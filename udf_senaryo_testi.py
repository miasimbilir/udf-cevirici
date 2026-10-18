#!/usr/bin/env python3
"""udf_senaryo_testi.py — docx_udf.py için senaryo testleri.

Her senaryo küçük bir DOCX'i sıfırdan üretir (dış kütüphane yok), dönüştürücüyü
çalıştırır ve çıkan UDF'yi denetler: XML geçerli mi, biçim dilimleri CDATA'yı
(UTF-16 birimiyle, Editör Java olduğu için) boşluksuz kaplıyor mu, tablo satırları
tutarlı mı, senaryoya özgü beklenti karşılandı mı.

Kullanım:
    python3 araclar/udf_senaryo_testi.py              # hepsi
    python3 araclar/udf_senaryo_testi.py -k gorsel    # ada göre süz
    python3 araclar/udf_senaryo_testi.py --tut KLASOR # DOCX/UDF'leri sakla (Word/Editör'de bakmak için)
"""
import argparse
import base64
import os
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import zipfile
import zlib
from xml.etree import ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
TOOL = os.path.join(HERE, 'docx_udf.py')
sys.path.insert(0, HERE)

NSDECL = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"')

STYLES = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NSDECL}>'
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" '
    'w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrDefault></w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
    '<w:pPr><w:jc w:val="both"/></w:pPr></w:style>'
    '<w:style w:type="table" w:styleId="TabloKlavuzu"><w:name w:val="Table Grid"/><w:tblPr>'
    '<w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>'
    '<w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>'
    '<w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/>'
    '</w:tblBorders></w:tblPr></w:style></w:styles>')

TEXT_W = (11906 - 2 * 1417) / 20.0          # metin genişliği, punto (453.6)


# Alt süreç çıktısı her sistemde UTF-8 (Windows'ta yerel kodlama cp1254 olunca emoji/Türkçe uyarı
# okunamıyordu — 18.10.2026 Windows denemesi)
UTF8 = {'text': True, 'encoding': 'utf-8', 'errors': 'replace', 'env': dict(os.environ, PYTHONIOENCODING='utf-8')}

# ---------- küçük yardımcılar ----------
def esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def R(text, rpr=''):
    return f'<w:r>{"<w:rPr>" + rpr + "</w:rPr>" if rpr else ""}<w:t xml:space="preserve">{esc(text)}</w:t></w:r>'


def P(inner, ppr=''):
    if not inner.lstrip().startswith('<'):
        inner = R(inner)
    return f'<w:p>{"<w:pPr>" + ppr + "</w:pPr>" if ppr else ""}{inner}</w:p>'


def TBL(rows, grid, tcpr=None):
    """rows: [[hücre_xml, ...]]; tcpr: {(satır, sütun): tcPr iç xml}"""
    tcpr = tcpr or {}
    out = ('<w:tbl><w:tblPr><w:tblStyle w:val="TabloKlavuzu"/><w:tblW w:w="0" w:type="auto"/>'
           '</w:tblPr><w:tblGrid>' + ''.join(f'<w:gridCol w:w="{g}"/>' for g in grid) + '</w:tblGrid>')
    for ri, row in enumerate(rows):
        out += '<w:tr>'
        for ci, cell in enumerate(row):
            out += f'<w:tc><w:tcPr>{tcpr.get((ri, ci), "")}</w:tcPr>{cell}</w:tc>'
        out += '</w:tr>'
    return out + '</w:tbl>'


def png(wd, ht, rgb=(31, 73, 125)):
    raw = bytearray()
    for y in range(ht):
        raw.append(0)
        for x in range(wd):
            edge = x < 3 or y < 3 or x >= wd - 3 or y >= ht - 3 or abs(x * ht - y * wd) < wd
            raw += bytes((255, 255, 255) if edge else rgb)

    def chunk(t, d):
        c = struct.pack('>I', len(d)) + t + d
        return c + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
    return (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', wd, ht, 8, 2, 0, 0, 0))
            + chunk(b'IDAT', zlib.compress(bytes(raw), 9)) + chunk(b'IEND', b''))


def jpeg_from_png(data):
    """Test JPEG'i: PIL ya da macOS sips; ikisi de yoksa None (senaryo atlanır)."""
    try:
        import io
        from PIL import Image
        buf = io.BytesIO()
        Image.open(io.BytesIO(data)).convert('RGB').save(buf, 'JPEG')
        return buf.getvalue()
    except Exception:
        pass
    if shutil.which('sips'):
        d = tempfile.mkdtemp()
        try:
            a, b = os.path.join(d, 'a.png'), os.path.join(d, 'b.jpg')
            open(a, 'wb').write(data)
            subprocess.run(['sips', '-s', 'format', 'jpeg', a, '--out', b],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            if os.path.exists(b):
                return open(b, 'rb').read()
        finally:
            shutil.rmtree(d, ignore_errors=True)
    return None


def inline(rid, w_pt, h_pt):
    cx, cy = int(w_pt * 12700), int(h_pt * 12700)
    return ('<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="1" name="Resim"/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name=""/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
            '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')


def anchor(rid, w_pt, h_pt):
    return (inline(rid, w_pt, h_pt)
            .replace('<wp:inline distT="0" distB="0" distL="0" distR="0">',
                     '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
                     'relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
                     '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column">'
                     '<wp:posOffset>3000000</wp:posOffset></wp:positionH>'
                     '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>')
            .replace('<wp:docPr', '<wp:wrapSquare wrapText="bothSides"/><wp:docPr')
            .replace('</wp:inline>', '</wp:anchor>'))


def make_docx(path, body, *, styles=True, numbering=None, media=None, footnotes=None,
              header=None, header_media=None, footer=None, sect=None, links=None):
    """media / header_media: {rId: (dosya_adı, bayt)}; links: {rId: url}"""
    media, header_media, links = media or {}, header_media or {}, links or {}
    rels = []
    if styles:
        rels.append(('rId1', 'styles', 'styles.xml', False))
    if numbering:
        rels.append(('rId2', 'numbering', 'numbering.xml', False))
    if footnotes:
        rels.append(('rId3', 'footnotes', 'footnotes.xml', False))
    if header:
        rels.append(('rId20', 'header', 'header1.xml', False))
    if footer:
        rels.append(('rId21', 'footer', 'footer1.xml', False))
    for rid, (name, _) in media.items():
        rels.append((rid, 'image', f'media/{name}', False))
    for rid, url in links.items():
        rels.append((rid, 'hyperlink', url, True))
    base = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'

    def rels_xml(items):
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships '
                'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                + ''.join(f'<Relationship Id="{i}" Type="{base}{t}" Target="{esc(tg)}"'
                          + (' TargetMode="External"' if ext else '') + '/>'
                          for i, t, tg, ext in items) + '</Relationships>')
    if sect is None:
        refs = ('<w:headerReference w:type="default" r:id="rId20"/>' if header else '') + \
               ('<w:footerReference w:type="default" r:id="rId21"/>' if footer else '')
        sect = (f'<w:sectPr>{refs}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" '
                'w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" '
                'w:gutter="0"/></w:sectPr>')
    doc = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NSDECL}>'
           f'<w:body>{body}{sect}</w:body></w:document>')
    over = [('document.xml', 'document.main')]
    if styles:
        over.append(('styles.xml', 'styles'))
    if numbering:
        over.append(('numbering.xml', 'numbering'))
    if footnotes:
        over.append(('footnotes.xml', 'footnotes'))
    if header:
        over.append(('header1.xml', 'header'))
    if footer:
        over.append(('footer1.xml', 'footer'))
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types '
          'xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Default Extension="png" ContentType="image/png"/>'
          '<Default Extension="jpeg" ContentType="image/jpeg"/>'
          '<Default Extension="emf" ContentType="image/x-emf"/>'
          + ''.join(f'<Override PartName="/word/{n}" ContentType="application/vnd.openxmlformats-'
                    f'officedocument.wordprocessingml.{t}+xml"/>' for n, t in over) + '</Types>')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', ct)
        z.writestr('_rels/.rels', rels_xml([('rId1', 'officeDocument', 'word/document.xml', False)]))
        z.writestr('word/document.xml', doc)
        z.writestr('word/_rels/document.xml.rels', rels_xml(rels))
        if styles:
            z.writestr('word/styles.xml', STYLES)
        if numbering:
            z.writestr('word/numbering.xml', numbering)
        if footnotes:
            z.writestr('word/footnotes.xml', footnotes)
        for part, xml, tag in (('header1', header, 'hdr'), ('footer1', footer, 'ftr')):
            if xml:
                z.writestr(f'word/{part}.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                           f'<w:{tag} {NSDECL}>{xml}</w:{tag}>')
        if header and header_media:
            z.writestr('word/_rels/header1.xml.rels',
                       rels_xml([(rid, 'image', f'media/{n}', False) for rid, (n, _) in header_media.items()]))
        for _, (name, data) in list(media.items()) + list(header_media.items()):
            z.writestr(f'word/media/{name}', data)


# ---------- UDF okuyucu ----------
class U:
    def __init__(self, path, stdout):
        x = zipfile.ZipFile(path).read('content.xml').decode('utf-8')
        self.xml = x
        self.root = ET.fromstring(x.split('?>', 1)[1].strip())
        self.cdata = self.root.find('content').text or ''
        self.cd16 = self.cdata.encode('utf-16-le')
        self.els = self.root.find('elements')
        self.warn = [l for l in stdout.splitlines() if l.startswith('UYARI')]
        self.page = self.root.find('properties/pageFormat')

    def sl(self, c):
        so, ln = int(c.get('startOffset')), int(c.get('length'))
        return self.cd16[2 * so:2 * (so + ln)].decode('utf-16-le', 'replace')

    def ptext(self, p):
        s = ''.join(self.sl(c) for c in p if c.get('startOffset') is not None)
        return s[:-1] if s.endswith('\n') else s

    def body(self):
        for e in self.els:
            if e.tag == 'paragraph':
                yield e
            elif e.tag == 'table':
                yield from e.iter('paragraph')
            # <page-break> içindeki taşıyıcı paragraf metin sayılmaz

    def texts(self):
        return [self.ptext(p) for p in self.body()]

    def alltext(self):
        return '\n'.join(self.texts())

    def contents(self, **want):
        """Gövdede verilen özniteliklere sahip content dilimlerinin metinleri."""
        out = []
        for p in self.body():
            for c in p:
                if c.tag == 'content' and all(c.get(k) == v for k, v in want.items()):
                    out.append(self.sl(c))
        return out

    def images(self, where=None):
        out = []
        for e in self.els:
            w_ = 'header' if e.tag == 'header' else 'footer' if e.tag == 'footer' else \
                'cell' if e.tag == 'table' else 'body'
            for im in e.iter('image'):
                if where in (None, w_):
                    out.append(im)
        return out

    def generic(self):
        errs = []
        pos = 0
        for p in self.els.iter('paragraph'):
            sl = [c for c in p if c.get('startOffset') is not None]
            for c in sl:
                so, ln = int(c.get('startOffset')), int(c.get('length'))
                if so != pos:
                    errs.append(f'dilim boşluğu: beklenen {pos}, gelen {so}')
                    pos = so
                pos += ln
            if sl and not self.sl(sl[-1]).endswith('\n'):
                errs.append('paragraf \\n ile bitmiyor')
        if pos != len(self.cd16) // 2:
            errs.append(f'CDATA kaplanmadı: {pos} / {len(self.cd16) // 2} (UTF-16 birimi)')
        for t in self.els.iter('table'):
            n = int(t.get('columnCount'))
            for r in t.findall('row'):
                want = len(r.get('columnSpans').split(',')) if r.get('columnSpans') else n
                if len(r.findall('cell')) != want:
                    errs.append('tablo satırında hücre sayısı sütun tanımıyla uyuşmuyor')
        return errs[:3]


def img_bytes(im):
    import base64
    return base64.b64decode(re.sub(r'\s', '', im.get('imageData')))


# ---------- senaryolar ----------
SCEN = []
# metin kaydırmalı görsel taklidi deneysel ve varsayılan kapalı: ilgili senaryolar bayrakla koşar
DOGRULAMA_HARIC = set()      # doğrulayıcının bilerek KALDI dediği senaryolar (gerekçesiyle eklenir)
SCEN_ARGS = {'gorsel_anchor': ['--yan-gorsel'], 'gorsel_yan_metin': ['--yan-gorsel']}
RENDER = ['gorsel_yan_metin', 'gorsel_onde', 'gorsel_ust_alt', 'dikey_birlesik_karmasik', 'hucre_dolgu_hizalama', 'sekme_sag_kenar', 'kenarliksiz_tablo', 'koyu_zemin_beyaz_yazi', 'altbilgi_sayfa_kalibi', 'gorsel_png_govde', 'gorsel_jpeg', 'gorsel_buyuk', 'gorsel_tablo_hucresi', 'gorsel_vml',
          'birlesik_hucre', 'dipnot', 'ust_alt_simge_vurgu', 'sayfa_sonu', 'cok_duzeyli_liste',
          'sekme_lider', 'yatay_sayfa', 'ilk_sayfa_farkli_ustbilgi', 'ustbilgi_tablolu_logo',
          'liste_alt_duzey_gorunumu', 'liste_bicim_ve_baslangic', 'altbilgi_sayfa_no_hizasi',
          'antet_tam_sayfa', 'antet_konumlu_logo', 'antet_vml_filigran', 'gorsel_pdf_wmf_adli',
          'metin_kutusu', 'pandoc_belgesi']   # --onizle ile Editör motorunda çizilecekler


def senaryo(name, note):
    def deco(fn):
        SCEN.append((name, note, fn))
        return fn
    return deco


@senaryo('gorsel_png_govde', 'Gövdede satır içi PNG')
def _(d):
    data = png(300, 120)
    make_docx(d, P('Görselden önce.') + P(inline('rId10', 150, 60), '<w:jc w:val="center"/>')
              + P('Görselden sonra.'), media={'rId10': ('image1.png', data)})

    def exp(u):
        ims = u.images('body')
        if len(ims) != 1:
            return [f'gövdede 1 görsel beklenirdi, {len(ims)} var']
        e = []
        if img_bytes(ims[0]) != data:
            e.append('PNG verisi değişmiş')
        if abs(float(ims[0].get('width')) - 150) > 0.6 or abs(float(ims[0].get('height')) - 60) > 0.6:
            e.append(f'boyut {ims[0].get("width")}x{ims[0].get("height")} ≠ 150x60')
        if u.texts() != ['Görselden önce.', '¸', 'Görselden sonra.']:
            e.append(f'metin akışı: {u.texts()}')
        return e
    return exp


@senaryo('gorsel_jpeg', 'Gövdede JPEG (UDF yalnız PNG taşır → çevrilmeli)')
def _(d):
    jp = jpeg_from_png(png(200, 100, (160, 30, 30)))
    if jp is None:
        return None
    make_docx(d, P('JPEG:') + P(inline('rId10', 100, 50)), media={'rId10': ('image1.jpeg', jp)})

    def exp(u):
        ims = u.images('body')
        if len(ims) != 1:
            return [f'JPEG taşınmadı ({len(ims)} görsel); uyarılar: {u.warn}']
        return [] if img_bytes(ims[0])[:4] == b'\x89PNG' else ['imageData PNG değil']
    return exp


@senaryo('gorsel_buyuk', 'Metin alanından geniş görsel (orantılı küçülmeli)')
def _(d):
    make_docx(d, P(inline('rId10', 700, 350)), media={'rId10': ('image1.png', png(400, 200))})

    def exp(u):
        im = u.images('body')
        if not im:
            return ['görsel yok']
        w_, h_ = float(im[0].get('width')), float(im[0].get('height'))
        e = []
        if w_ > TEXT_W + 0.5:
            e.append(f'genişlik {w_} > metin alanı {TEXT_W}')
        if abs(w_ / h_ - 2.0) > 0.02:
            e.append(f'en-boy oranı bozuldu: {w_}x{h_}')
        return e
    return exp


@senaryo('gorsel_tablo_hucresi', 'Tablo hücresinde görsel (hücreye sığmalı)')
def _(d):
    make_docx(d, TBL([[P('İmza'), P(inline('rId10', 300, 100))]], [6000, 3000])
              + P('son'), media={'rId10': ('image1.png', png(300, 100))})

    def exp(u):
        im = u.images('cell')
        if len(im) != 1:
            return [f'hücrede 1 görsel beklenirdi, {len(im)}']
        cellw = TEXT_W * 3000 / 9000
        return [] if float(im[0].get('width')) <= cellw + 0.5 else \
            [f'görsel {im[0].get("width")} > hücre genişliği {cellw:.0f}']
    return exp


def anchor2(rid, w_pt, h_pt, wrap='wrapSquare', align='left'):
    """Hizası align ile verilen serbest görsel; wrap: wrapSquare | wrapTopAndBottom | wrapNone"""
    wx = {'wrapSquare': '<wp:wrapSquare wrapText="bothSides"/>', 'wrapTopAndBottom': '<wp:wrapTopAndBottom/>',
          'wrapNone': '<wp:wrapNone/>'}[wrap]
    return (inline(rid, w_pt, h_pt)
            .replace('<wp:inline distT="0" distB="0" distL="0" distR="0">',
                     '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
                     'relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
                     '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column">'
                     f'<wp:align>{align}</wp:align></wp:positionH>'
                     '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>')
            .replace('<wp:docPr', wx + '<wp:docPr')
            .replace('</wp:inline>', '</wp:anchor>'))


UZUN = ('Bu paragraf, görselin yanına sığan satır sayısını aşacak kadar uzun tutulmuştur; dönüştürücü metnin '
        'ne kadarının görselin yanında kalacağını satır satır tahmin eder ve taşan kısmı görselin altına, '
        'sayfa genişliğine yayarak devam ettirir. ') * 6


@senaryo('gorsel_anchor', 'Sağa yaslı, metin kaydırmalı serbest görsel → kenarlıksız iki sütunlu tablo')
def _(d):
    make_docx(d, P(R('Metin ') + anchor('rId10', 120, 60) + R('devam.')) + P('Sonraki paragraf.'),
              media={'rId10': ('image1.png', png(240, 120))})

    def exp(u):
        t = u.els.find('table')
        if t is None:
            return [f'yan yana tablo kurulmamış; uyarılar: {u.warn}']
        cells = t.find('row').findall('cell')
        e = []
        if t.get('border') != 'borderNone' or len(cells) != 2:
            e.append(f'tablo: border={t.get("border")} hücre={len(cells)}')
        elif cells[1].find('.//image') is None:
            e.append('sağa yaslı görsel sağ hücrede değil')
        if 'Metin devam.' not in u.alltext():
            e.append(f'metin: {u.texts()}')
        if not any('taklit' in x for x in u.warn):
            e.append('taklit için UYARI yok')
        return e
    return exp


@senaryo('gorsel_yan_metin', 'Sola yaslı görselin yanında uzun metin: satır sınırından bölünüp altta sürmeli')
def _(d):
    make_docx(d, P('BAŞLIK') + P(anchor2('rId10', 150, 200) + R(UZUN)) + P('Son paragraf.'),
              media={'rId10': ('image1.png', png(150, 200))})

    def exp(u):
        kinds = [x.tag for x in u.els if x.tag in ('paragraph', 'table')]
        t = u.els.find('table')
        e = []
        if t is None or kinds[:3] != ['paragraph', 'table', 'paragraph']:
            return [f'yapı: {kinds}']
        cells = t.find('row').findall('cell')
        if cells[0].find('.//image') is None:
            e.append('sola yaslı görsel sol hücrede değil')
        yan = ' '.join(u.ptext(p) for p in cells[1].findall('paragraph'))
        alt = u.ptext([x for x in u.els if x.tag == 'paragraph'][1])
        if (yan + ' ' + alt).split() != UZUN.split():
            e.append('bölünen paragrafın iki parçası birleşince özgün metni vermiyor')
        if not (0.25 < len(yan) / max(1, len(UZUN)) < 0.9):
            e.append(f'bölme oranı tuhaf: yan={len(yan)} toplam={len(UZUN)}')
        return e
    return exp


@senaryo('gorsel_kaydirmali_varsayilan', 'Metin kaydırmalı görsel, bayrak YOKKEN: satır içi + uyarı (taklit kapalı)')
def _(d):
    make_docx(d, P(anchor2('rId10', 150, 200) + R('Yanındaki metin.')) + P('Son.'),
              media={'rId10': ('image1.png', png(150, 200))})

    def exp(u):
        e = []
        if u.els.find('table') is not None:
            e.append('bayrak verilmediği hâlde iki sütunlu tablo kurulmuş')
        if len(u.images('body')) != 1:
            e.append('görsel satır içine alınmamış')
        if not any('serbest' in x for x in u.warn):
            e.append('UYARI yok')
        return e
    return exp


@senaryo('gorsel_onde', 'Metnin önünde/arkasında duran serbest görsel (kaydırma yok) → satır içi + uyarı')
def _(d):
    make_docx(d, P(R('Metin ') + anchor2('rId10', 80, 40, 'wrapNone') + R('devam.')),
              media={'rId10': ('image1.png', png(160, 80))})

    def exp(u):
        e = []
        if len(u.images('body')) != 1 or u.els.find('table') is not None:
            e.append('görsel satır içinde değil')
        if not any('serbest' in x for x in u.warn):
            e.append('konum kaybı için UYARI yok')
        return e
    return exp


@senaryo('gorsel_ust_alt', 'Üst-alt kaydırmalı serbest görsel → kendi satırında, ortalı')
def _(d):
    make_docx(d, P(anchor2('rId10', 200, 80, 'wrapTopAndBottom', 'center') + R('Görselin altındaki metin.')),
              media={'rId10': ('image1.png', png(200, 80))})

    def exp(u):
        ps = list(u.body())
        ok = len(ps) == 2 and ps[0].find('image') is not None and ps[0].get('Alignment') == '1' \
            and u.ptext(ps[1]) == 'Görselin altındaki metin.'
        return [] if ok else [f'yapı: {[(p.get("Alignment"), u.ptext(p)) for p in ps]}']
    return exp


@senaryo('gorsel_vml', 'Eski tip VML görsel (w:pict)')
def _(d):
    pict = ('<w:r><w:pict><v:shape id="s1" type="#_x0000_t75" style="width:90pt;height:45pt">'
            '<v:imagedata r:id="rId10" o:title=""/></v:shape></w:pict></w:r>')
    make_docx(d, P(R('Mühür: ') + pict), media={'rId10': ('image1.png', png(180, 90))})

    def exp(u):
        im = u.images('body')
        if len(im) != 1:
            return [f'VML görsel taşınmadı; uyarılar: {u.warn}']
        return [] if abs(float(im[0].get('width')) - 90) < 0.6 else [f'genişlik {im[0].get("width")} ≠ 90']
    return exp


@senaryo('gorsel_emf', 'Desteklenmeyen biçim (EMF): atla + uyar, belge sağlam kalsın')
def _(d):
    make_docx(d, P('Önce') + P(inline('rId10', 100, 50)) + P('Sonra'),
              media={'rId10': ('image1.emf', b'\x01\x00\x00\x00' + b'\x00' * 120)})

    def exp(u):
        e = []
        if u.images():
            e.append('EMF görsel olarak gömülmüş')
        if not u.warn:
            e.append('UYARI yok')
        if u.texts()[0] != 'Önce' or u.texts()[-1] != 'Sonra':
            e.append(f'metin: {u.texts()}')
        return e
    return exp


@senaryo('ustbilgi_tablolu_logo', 'Üstbilgide tablo içinde logo + metin')
def _(d):
    hdr = TBL([[P(inline('rId1', 80, 40)), P('BİLİR AVUKATLIK')]], [3000, 6000])
    make_docx(d, P('Gövde'), header=hdr, header_media={'rId1': ('logo.png', png(160, 80))})

    def exp(u):
        h = u.els.find('header')
        if h is None:
            return ['üstbilgi yok']
        e = []
        if len(list(h.iter('image'))) != 1:
            e.append('logo yok')
        if 'BİLİR AVUKATLIK' not in ''.join(u.ptext(p) for p in h.iter('paragraph')):
            e.append('üstbilgi metni yok')
        return e
    return exp


@senaryo('metin_kutusu', 'Metin kutusu (AlternateContent: metin iki kez gelmemeli)')
def _(d):
    inner = '<w:txbxContent><w:p><w:r><w:t>KUTU METNİ</w:t></w:r></w:p></w:txbxContent>'
    box = ('<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" '
           'distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2" behindDoc="0" locked="0" '
           'layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH '
           'relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV '
           'relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent '
           'cx="1800000" cy="600000"/><wp:wrapNone/><wp:docPr id="2" name="Kutu"/><a:graphic>'
           '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
           f'<wps:wsp><wps:spPr/><wps:txbx>{inner}</wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData>'
           '</a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape id="k1" '
           f'style="width:140pt;height:47pt"><v:textbox>{inner}</v:textbox></v:shape></w:pict>'
           '</mc:Fallback></mc:AlternateContent></w:r>')
    make_docx(d, P(R('Önce ') + box + R('sonra')) + P('Bitiş'))

    def exp(u):
        t = u.alltext()
        e = []
        if t.count('KUTU METNİ') != 1:
            e.append(f'kutu metni {t.count("KUTU METNİ")} kez geçiyor (1 olmalı)')
        if 'Önce sonra' not in t:
            e.append(f'taşıyan paragraf bozuldu: {u.texts()}')
        return e
    return exp


@senaryo('birlesik_hucre', 'Birleştirilmiş hücreler: yatay → satır columnSpans, dikey → iç içe tablo')
def _(d):
    rows = [[P('Başlık (2 sütun)'), P('C')], [P('A1'), P('B1'), P('dikey')], [P('A2'), P('B2'), P('')]]
    tcpr = {(0, 0): '<w:gridSpan w:val="2"/>', (1, 2): '<w:vMerge w:val="restart"/>', (2, 2): '<w:vMerge/>'}
    make_docx(d, TBL(rows, [3000, 3000, 3000], tcpr) + P('son'))

    def exp(u):
        t = u.els.find('table')
        rows = t.findall('row')
        e = []
        if len(rows) != 2:
            return [f'dış tabloda {len(rows)} satır var (2 beklenir: başlık + dikey birleşik grup)']
        if rows[0].get('columnSpans') != '200,100':
            e.append(f'yatay birleştirme: columnSpans={rows[0].get("columnSpans")!r} ≠ 200,100')
        cells = rows[1].findall('cell')
        inner = cells[0].find('table') if cells else None
        if len(cells) != 2 or inner is None or rows[1].get('columnSpans') != '200,100':
            e.append('dikey grup: [iç tablo | birleşik hücre] yapısı kurulmamış')
        elif [len(r.findall('cell')) for r in inner.findall('row')] != [2, 2]:
            e.append('iç tablo 2x2 değil')
        if 'dikey' not in u.ptext(cells[-1].find('paragraph')):
            e.append('birleşik hücre metni yerinde değil')
        return e
    return exp


@senaryo('dikey_birlesik_karmasik', 'Örtüşen/karmaşık dikey birleştirme → güvenli geri dönüş (boş hücre + uyarı)')
def _(d):
    rows = [[P('A1'), P('B1')], [P(''), P('B2')], [P('A3'), P('')]]
    tcpr = {(0, 0): '<w:vMerge w:val="restart"/>', (1, 0): '<w:vMerge/>',
            (1, 1): '<w:vMerge w:val="restart"/>', (2, 1): '<w:vMerge/>'}
    make_docx(d, TBL(rows, [4500, 4500], tcpr) + P('son'))

    def exp(u):
        e = []
        if not any('dikey' in x for x in u.warn):
            e.append('karmaşık dikey birleştirme için UYARI yok')
        if not all(k in u.alltext() for k in ('A1', 'B1', 'B2', 'A3')):
            e.append(f'metin kaybı: {u.texts()}')
        return e
    return exp


@senaryo('hucre_dolgu_hizalama', 'Hücre dolgusu (bant), iç boşluk ve dikey ortalama taklidi')
def _(d):
    shd = '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/><w:vAlign w:val="center"/>'
    mar = '<w:tcMar><w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/>'
    rows = [[P(R('Sıra', '<w:b/>'), '<w:jc w:val="center"/>'), P(R('Açıklama', '<w:b/>'), '<w:jc w:val="center"/>')],
            [P('1'), P('Bu açıklama hücresi iki satıra sarılacak kadar uzundur; yanındaki tek satırlık hücre '
                       'dikeyde ortalanmalıdır ki Word görünümüne yaklaşsın.')]]
    make_docx(d, TBL(rows, [1500, 7500], {(0, 0): shd, (0, 1): shd, (1, 0): mar, (1, 1): mar}) + P('son'))

    def exp(u):
        t = u.els.find('table')
        ps = [r.findall('cell')[0].find('paragraph') for r in t.findall('row')]
        e = []
        first = [c for c in ps[0] if c.tag == 'content' and c.get('background')]
        if len(first) < 2 or '\u00a0' not in ''.join(u.sl(c) for c in first):
            e.append('başlık hücresinde dolgu bandı (zeminli bölünmez boşluk) yok')
        if float(ps[1].get('SpaceAbove') or 0) < 8:
            e.append(f'tek satırlık hücre ortalanmamış: SpaceAbove={ps[1].get("SpaceAbove")}')
        return e
    return exp


@senaryo('ic_ice_tablo', 'Hücre içinde tablo')
def _(d):
    innert = TBL([[P('iç-1'), P('iç-2')]], [2000, 2000])
    make_docx(d, TBL([[P('dış') + innert + P('dış-son'), P('yan')]], [5000, 4000]) + P('son'))

    def exp(u):
        t = u.alltext()
        return [] if all(k in t for k in ('iç-1', 'iç-2', 'dış-son', 'yan')) else [f'metin kaybı: {u.texts()}']
    return exp


@senaryo('son_blok_tablo', 'Belge tabloyla bitiyor (ardından paragraf gerekir)')
def _(d):
    make_docx(d, P('Giriş') + TBL([[P('a'), P('b')]], [4500, 4500]))

    def exp(u):
        kinds = [e.tag for e in u.els if e.tag in ('paragraph', 'table')]
        return [] if kinds and kinds[-1] == 'paragraph' else ['son gövde öğesi tablo; ardından paragraf yok']
    return exp


FOOTNOTES = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes {NSDECL}>'
             '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
             '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/>'
             '</w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:rPr><w:vertAlign '
             'w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> '
             'Yargıtay 9. HD, 2020/123 E., 2021/456 K.</w:t></w:r></w:p></w:footnote></w:footnotes>')


@senaryo('dipnot', 'Dipnot (işaret üst simge, metin belge sonunda)')
def _(d):
    ref = '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>'
    make_docx(d, P(R('Karar bu yöndedir') + ref + R('.')) + P('İkinci paragraf.'), footnotes=FOOTNOTES)

    def exp(u):
        e = []
        if 'Yargıtay 9. HD, 2020/123 E.' not in u.alltext():
            e.append('dipnot metni kayboldu')
        if '1' not in [s.strip('()') for s in u.contents(superscript='true')]:
            e.append(f'üst simge dipnot işareti yok: {u.contents(superscript="true")}')
        return e
    return exp


@senaryo('ust_alt_simge_vurgu', 'Üst/alt simge, üstü çizili, vurgu (highlight)')
def _(d):
    make_docx(d, P(R('m') + R('2', '<w:vertAlign w:val="superscript"/>') + R(' H')
                   + R('2', '<w:vertAlign w:val="subscript"/>') + R('O ')
                   + R('iptal', '<w:strike/>') + R(' ') + R('sarı', '<w:highlight w:val="yellow"/>')))

    def exp(u):
        e = []
        for k, v in (('superscript', '2'), ('strikethrough', 'iptal')):
            if v not in u.contents(**{k: 'true'}):
                e.append(f'{k} taşınmadı')
        # Editör motoru subscript özniteliğini yukarıda çiziyor → Unicode alt simge beklenir
        if 'H\u2082O' not in u.texts()[0]:
            e.append(f'alt simge Unicode karaktere çevrilmedi: {u.texts()[0]!r}')
        if not any(c.get('background') for p in u.body() for c in p if u.sl(c) == 'sarı'):
            e.append('vurgu rengi (background) taşınmadı')
        return e
    return exp


@senaryo('buyuk_harf_caps', 'Word "Tümü büyük harf" biçimi (Türkçe i/ı)')
def _(d):
    make_docx(d, P(R('istinaf dilekçesi ığ', '<w:caps/>')))
    return lambda u: [] if u.texts()[0] == 'İSTİNAF DİLEKÇESİ IĞ' else [f'metin: {u.texts()[0]!r}']


@senaryo('izlenen_degisiklik', 'Değişiklik izleme: ekleme kalır, silme ve taşıma-kaynağı gider')
def _(d):
    make_docx(d, P(R('Sabit ') + '<w:ins w:id="1" w:author="a">' + R('eklenen ') + '</w:ins>'
                   '<w:del w:id="2" w:author="a"><w:r><w:delText>silinen </w:delText></w:r></w:del>'
                   '<w:moveFrom w:id="3" w:author="a">' + R('taşınan-eski ') + '</w:moveFrom>'
                   '<w:moveTo w:id="4" w:author="a">' + R('taşınan-yeni') + '</w:moveTo>'))
    return lambda u: [] if u.texts()[0] == 'Sabit eklenen taşınan-yeni' else [f'metin: {u.texts()[0]!r}']


@senaryo('kopru', 'Köprü (w:hyperlink ve HYPERLINK alanı)')
def _(d):
    fld = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> '
           'HYPERLINK "http://ornek.org" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/>'
           '</w:r>' + R('alan metni') + '<w:r><w:fldChar w:fldCharType="end"/></w:r>')
    make_docx(d, P('<w:hyperlink r:id="rId30">' + R('bağlantı metni') + '</w:hyperlink>' + R(' ve ') + fld),
              links={'rId30': 'http://ornek.org'})
    return lambda u: [] if u.texts()[0] == 'bağlantı metni ve alan metni' else [f'metin: {u.texts()[0]!r}']


@senaryo('icindekiler_pageref', 'İçindekiler tablosu: PAGEREF sayfa numaraları metindir, ATILMAMALI')
def _(d):
    fc = lambda t: f'<w:r><w:fldChar w:fldCharType="{t}"/></w:r>'
    ins = lambda t: f'<w:r><w:instrText xml:space="preserve"> {t} </w:instrText></w:r>'
    satir = lambda ad, no, bas='': P(bas + R(ad) + '<w:r><w:tab/></w:r>' + fc('begin') + ins(f'PAGEREF _Toc{no} \\h')
                                     + fc('separate') + R(str(no)) + fc('end'),
                                     '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs>')
    toc = (satir('Giriş', 12, fc('begin') + ins('TOC \\o "1-3" \\h') + fc('separate'))
           + satir('Birinci Bölüm', 44) + P(R('Sonuç') + fc('end')))
    make_docx(d, P('İÇİNDEKİLER') + toc + P('Gövde metni alanlardan sonra da görünmeli.'),
              footer=P(R('Sayfa ') + fc('begin') + ins('PAGE') + fc('separate') + R('7') + fc('end')))

    def exp(u):
        t = u.texts()
        e = []
        if 'Giriş\t12' not in t or 'Birinci Bölüm\t44' not in t:
            e.append(f'içindekiler sayfa numaraları kayıp: {t[:4]}')
        if 'Gövde metni alanlardan sonra da görünmeli.' not in t or 'Sonuç' not in t:
            e.append('iç içe alandan sonraki metin gizlenmiş')
        f = u.els.find('footer')
        if any('7' in u.ptext(p) for p in f.iter('paragraph')):
            e.append('altbilgideki PAGE alanının önbellek değeri metne sızmış')
        if not any('içindekiler' in x for x in u.warn):
            e.append('içindekiler için UYARI yok')
        return e
    return exp


@senaryo('sembol', 'Simge (w:sym, Wingdings onay işareti)')
def _(d):
    make_docx(d, P(R('Kabul ') + '<w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r>' + R(' edildi')))
    return lambda u: [] if u.texts()[0] == 'Kabul ✓ edildi' else [f'metin: {u.texts()[0]!r}']


@senaryo('yumusak_satir_sonu', 'Paragraf içi satır sonu (Shift+Enter)')
def _(d):
    make_docx(d, P(R('İSTANBUL', '<w:b/>') + '<w:r><w:br/></w:r>' + R('19. İDARE MAHKEMESİNE', '<w:b/>'),
                   '<w:jc w:val="center"/>'))
    return lambda u: [] if u.texts() == ['İSTANBUL', '19. İDARE MAHKEMESİNE'] else [f'metin: {u.texts()}']


@senaryo('sayfa_sonu', 'Sayfa sonu (Ctrl+Enter ve "önce sayfa sonu") → <page-break>')
def _(d):
    make_docx(d, P('Birinci sayfa') + P('<w:r><w:br w:type="page"/></w:r>') + P('EKLER', '<w:pageBreakBefore/>'))

    def exp(u):
        e = []
        if [t for t in u.texts() if t] != ['Birinci sayfa', 'EKLER']:
            e.append(f'metin: {u.texts()}')
        n = len(u.els.findall('page-break'))
        if n != 2:
            e.append(f'{n} page-break öğesi var (2 beklenir)')
        return e
    return exp


NUMBERING = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering {NSDECL}>'
             '<w:abstractNum w:abstractNumId="0">'
             '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>'
             '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>'
             '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/>'
             '<w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>'
             '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%3."/>'
             '<w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>'
             '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/>'
             '<w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" '
             'w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>'
             '<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/>'
             '<w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0D8;"/><w:pPr><w:ind w:left="720" '
             'w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr>'
             '</w:lvl></w:abstractNum>'
             '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
             '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
             '<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num></w:numbering>')


def LI(text, numid, ilvl=0):
    return P(text, f'<w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="{numid}"/></w:numPr>')


@senaryo('cok_duzeyli_liste', 'Çok düzeyli liste, büyük Roma rakamı, ok işaretli madde')
def _(d):
    make_docx(d, LI('bir', 1) + LI('bir-a', 1, 1) + LI('bir-a-i', 1, 2) + LI('iki', 1)
              + LI('USUL', 2) + LI('ESAS', 2) + LI('ok maddesi', 3), numbering=NUMBERING)

    def exp(u):
        ps = list(u.body())
        e = []
        want = [('NUMBER_TYPE_NUMBER_DOT', '1'), ('NUMBER_TYPE_CHAR_SMALL_PARANTHESE', '2'),   # Word'de "a)"
                ('NUMBER_TYPE_ROMAN_SMALL_DOT', '3'), ('NUMBER_TYPE_NUMBER_DOT', '1'),
                ('NUMBER_TYPE_ROMAN_BIG_DOT', '1'), ('NUMBER_TYPE_ROMAN_BIG_DOT', '1')]
        for p, (nt, lv) in zip(ps, want):
            if p.get('NumberType') != nt or p.get('ListLevel') != lv:
                e.append(f'{u.ptext(p)!r}: {p.get("NumberType")}/{p.get("ListLevel")} ≠ {nt}/{lv}')
        if u.texts()[4] != 'USUL':
            e.append(f'Roma rakamı düz metne düşmüş: {u.texts()[4]!r}')
        if ps[6].get('BulletType') != 'BULLET_TYPE_ARROW':
            e.append(f'ok işareti: {ps[6].get("BulletType")}')
        return e[:4]
    return exp


NUMBERING2 = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering {NSDECL}>'
              '<w:abstractNum w:abstractNumId="0">'
              '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1)"/></w:lvl>'
              '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2."/></w:lvl>'
              '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0A7;"/>'
              '<w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr></w:lvl></w:abstractNum>'
              '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/>'
              '<w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%1)"/></w:lvl></w:abstractNum>'
              '<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="4"/>'
              '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1-"/></w:lvl></w:abstractNum>'
              '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
              '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
              '<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>'
              '<w:num w:numId="4"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0">'
              '<w:startOverride w:val="3"/></w:lvlOverride></w:num></w:numbering>')


@senaryo('liste_alt_duzey_gorunumu', 'Alt düzey görünümü: Editör NumberType\'a bakmaz, SecListTypeLevelN ister')
def _(d):
    make_docx(d, LI('ana', 1) + LI('alt rakam', 1, 1) + LI('alt işaret', 1, 2) + LI('ana iki', 1),
              numbering=NUMBERING2)

    def exp(u):
        ps = list(u.body())
        e = []
        if ps[0].get('NumberType') != 'NUMBER_TYPE_CHAR_BIG_PARANTHESE' or ps[0].get('SecListTypeLevel1'):
            e.append(f'1. düzey: {ps[0].attrib}')
        if ps[1].get('SecListTypeLevel2') != 'NUMBER_TYPE_NUMBER_DOT':
            e.append(f'2. düzeyde SecListTypeLevel2 yok/yanlış: {ps[1].get("SecListTypeLevel2")} '
                     '(Editör "1." yerine "a." çizer)')
        if ps[2].get('SecListTypeLevel3') != 'BULLET_TYPE_RECTANGLE':
            e.append(f'3. düzey işaret: {ps[2].get("SecListTypeLevel3")}')
        return e
    return exp


@senaryo('liste_bicim_ve_baslangic', 'Numara ekleri "(a)" "4-" ve 1\'den başlamayan liste -> NumberSetted')
def _(d):
    make_docx(d, LI('parantezli', 2) + LI('dördüncü', 3) + LI('beşinci', 3) + LI('C ile başlar', 4) + LI('D', 4),
              numbering=NUMBERING2)

    def exp(u):
        ps = list(u.body())
        e = []
        if ps[0].get('NumberType') != 'NUMBER_TYPE_CHAR_SMALL_D_PARANTHESE':
            e.append(f'(a): {ps[0].get("NumberType")}')
        if ps[1].get('NumberType') != 'NUMBER_TYPE_NUMBER_TRE' or ps[1].get('NumberSetted') != '3':
            e.append(f'4-: {ps[1].get("NumberType")} NumberSetted={ps[1].get("NumberSetted")}')
        if ps[2].get('NumberSetted'):
            e.append('NumberSetted yalnız ilk kalemde olmalı')
        if ps[3].get('NumberSetted') != '2' or ps[4].get('NumberSetted'):
            e.append(f'startOverride=3: NumberSetted={ps[3].get("NumberSetted")}/{ps[4].get("NumberSetted")}')
        return e
    return exp


@senaryo('liste_harf_alfabe_uyarisi', 'Harfli listede 4. kalem: Editör Türk alfabesiyle sayar (ç) → uyarı verilmeli')
def _(d):
    make_docx(d, ''.join(LI(f'bent {i + 1}', 2) for i in range(5)) + LI('üç kalemlik', 1), numbering=NUMBERING2)

    def exp(u):
        return [] if any('Türk alfabesi' in x for x in u.warn) else ['alfabe kayması uyarısı yok']
    return exp


@senaryo('altbilgi_sayfa_no_hizasi', 'Sayfa numarası Word\'deki hizada ve biçimde (ortalı, yalnız sayfa no)')
def _(d):
    fld = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>'
           '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' + R('1') + '<w:r><w:fldChar w:fldCharType="end"/></w:r>')
    make_docx(d, P('Gövde'), footer=P(fld, '<w:jc w:val="center"/>'))

    def exp(u):
        f = u.els.find('footer')
        spec = f.get('pageNumber-spec') if f is not None else None
        return [] if spec == 'BSP32_40' else [f'pageNumber-spec {spec} (beklenen BSP32_40: ortalı, toplam yok)']
    return exp


@senaryo('sekme_lider', 'Sağa hizalı, nokta liderli sekme durağı')
def _(d):
    make_docx(d, P(R('Ek-1') + '<w:r><w:tab/></w:r>' + R('s. 12'),
                   '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs>'))

    def exp(u):
        ts = next(u.body()).get('TabSet') or ''
        first = ts.split(',')[0].split(':') if ts else ['0', '', '']
        ok = 449.0 <= float(first[0]) <= 450.1 and first[1:] == ['1', '1']      # kenara yakınsa 4 pt içeri alınır
        return [] if ok else [f'TabSet: {ts[:40]!r}']
    return exp


@senaryo('sekme_sag_kenar', 'Sağ kenara dayalı sağ sekme (Editör sığdıramayıp satırı kırıyordu)')
def _(d):
    make_docx(d, P(R('Ek-1') + '<w:r><w:tab/></w:r>' + R('4 sayfa'),
                   '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9072"/></w:tabs>'))

    def exp(u):
        ts = next(u.body()).get('TabSet') or ''
        pos = float(ts.split(':')[0]) if ts else 0
        return [] if 0 < pos <= TEXT_W - 3.9 else [f'sağ sekme kenarda kaldı: {ts[:30]!r} (metin alanı {TEXT_W})']
    return exp


@senaryo('sekme_girintili', 'Girintili paragrafta sekme: Editör konumu girintiden ölçer, Word kenardan')
def _(d):
    make_docx(d, P(R('Alt başlık') + '<w:r><w:tab/></w:r>' + R('44'),
                   '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs><w:ind w:left="720"/>')
              + P(R('Etiket') + '<w:r><w:tab/></w:r>' + R('değer'), '<w:ind w:left="400"/>'))

    def exp(u):
        ps = list(u.body())
        e = []
        t1 = (ps[0].get('TabSet') or '0').split(':')[0]
        if abs(float(t1) - (TEXT_W - 4.0 - 36.0)) > 0.6:             # 449,6 (kenar payı) − 36 (LeftIndent)
            e.append(f'sağ durak girintiden arındırılmamış: {t1} (beklenen {TEXT_W - 40.0:.1f})')
        t2 = (ps[1].get('TabSet') or '0').split(':')[0]
        if abs(float(t2) - 16.0) > 0.6:                              # varsayılan 36'lık durak − 20 (LeftIndent)
            e.append(f'varsayılan duraklar girintiye göre kaydırılmamış: {t2} (beklenen 16.0)')
        return e
    return exp


@senaryo('kenarliksiz_tablo', 'Kenarlığı hücre düzeyinde kapatılmış tablo (imza bloğu)')
def _(d):
    nb = ('<w:tcBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/>'
          '<w:right w:val="none"/></w:tcBorders>')
    make_docx(d, TBL([[P(''), P('Av. Deniz ÖRNEK')]], [4500, 4500], {(0, 0): nb, (0, 1): nb}) + P('son'))
    return lambda u: [] if u.els.find('table').get('border') == 'borderNone' else \
        [f'border={u.els.find("table").get("border")} (borderNone beklenir)']


@senaryo('koyu_zemin_beyaz_yazi', 'Koyu dolgulu hücrede beyaz yazı (dolgu taşınmaz → yazı kaybolmamalı)')
def _(d):
    make_docx(d, TBL([[P(R('BAŞLIK', '<w:b/><w:color w:val="FFFFFF"/>'))]], [9000],
                     {(0, 0): '<w:shd w:val="clear" w:color="auto" w:fill="1F497D"/>'}) + P('son'))

    def exp(u):
        c = [x for p in u.body() for x in p if x.tag == 'content' and u.sl(x) == 'BAŞLIK']
        return [] if c and c[0].get('background') else ['beyaz yazı zeminsiz kaldı (Editör\'de görünmez olur)']
    return exp


@senaryo('altbilgi_sayfa_kalibi', 'Altbilgide "Sayfa X / Y" kalıbı (artık metin kalmamalı)')
def _(d):
    fld = lambda ins: ('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>' + ins +
                       '</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' + R('7') +
                       '<w:r><w:fldChar w:fldCharType="end"/></w:r>')
    make_docx(d, P('Gövde'), footer=P('Büro adresi') + P(R('Sayfa ') + fld('PAGE') + R(' / ') + fld('NUMPAGES')))

    def exp(u):
        f = u.els.find('footer')
        tx = [u.ptext(p) for p in f.iter('paragraph')]
        e = []
        if any('Sayfa' in t for t in tx):
            e.append(f'altbilgide kalıp artığı: {tx}')
        if 'Büro adresi' not in tx:
            e.append('altbilgi metni kayboldu')
        if (f.get('pageNumber-spec') or 'BSP32_0') == 'BSP32_0':
            e.append('sayfa numarası kapalı kaldı')
        return e
    return exp


@senaryo('cdata_kacis', 'Metinde ]]> , & ve < karakterleri (Editör bölünmüş CDATA okuyamaz)')
def _(d):
    s = 'Koşul: a<b && c>d ise x[i[0]]>5 olur'
    make_docx(d, P(R(s) + R(' KALIN', '<w:b/>')))

    def exp(u):
        e = [] if u.texts()[0] == s.replace(']]>', ']] >') + ' KALIN' else [f'metin: {u.texts()[0]!r}']
        if u.contents(bold='true') != [' KALIN']:
            e.append(f'kalın dilim kaydı: {u.contents(bold="true")}')
        return e
    return exp


@senaryo('bmp_disi_karakter', 'Emoji / BMP dışı karakter: Editör açamaz → değiştirilmeli, dilimler kaymamalı')
def _(d):
    make_docx(d, P(R('Not 📎 ek: ') + R('KALIN', '<w:b/>') + R(' son')))

    def exp(u):
        e = []
        if any(ord(ch) > 0xFFFF for ch in u.cdata):
            e.append('CDATA\'da BMP dışı karakter kaldı (Editör bu dosyayı açamaz)')
        if u.contents(bold='true') != ['KALIN']:
            e.append(f'kalın dilim kaydı: {u.contents(bold="true")}')
        if not any('emoji' in x for x in u.warn):
            e.append('değiştirme için UYARI yok')
        return e
    return exp


@senaryo('kontrol_karakteri', 'XML\'de geçersiz denetim karakteri')
def _(d):
    make_docx(d, P('temiz'))
    # belgeye sonradan ham denetim karakteri sok (Word bunu üretmez ama bozuk dosyalarda görülür)
    z = zipfile.ZipFile(d)
    items = {n: z.read(n) for n in z.namelist()}
    z.close()
    items['word/document.xml'] = items['word/document.xml'].replace('temiz'.encode(), 'te&#x2028;miz­'.encode())
    with zipfile.ZipFile(d, 'w', zipfile.ZIP_DEFLATED) as z2:
        for n, b in items.items():
            z2.writestr(n, b)
    return lambda u: [] if 'te' in u.texts()[0] and 'miz' in u.texts()[0] else [f'metin: {u.texts()!r}']


@senaryo('yatay_sayfa', 'Yatay (landscape) sayfa')
def _(d):
    sect = ('<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1417" '
            'w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')
    make_docx(d, P('Yatay belge'), sect=sect)
    return lambda u: [] if u.page.get('paperOrientation') == '0' else \
        [f'paperOrientation={u.page.get("paperOrientation")} (yatay için 0 beklenir)']


@senaryo('kagit_a5', 'A5 sayfa (kitapçık): Editör kâğıt kodu 2 yazılmalı, A4\'e zorlanmamalı')
def _(d):
    sect = ('<w:sectPr><w:pgSz w:w="8391" w:h="11906" w:code="11"/><w:pgMar w:top="851" w:right="851" '
            'w:bottom="851" w:left="851" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>')
    make_docx(d, P('A5 kitapçık sayfası'), sect=sect)
    return lambda u: [] if u.page.get('mediaSizeName') == '2' else \
        [f'mediaSizeName={u.page.get("mediaSizeName")} (A5 için 2 beklenir)']


@senaryo('stilsiz_docx', 'styles.xml / numbering.xml olmayan DOCX')
def _(d):
    make_docx(d, P('Stilsiz belge') + TBL([[P('a'), P('b')]], [4500, 4500]) + P('son'), styles=False)
    return lambda u: [] if u.texts()[0] == 'Stilsiz belge' else [f'metin: {u.texts()}']


@senaryo('bos_belge', 'Boş gövde')
def _(d):
    make_docx(d, '')
    return lambda u: []


@senaryo('ilk_sayfa_farkli_ustbilgi', 'İlk sayfa üstbilgisi ≠ diğer sayfalar (ikisi de dolu)')
def _(d):
    sect = ('<w:sectPr><w:headerReference w:type="default" r:id="rId20"/><w:headerReference '
            'w:type="first" r:id="rId22"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" '
            'w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>'
            '<w:titlePg/></w:sectPr>')
    make_docx(d, P('Gövde'), header=P('DEVAM SAYFASI'), sect=sect)
    z = zipfile.ZipFile(d)
    items = {n: z.read(n) for n in z.namelist()}
    z.close()
    items['word/header2.xml'] = items['word/header1.xml'].replace('DEVAM SAYFASI'.encode(), 'İLK SAYFA ANTETİ'.encode())
    items['word/_rels/document.xml.rels'] = items['word/_rels/document.xml.rels'].replace(
        b'</Relationships>', b'<Relationship Id="rId22" Type="http://schemas.openxmlformats.org/'
        b'officeDocument/2006/relationships/header" Target="header2.xml"/></Relationships>')
    with zipfile.ZipFile(d, 'w', zipfile.ZIP_DEFLATED) as z2:
        for n, b in items.items():
            z2.writestr(n, b)

    def exp(u):
        hs = u.els.findall('header')
        txt = [(h.get('startPage'), h.get('stopPage'), ''.join(u.ptext(p) for p in h.iter('paragraph'))) for h in hs]
        ok = any('İLK SAYFA' in t for _, _, t in txt) and any('DEVAM' in t for _, _, t in txt)
        return [] if ok or u.warn else [f'ikinci üstbilgi sessizce kayboldu: {txt}']
    return exp


def anchor_zemin(rid, w_pt, h_pt, x_pt=0.0, y_pt=0.0):
    """Üstbilgide metnin arkasında duran (wrapNone, behindDoc) serbest görsel: antet / sayfa zemini."""
    return (anchor2(rid, w_pt, h_pt, wrap='wrapNone')
            .replace('behindDoc="0"', 'behindDoc="1"')
            .replace('<wp:positionH relativeFrom="column"><wp:align>left</wp:align></wp:positionH>',
                     f'<wp:positionH relativeFrom="page"><wp:posOffset>{int(x_pt * 12700)}</wp:posOffset></wp:positionH>')
            .replace('<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>',
                     f'<wp:positionV relativeFrom="page"><wp:posOffset>{int(y_pt * 12700)}</wp:posOffset></wp:positionV>'))


@senaryo('antet_tam_sayfa', 'Üstbilgide metnin arkasında tam sayfa antet → sayfa arka planı (üstbilgiye gömülürse sayfa sayısı katlanır)')
def _(d):
    make_docx(d, ''.join(P(f'Gövde paragrafı {i + 1}. ' + 'Metin ' * 30) for i in range(6)),
              header=P(anchor_zemin('rId1', 595.3, 841.9)), header_media={'rId1': ('antet.png', png(120, 170))})

    def exp(u):
        bg = u.root.find('properties/bgImage')
        e = []
        if bg is None or len(bg.get('bgImageData') or '') < 100:
            return ['antet sayfa arka planına (bgImage) yazılmadı']
        if any(bg.get(k) != '0' for k in ('bgImageLeftMargin', 'bgImageUpMargin', 'bgImageRigtMargin', 'bgImageBottomMargin')):
            e.append(f'tam sayfa antette kenar payı olmamalı: { {k: v for k, v in bg.attrib.items() if "Margin" in k} }')
        if u.images():
            e.append('antet ayrıca satır içi görsel olarak da yazılmış (üstbilgi sayfayı doldurur)')
        if u.els.find('header') is not None:
            e.append('boş üstbilgi öğesi yazılmış')
        if not any('arka plan' in x for x in u.warn):
            e.append('arka plan uyarısı yok')
        return e
    return exp


@senaryo('antet_konumlu_logo', 'Üstbilgide metnin arkasında konumlu logo + üstbilgi yazısı → kenar paylı arka plan, yazı korunur')
def _(d):
    make_docx(d, P('Gövde'), header=P(anchor_zemin('rId1', 120, 60, 100, 20) + R('BÜRO ADI')),
              header_media={'rId1': ('logo.png', png(120, 60))})

    def exp(u):
        bg = u.root.find('properties/bgImage')
        got = tuple(bg.get(k) for k in ('bgImageLeftMargin', 'bgImageUpMargin', 'bgImageRigtMargin', 'bgImageBottomMargin'))
        e = []
        if got != ('100', '20', '375', '762'):
            e.append(f'kenar payları {got} (beklenen 100, 20, 375, 762)')
        h = u.els.find('header')
        if h is None or 'BÜRO ADI' not in ''.join(u.ptext(p) for p in h.iter('paragraph')):
            e.append('üstbilgi yazısı kayboldu')
        if h is not None and list(h.iter('image')):
            e.append('logo üstbilgide satır içi de yazılmış')
        return e
    return exp


def vml_filigran(rid, w_pt, h_pt):
    """Word'ün "Filigran → Resim" ile eklediği üstbilgi görseli (VML, konumlu, metnin arkasında)."""
    return ('<w:r><w:pict><v:shape id="WordPictureWatermark1" type="#_x0000_t75" style="position:absolute;'
            f'margin-left:0;margin-top:0;width:{w_pt}pt;height:{h_pt}pt;z-index:-251653120;'
            'mso-position-horizontal:center;mso-position-horizontal-relative:margin;'
            'mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f">'
            f'<v:imagedata r:id="{rid}" o:title="antet"/></v:shape></w:pict></w:r>')


@senaryo('antet_vml_filigran', 'Antet Word filigranı (VML) olarak eklenmiş → sayfa arka planı, üstbilgi boş kalır')
def _(d):
    make_docx(d, P('Gövde'), header=P(vml_filigran('rId1', 596, 842)),
              header_media={'rId1': ('antet.png', png(120, 170))})

    def exp(u):
        bg = u.root.find('properties/bgImage')
        e = []
        if bg is None or len(bg.get('bgImageData') or '') < 100:
            return ['VML filigran anteti sayfa arka planına yazılmadı']
        got = tuple(bg.get(k) for k in ('bgImageLeftMargin', 'bgImageUpMargin', 'bgImageRigtMargin', 'bgImageBottomMargin'))
        if got != ('0', '0', '0', '0'):
            e.append(f'tam sayfa filigranda pay olmamalı: {got}')
        if u.images() or u.els.find('header') is not None:
            e.append('antet üstbilgiye de yazılmış')
        return e
    return exp


def _pdf_ornek():
    """Tek sayfalık, içinde dolu bir dikdörtgen olan en küçük geçerli PDF (Canva anteti gibi vektör görsel)."""
    objs = [b'<< /Type /Catalog /Pages 2 0 R >>', b'<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
            b'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /Contents 4 0 R >>']
    st = b'0.1 0.3 0.5 rg 10 10 180 80 re f'
    objs.append(b'<< /Length %d >>\nstream\n' % len(st) + st + b'\nendstream')
    out, offs = b'%PDF-1.4\n', []
    for i, o in enumerate(objs, 1):
        offs.append(len(out))
        out += b'%d 0 obj\n' % i + o + b'\nendobj\n'
    x = len(out)
    out += b'xref\n0 %d\n0000000000 65535 f \n' % (len(objs) + 1) + b''.join(b'%010d 00000 n \n' % o for o in offs)
    return out + b'trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n' % (len(objs) + 1, x)


@senaryo('gorsel_pdf_wmf_adli', 'Mac Word\'ün .wmf adıyla sakladığı PDF görsel → PNG\'ye çizilip taşınmalı')
def _(d):
    if not (shutil.which('sips') or shutil.which('powershell')):
        return None                               # çizici yok: bu makinede sınanamaz
    make_docx(d, P('Önce') + P(inline('rId10', 200, 100)) + P('Sonra'), media={'rId10': ('image1.wmf', _pdf_ornek())})

    def exp(u):
        ims = u.images('body')
        if len(ims) != 1:
            return [f'PDF görsel taşınmadı ({len(ims)} görsel); uyarılar: {u.warn}']
        data = base64.b64decode(ims[0].get('imageData'))
        return [] if data[:4] == b'\x89PNG' else ['görsel PNG değil']
    return exp


@senaryo('altbilgi_bos', 'Word\'de altbilgi tanımlı ama boş → "Word\'deki gibi" ayarında sayfa numarası EKLENMEMELİ')
def _(d):
    make_docx(d, P('Gövde'), footer=P(''))

    def exp(u):
        f = u.els.find('footer')
        spec = f.get('pageNumber-spec') if f is not None else None
        return [] if spec in (None, 'BSP32_0') else [f'boş altbilgiye sayfa numarası eklendi ({spec})']
    return exp


@senaryo('satir_sonu_cr', 'Word\'ün eski satır sonu işareti (w:cr, .doc dönüşümlerinde çıkıyor): iki satır birleşmemeli')
def _(d):
    make_docx(d, P(R('Çankaya / ANKARA') + '<w:r><w:cr/></w:r>' + R('VEKİLİ', '<w:b/>')))
    return lambda u: [] if u.texts() == ['Çankaya / ANKARA', 'VEKİLİ'] else [f'metin: {u.texts()}']


@senaryo('tablo_hucre_kenarligi', 'Kenarlığı yalnız hücrelerde tanımlı tablo (webden/yapay zekâdan yapıştırılan) → çizgili kalmalı')
def _(d):
    kenar = ''.join(f'<w:{k} w:val="inset" w:sz="6" w:space="0" w:color="A0A0A0"/>' for k in ('top', 'left', 'bottom', 'right'))
    hucre = lambda t: (f'<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:tcBorders>{kenar}</w:tcBorders>'
                       f'<w:vAlign w:val="center"/></w:tcPr>{P(t)}</w:tc>')
    tablo = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblCellSpacing w:w="15" w:type="dxa"/>'
             '<w:tblBorders>' + ''.join(f'<w:{k} w:val="outset" w:sz="6" w:space="0" w:color="A0A0A0"/>'
                                         for k in ('top', 'left', 'bottom', 'right')) + '</w:tblBorders></w:tblPr>'
             '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>'
             f'<w:tr>{hucre("İnceleme Unsuru")}{hucre("Açıklama")}</w:tr>'
             f'<w:tr>{hucre("Hologram")}{hucre("Bulunmamaktadır")}</w:tr></w:tbl>')
    make_docx(d, tablo + P('son'))

    def exp(u):
        t = u.els.find('table')
        return [] if t is not None and t.get('border') == 'borderCell' else [f'kenarlık: {t.get("border") if t is not None else None}']
    return exp


@senaryo('gorsel_cozunurluk', 'Kameradan gelen büyük fotoğraf sayfadaki boyutuna (200 dpi) indirilmeli; UDF şişmemeli')
def _(d):
    data = png(3000, 2000, (90, 120, 150))
    make_docx(d, P(inline('rId10', 150, 100)), media={'rId10': ('kamera.png', data)})

    def exp(u):
        ims = u.images('body')
        if len(ims) != 1:
            return ['görsel kayboldu']
        px = struct.unpack('>II', base64.b64decode(ims[0].get('imageData'))[16:24])
        try:
            import PIL  # noqa: F401  (küçültme Pillow ister; yoksa görsel aynen kalır)
        except ImportError:
            return []
        hedef = int(150 / 72 * 200 + 0.5)
        return [] if abs(px[0] - hedef) <= 2 else [f'genişlik {px[0]} px (beklenen ~{hedef})']
    return exp


@senaryo('pandoc_belgesi', 'Pandoc üretimi DOCX (tablo, görsel, dipnot, iç içe liste, alıntı)')
def _(d):
    if not shutil.which('pandoc'):
        return None
    wd = tempfile.mkdtemp()
    open(os.path.join(wd, 'g.png'), 'wb').write(png(200, 80))
    md = ('# AÇIKLAMALAR\n\nMetin[^1] ve **kalın**.\n\n[^1]: Dipnot metni burada.\n\n'
          '| Tarih | Olay |\n|---|---|\n| 01.01.2026 | Tebliğ |\n\n![Şekil](g.png)\n\n'
          '1. bir\n   a. alt\n2. iki\n\n> Alıntı paragrafı.\n')
    open(os.path.join(wd, 'a.md'), 'w', encoding='utf-8').write(md)
    subprocess.run(['pandoc', 'a.md', '-o', d], cwd=wd, check=True)
    shutil.rmtree(wd, ignore_errors=True)

    def exp(u):
        t = u.alltext()
        e = []
        for k in ('AÇIKLAMALAR', 'Tebliğ', 'Alıntı paragrafı.', 'Dipnot metni burada.'):
            if k not in t:
                e.append(f'kayıp: {k}')
        if not u.images('body'):
            e.append('görsel yok')
        if u.els.find('table') is None:
            e.append('tablo yok')
        return e
    return exp


# ---------- 18.10.2026: Argon (Gemini) denetiminden çıkan ve düzeltilen hatalar ----------
def _fld(*parts):
    """Karmaşık alan; yönerge birden çok instrText run'ına bölünebilir (Word bunu yapıyor)."""
    return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            + ''.join(f'<w:r><w:instrText xml:space="preserve">{x}</w:instrText></w:r>' for x in parts)
            + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' + R('1')
            + '<w:r><w:fldChar w:fldCharType="end"/></w:r>')


def _hf(u, tag):
    e = u.els.find(tag)
    return (None, []) if e is None else (e.get('pageNumber-spec'), [u.ptext(p) for p in e.iter('paragraph')])


@senaryo('gorsel_goreli_iliski', 'İlişki hedefi göreli yol ("../word/media/…"): görsel düşmemeli')
def _(d):
    data = png(120, 60)
    make_docx(d, P('Önce') + P(inline('rId10', 60, 30)) + P('Sonra'), media={'rId10': ('image1.png', data)})
    with zipfile.ZipFile(d) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    parts['word/_rels/document.xml.rels'] = parts['word/_rels/document.xml.rels'].replace(
        b'Target="media/image1.png"', b'Target="../word/./media/image1.png"')
    with zipfile.ZipFile(d, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, b in parts.items():
            z.writestr(n, b)

    def exp(u):
        ims = u.images('body')
        return [] if len(ims) == 1 and img_bytes(ims[0]) == data else [f'görsel taşınmadı ({len(ims)}); {u.warn}']
    return exp


@senaryo('altbilgi_metin_ve_sayfa_no', 'Altbilgide "Av. Ali Kaya - Sayfa" + PAGE: metin silinmemeli')
def _(d):
    make_docx(d, P('Gövde'), footer=P(R('Av. Ali Kaya - Sayfa ') + _fld(' PAGE ')))

    def exp(u):
        spec, tx = _hf(u, 'footer')
        e = []
        if not any('Av. Ali Kaya' in t for t in tx):
            e.append(f'altbilgi metni kayboldu: {tx}')
        if spec in (None, 'BSP32_0'):
            e.append(f'sayfa numarası yok: {spec}')
        return e
    return exp


@senaryo('ustbilgi_sayfa_no', 'Sayfa numarası yalnız üstbilgide, altbilgi metinli: numara ve büro adı kalmalı')
def _(d):
    make_docx(d, P('Gövde'), header=P(R('Bilir Hukuk Bürosu ') + _fld(' PAGE ')), footer=P('Adres satırı'))

    def exp(u):
        e = []
        _, htx = _hf(u, 'header')
        spec, ftx = _hf(u, 'footer')
        if not any('Bilir Hukuk Bürosu' in t for t in htx):
            e.append(f'üstbilgi metni kayboldu: {htx}')
        if spec in (None, 'BSP32_0'):
            e.append(f'sayfa numarası altbilgiye yazılmadı: {spec}')
        if not any('Adres' in t for t in ftx):
            e.append(f'altbilgi metni: {ftx}')
        return e
    return exp


@senaryo('alan_yonergesi_bolunmus', 'PAGE yönergesi iki run\'a bölünmüş ("PA" + "GE"): sabit "1" yazılmamalı')
def _(d):
    make_docx(d, P('Gövde'), footer=P(R('Sayfa ') + _fld(' PA', 'GE ')))

    def exp(u):
        spec, tx = _hf(u, 'footer')
        e = []
        if any('1' in t for t in tx):
            e.append(f'alan önbelleği metne girdi: {tx}')
        if spec in (None, 'BSP32_0'):
            e.append(f'sayfa numarası tanınmadı: {spec}')
        return e
    return exp


@senaryo('tablo_dar_son_sutun', 'Çok dar son sütun: columnSpans sıfır/negatif olmamalı')
def _(d):
    make_docx(d, TBL([[P(f'h{i}') for i in range(5)]], [1248] * 4 + [20]))

    def exp(u):
        t = u.els.find('table')
        sp = [int(x) for x in t.get('columnSpans').split(',')]
        e = [] if min(sp) >= 1 and sum(sp) == 100 * len(sp) else [f'columnSpans {sp}']
        return e
    return exp


def bad_inputs(workdir):
    """Geçersiz girdiler: temiz hata mesajı (traceback yok) ve sıfırdan farklı çıkış kodu."""
    res = []
    cases = {'eski_doc.docx': b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1' + b'\x00' * 600,
             'duz_metin.docx': 'bu bir docx değil'.encode(),
             'olmayan.docx': None}
    for name, data in cases.items():
        p = os.path.join(workdir, name)
        if data is not None:
            open(p, 'wb').write(data)
        r = subprocess.run([sys.executable, TOOL, p, '-o', p + '.udf'], capture_output=True, **UTF8)
        errs = []
        if r.returncode == 0:
            errs.append('çıkış kodu 0')
        if 'Traceback' in r.stderr:
            errs.append('Python traceback sızdı')
        res.append((f'gecersiz_girdi:{name}', 'Geçersiz/bozuk girdi → temiz hata', errs))
    return res


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('-k', default='', help='senaryo adına göre süz')
    ap.add_argument('--tut', help='üretilen DOCX/UDF dosyalarını bu klasörde sakla')
    ap.add_argument('--editor', action='store_true',
                    help='her UDF\'yi kurulu UYAP Editör motoruyla da yükle (udf_onizle.py); açılamazsa KALDI')
    a = ap.parse_args()
    work = a.tut or tempfile.mkdtemp(prefix='udf_senaryo_')
    os.makedirs(work, exist_ok=True)
    rows = []
    for name, note, fn in SCEN:
        if a.k and a.k not in name:
            continue
        docx = os.path.join(work, name + '.docx')
        udf = os.path.join(work, name + '.udf')
        try:
            exp = fn(docx)
        except Exception as ex:                                # fixture üretilemedi
            rows.append((name, note, 'ATLANDI', [f'fixture: {ex}']))
            continue
        if exp is None:
            rows.append((name, note, 'ATLANDI', ['gerekli araç yok']))
            continue
        r = subprocess.run([sys.executable, TOOL, docx, '-o', udf] + SCEN_ARGS.get(name, []),
                           capture_output=True, **UTF8)
        if r.returncode != 0:
            last = (r.stderr.strip().splitlines() or r.stdout.strip().splitlines() or ['?'])[-1]
            rows.append((name, note, 'ÇÖKTÜ', [last[:160]]))
            continue
        try:
            u = U(udf, r.stdout)
            errs = u.generic() + exp(u)
        except Exception as ex:
            errs = [f'UDF okunamadı/denetim hatası: {type(ex).__name__}: {ex}'[:200]]
        if not errs:
            # Masaüstü uygulaması her çıktıyı udf_dogrula.dogrula() ile denetler ve KALDI derse UDF'yi
            # yazmaz; doğrulayıcının yanlış alarmı kullanıcıya "çevrilemedi" olarak yansır (21.09.2026'da
            # yaşandı). Bu yüzden aynı denetim her senaryoda burada da koşar.
            try:
                import udf_dogrula
                ok, lines = udf_dogrula.dogrula(docx, udf)
                if not ok and name not in DOGRULAMA_HARIC:
                    errs = ['udf_dogrula KALDI: ' + '; '.join(l for l in lines if l.startswith('HATA'))[:200]]
            except Exception as ex:
                errs = [f'udf_dogrula çöktü: {type(ex).__name__}: {ex}'[:200]]
        if a.editor and not errs:
            rr = subprocess.run([sys.executable, os.path.join(HERE, 'udf_onizle.py'), udf,
                                 '-o', os.path.join(work, '_onizleme')], capture_output=True, **UTF8)
            if rr.returncode != 0:
                errs = ['Editör motoru bu UDF\'yi açamadı/çizemedi: '
                        + ((rr.stderr or rr.stdout).strip().splitlines() or ['?'])[0][:140]]
        rows.append((name, note, 'GEÇTİ' if not errs else 'KALDI', errs))
    if not a.k or 'gecersiz' in a.k:
        for name, note, errs in bad_inputs(work):
            rows.append((name, note, 'GEÇTİ' if not errs else 'KALDI', errs))
    wn = max(len(r[0]) for r in rows)
    for name, note, st, errs in rows:
        print(f'{st:8} {name:{wn}}  {note}')
        for e in errs:
            print(f'{"":8} {"":{wn}}    ↳ {e}')
    tot = {s: sum(1 for r in rows if r[2] == s) for s in ('GEÇTİ', 'KALDI', 'ÇÖKTÜ', 'ATLANDI')}
    print('\nÖZET:', ', '.join(f'{k} {v}' for k, v in tot.items() if v), f'| dosyalar: {work}')
    if not a.tut:
        shutil.rmtree(work, ignore_errors=True)
    sys.exit(0 if tot['KALDI'] == 0 and tot['ÇÖKTÜ'] == 0 else 1)


if __name__ == '__main__':
    main()
