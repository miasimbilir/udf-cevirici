#!/usr/bin/env python3
"""udf_dogrula.py — DOCX ile ondan üretilen UDF'yi (ya da UDF ile ondan üretilen DOCX'i) karşılaştırır.

Kesin denetimler (biri bozuksa KALDI):
  1. UDF geçerli bir zip ve XML mi?
  2. Biçim dilimleri CDATA metnini boşluksuz ve taşmasız kaplıyor mu? (Editör Java
     olduğu için ofsetler UTF-16 birimiyle sayılır.)
  3. Her paragraf '\\n' ile bitiyor mu; tablo satırlarında hücre sayısı sütun tanımına
     (tablonun columnCount'u ya da satırın kendi columnSpans'ı) uyuyor mu?
  4. DOCX'teki her kelime UDF'de var mı? (Sıra aranmaz: dipnot belge sonuna, metin
     kutusu paragrafın ardına taşındığı için. Büyük/küçük harf farkı sayılmaz.)
Bilgi satırları: paragraf/tablo/görsel/sayfa sonu sayıları, üstbilgi/altbilgi.

Ters yön (UDF -> DOCX, udf_docx.py çıktısı) için kesin denetimler:
  1. DOCX geçerli bir zip mi, zorunlu parçalar var mı, bütün XML parçaları ayrıştırılıyor mu?
  2. Her ilişkinin (rels) hedefi arşivde var mı, her parçanın içerik türü tanımlı mı?
  3. UDF'deki her kelime DOCX'te (gövde + üstbilgi + altbilgi) var mı?
  4. Tablo ve görsel sayıları tutuyor mu; her tablo satırının sütun toplamı ızgaraya uyuyor mu?

Kullanım:
    python3 araclar/udf_dogrula.py belge.docx belge.udf      (DOCX -> UDF denetimi)
    python3 araclar/udf_dogrula.py belge.udf belge.docx      (UDF -> DOCX denetimi)
Çıkış kodu 0 = GEÇTİ, 1 = KALDI.
"""
import base64
import posixpath
import re
import sys
import unicodedata
import zipfile
from xml.etree import ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def w(t):
    return f'{{{W}}}{t}'


def norm(s):
    # NFKC: araç alt simgeyi Unicode karaktere çevirir (CO₂); karşılaştırmada CO2 ile eşlensin
    return unicodedata.normalize('NFKC', s).replace('i', 'İ').replace('ı', 'I').upper()


def words(text):
    return set(re.findall(r'[^\W_]+', norm(text), re.UNICODE))


def docx_words(path, ust_alt=False, ayir=False):
    z = zipfile.ZipFile(path)
    out = set()
    pat = r'word/(footnotes|endnotes|header\d*|footer\d*)\.xml' if ust_alt else r'word/(footnotes|endnotes)\.xml'
    parts = ['word/document.xml'] + [n for n in z.namelist() if re.fullmatch(pat, n)]
    for part in parts:
        root = ET.fromstring(z.read(part))
        dead = set()
        for tag in ('del', 'moveFrom'):
            for el in root.iter(w(tag)):
                dead.update(id(t) for t in el.iter(w('t')))
        # ayir: üst/alt simge (dipnot işareti, m²) komşu kelimeden ayrılır. Yalnız UDF -> DOCX denetiminde
        # kullanılır; DOCX -> UDF yönünde araç "CO" + alt simge "2"yi tek kelime "CO₂" yaptığı için AYRILMAZ
        # (21.09.2026: ayırma iki yöne birden uygulanınca CO₂ içeren her belge yanlış alarmla KALDI).
        raised = {id(t) for r in root.iter(w('r')) if r.find(f'{w("rPr")}/{w("vertAlign")}') is not None
                  for t in r.iter(w('t'))} if ayir else set()
        for p in root.iter(w('p')):               # kelime run'lara bölünmüş olabilir: paragraf bazında birleştir
            inner = {id(t) for q in p.iter(w('p')) if q is not p for t in q.iter()}   # metin kutusu vb.
            out |= words(''.join((f' {t.text or ""} ' if id(t) in raised else (t.text or '')) if t.tag == w('t') else ' '
                                 for t in p.iter() if t.tag in (w('t'), w('tab'), w('br'), w('cr'))
                                 and id(t) not in dead and id(t) not in inner))
    return out


def dogrula(docx, udf):
    """(gecti_mi, satirlar) döndürür; satirlar ekrana/rapora yazılacak metinlerdir."""
    out, errs = [], []
    try:
        z = zipfile.ZipFile(udf)
        x = z.read('content.xml').decode('utf-8')
        root = ET.fromstring(x.split('?>', 1)[1].strip())
    except Exception as ex:
        return False, [f'HATA: UDF okunamadı: {ex}', 'SONUÇ: KALDI']
    cdata = root.find('content').text or ''
    cd16 = cdata.encode('utf-16-le')
    els = root.find('elements')

    def sl(c):
        so, ln = int(c.get('startOffset')), int(c.get('length'))
        return cd16[2 * so:2 * (so + ln)].decode('utf-16-le', 'replace')

    pos, gaps, noeol, npara = 0, 0, 0, 0
    for p in els.iter('paragraph'):
        cs = [c for c in p if c.get('startOffset') is not None]
        npara += 1
        for c in cs:
            so = int(c.get('startOffset'))
            if so != pos:
                gaps += 1
                pos = so
            pos += int(c.get('length'))
        if cs and not sl(cs[-1]).endswith('\n'):
            noeol += 1
    total = len(cd16) // 2
    if gaps or pos != total:
        errs.append(f'CDATA kapsama bozuk: {gaps} boşluk, son ofset {pos} / {total}')
    if noeol:
        errs.append(f'{noeol} paragraf satır sonu ile bitmiyor')
    out.append(f'Ofset zinciri: {pos}/{total} (UTF-16 birimi)' + (' OK' if not gaps and pos == total else ' HATA'))

    astral = sorted({ch for ch in cdata if ord(ch) > 0xFFFF})
    if astral:
        errs.append('metinde Editör\'ün açamadığı karakter var (emoji vb.): ' + ' '.join(astral))

    for t in els.iter('table'):
        n = int(t.get('columnCount') or 0)
        rows = t.findall('row')
        bad = [i + 1 for i, r in enumerate(rows)
               if len(r.findall('cell')) != (len(r.get('columnSpans').split(',')) if r.get('columnSpans') else n)]
        merged = sum(1 for r in rows if r.get('columnSpans'))
        out.append(f'Tablo: {n} sütun, {len(rows)} satır, kenarlık {t.get("border")}'
                   + (f', {merged} satırda yatay birleştirme' if merged else '')
                   + (f'  HATA: tutarsız satırlar {bad}' if bad else ''))
        if bad:
            errs.append(f'tabloda tutarsız satırlar: {bad}')

    # üst/alt simge dilimleri (dipnot işareti gibi) komşu kelimeye yapışmasın diye boşlukla ayrılır
    pieces = []
    for c in els.iter('content'):
        if c.get('startOffset') is not None:
            t = sl(c)
            pieces.append(f' {t} ' if c.get('superscript') or c.get('subscript') else t)
    udf_words = words(''.join(pieces)) | words(cdata)
    try:
        dw = docx_words(docx)
    except Exception as ex:
        return False, out + [f'HATA: DOCX okunamadı: {ex}', 'SONUÇ: KALDI']
    missing = sorted(dw - udf_words)
    out.append(f'Kelime kapsaması: DOCX\'teki {len(dw)} farklı kelimenin {len(dw) - len(missing)} tanesi UDF\'de var')
    if missing:
        errs.append(f'UDF\'de bulunmayan {len(missing)} kelime: {", ".join(missing[:15])}'
                    + (' …' if len(missing) > 15 else ''))

    body_imgs = sum(1 for e in els if e.tag in ('paragraph', 'table') for _ in e.iter('image'))
    out.append(f'Sayım: {npara} paragraf, {len(els.findall("table"))} tablo, '
               f'{len(els.findall("page-break"))} sayfa sonu, {body_imgs} gövde görseli')
    for tag, ad in (('header', 'Üstbilgi'), ('footer', 'Altbilgi')):
        for e in els.findall(tag):
            rng = ' (yalnız ilk sayfa)' if e.get('stopPage') == '1' else \
                f' ({e.get("startPage")}. sayfadan itibaren)' if e.get('startPage') else ''
            pn = ', sayfa numaralı' if (e.get('pageNumber-spec') or 'BSP32_0') != 'BSP32_0' else ''
            out.append(f'{ad}: {len(e.findall("paragraph"))} paragraf, {len(list(e.iter("image")))} görsel{rng}{pn}')
    out += [f'HATA: {e}' for e in errs]
    out.append('SONUÇ: ' + ('GEÇTİ' if not errs else 'KALDI'))
    return not errs, out


def _img_ok(b64):
    try:
        data = base64.b64decode(re.sub(r'\s+', '', b64 or ''))
    except (ValueError, TypeError):
        return False
    return data[:8] == b'\x89PNG\r\n\x1a\n' or data[:2] in (b'\xff\xd8', b'BM') or data[:6] in (b'GIF87a', b'GIF89a')


def dogrula_docx(udf, docx):
    """Ters yön: UDF'den üretilen DOCX'in denetimi. (gecti_mi, satirlar) döndürür."""
    out, errs = [], []
    try:
        raw = zipfile.ZipFile(udf).read('content.xml')
        try:
            root = ET.fromstring(raw)
        except ET.ParseError:                     # geçersiz denetim karakteri: ofset kaymasın diye boşlukla değiştir
            txt = re.sub('[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]', ' ', raw.decode('utf-8', 'replace'))
            root = ET.fromstring(txt.split('?>', 1)[1].strip() if txt.lstrip().startswith('<?xml') else txt)
    except Exception as ex:
        return False, [f'HATA: UDF okunamadı: {ex}', 'SONUÇ: KALDI']
    try:
        z = zipfile.ZipFile(docx)
        names = set(z.namelist())
    except Exception as ex:
        return False, [f'HATA: DOCX okunamadı: {ex}', 'SONUÇ: KALDI']
    for need in ('[Content_Types].xml', '_rels/.rels', 'word/document.xml', 'word/styles.xml'):
        if need not in names:
            errs.append(f'DOCX içinde zorunlu parça yok: {need}')
    trees = {}
    for n in sorted(names):
        if n.endswith(('.xml', '.rels')):
            try:
                trees[n] = ET.fromstring(z.read(n))
            except ET.ParseError as ex:
                errs.append(f'{n} ayrıştırılamadı: {ex}')
    out.append(f'DOCX paketi: {len(names)} parça, {len(trees)} XML' + (' OK' if not errs else ' HATA'))

    # ilişkiler ve içerik türleri
    broken = []
    for n, t in trees.items():
        if not n.endswith('.rels'):
            continue
        base = posixpath.dirname(posixpath.dirname(n))
        for r in t:
            if r.get('TargetMode') == 'External':
                continue
            target = posixpath.normpath(posixpath.join(base, r.get('Target') or ''))
            if target not in names:
                broken.append(f'{n} -> {r.get("Target")}')
    if broken:
        errs.append('ilişki hedefi arşivde yok: ' + ', '.join(broken[:5]))
    ct = trees.get('[Content_Types].xml')
    if ct is not None:
        defaults = {e.get('Extension', '').lower() for e in ct if e.tag.endswith('Default')}
        overrides = {e.get('PartName') for e in ct if e.tag.endswith('Override')}
        untyped = [n for n in names if n != '[Content_Types].xml' and '/' + n not in overrides
                   and n.rsplit('.', 1)[-1].lower() not in defaults]
        if untyped:
            errs.append('içerik türü tanımsız parça: ' + ', '.join(untyped[:5]))

    # kelime kapsaması: UDF -> DOCX
    cdata = (root.find('content').text if root.find('content') is not None else '') or ''
    cd16 = cdata.encode('utf-16-le')
    pieces = []
    for c in root.iter():
        if c.get('startOffset') is None or c.tag == 'image':
            continue
        try:
            so, ln = int(c.get('startOffset')), int(c.get('length'))
        except ValueError:
            continue
        t = cd16[2 * so:2 * (so + ln)].decode('utf-16-le', 'replace')
        pieces.append(f' {t} ' if c.get('superscript') or c.get('subscript') else t)
    uw = words(''.join(pieces))
    try:
        dw = docx_words(docx, ust_alt=True, ayir=True) | docx_words(docx, ust_alt=True)
    except Exception as ex:
        return False, out + [f'HATA: DOCX metni okunamadı: {ex}', 'SONUÇ: KALDI']
    missing = sorted(uw - dw)
    out.append(f'Kelime kapsaması: UDF\'deki {len(uw)} farklı kelimenin {len(uw) - len(missing)} tanesi DOCX\'te var')
    if missing:
        errs.append(f'DOCX\'te bulunmayan {len(missing)} kelime: {", ".join(missing[:15])}'
                    + (' …' if len(missing) > 15 else ''))

    # sayımlar
    els = root.find('elements')
    doc = trees.get('word/document.xml')
    if els is not None and doc is not None:
        body_els = [e for e in els if e.tag not in ('header', 'footer')]
        # iç içe tablo sayılmaz: dikey birleştirme taklidi Word'de gerçek birleştirmeye açılır
        u_tbl = sum(1 for e in body_els if e.tag == 'table')
        # udf_docx okunamayan görseli uyarıyla atlar; yalnız tanınan görseller sayılır (18.10.2026)
        u_img = sum(1 for e in body_els for i in e.iter('image') if _img_ok(i.get('imageData')))
        d_tbl = len(doc.findall(f'{w("body")}/{w("tbl")}'))
        d_img = sum(1 for _ in doc.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}blip'))
        out.append(f'Sayım: UDF {u_tbl} tablo / {u_img} gövde görseli; DOCX {d_tbl} tablo / {d_img} gövde görseli '
                   '(tablolar üst düzey)')
        if u_tbl != d_tbl:
            errs.append(f'tablo sayısı tutmuyor (UDF {u_tbl}, DOCX {d_tbl})')
        if u_img != d_img:
            errs.append(f'gövde görseli sayısı tutmuyor (UDF {u_img}, DOCX {d_img})')
        for t in doc.iter(w('tbl')):
            grid = len(t.findall(f'{w("tblGrid")}/{w("gridCol")}'))
            for i, tr in enumerate(t.findall(w('tr'))):
                tot = 0
                for tc in tr.findall(w('tc')):
                    gs = tc.find(f'{w("tcPr")}/{w("gridSpan")}')
                    tot += int(gs.get(w('val'))) if gs is not None else 1
                if tot != grid:
                    errs.append(f'tablo satırı {i + 1}: sütun toplamı {tot}, ızgara {grid}')
                if tr.find(w('tc')) is not None and any(len(tc) and tc[-1].tag != w('p') for tc in tr.findall(w('tc'))):
                    errs.append(f'tablo satırı {i + 1}: hücre paragrafla bitmiyor (Word açamaz)')
    hf = sorted(n for n in names if re.fullmatch(r'word/(header|footer)\d+\.xml', n))
    if hf:
        out.append('Üstbilgi/altbilgi parçaları: ' + ', '.join(n.split('/')[-1] for n in hf))
    out += [f'HATA: {e}' for e in errs]
    out.append('SONUÇ: ' + ('GEÇTİ' if not errs else 'KALDI'))
    return not errs, out


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    if sys.argv[1].lower().endswith('.udf') and not sys.argv[2].lower().endswith('.udf'):
        ok, satirlar = dogrula_docx(sys.argv[1], sys.argv[2])
    else:
        ok, satirlar = dogrula(sys.argv[1], sys.argv[2])
    print('\n'.join(satirlar))
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()
