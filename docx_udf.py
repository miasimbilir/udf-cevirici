#!/usr/bin/env python3
"""DOCX -> UDF (UYAP Editör) dönüştürücü.

UDF, içinde tek bir content.xml bulunan zip arşividir:

    <template format_id="1.8">
      <content><![CDATA[ ...belgenin düz metni... ]]></content>
      <properties><pageFormat .../><bgImage .../></properties>
      <elements resolver="hvl-default">
        <header [stopPage="1"]><paragraph><image imageData=".." width=".." height=".."
                startOffset="0" length="1"/>..</paragraph></header>
        <paragraph Alignment=".." ...>
          <content bold=".." italic=".." underline=".." startOffset="N" length="M"/>
        </paragraph>
        <table tableName="Sabit" columnCount="N" columnSpans="a,b,.." border="borderCell">
          <row rowName="row1" rowType="dataRow"><cell><paragraph>..</paragraph></cell>..</row>
        </table>
        <footer ...><paragraph>..</paragraph></footer>
      </elements>
      <styles>...</styles>
    </template>

Biçim bilgisi metnin içinde değil; her <content>/<image> parçası CDATA metnine
startOffset/length ile işaret eder. Bu nedenle offsetlerin karakter bazında
birebir tutması şarttır. CDATA sırası: üstbilgi metni, gövde (tablo hücreleri
dâhil), altbilgi metni; her paragraf '\\n' ile biter, görsel tek bir '¸'
karakteriyle temsil edilir (Editör'ün kendi ürettiği dosyalardan alındı).

Taşınanlar: paragraf hizası, girinti (sol/sağ/asılı/ilk satır), üst-alt boşluk,
satır aralığı, özel sekme durakları, kalın/italik/altı çizili/üstü çizili/üst-alt
simge/punto/renk/vurgu/yazı tipi, "tümü büyük harf" (Türkçe i/ı ile), Word listeleri
(yerel UDF listesi; 1. a. i. I. A. ve madde işaretleri), tablolar (sütun oranı,
kenarlık türü, satır yüksekliği, YATAY birleştirme satır düzeyi columnSpans ile),
sayfa sonu (<page-break>), üstbilgi/altbilgi (metin + görsel; ilk sayfaya özel
olan stopPage="1", devamı startPage="2"), görseller (PNG doğrudan; JPEG/GIF/BMP
PNG'ye çevrilir; serbest konumlu ve VML görseller satır içine alınır; alana
sığmayan görsel orantılı küçültülür), dipnot/sonnot (işaret üst simge, metin belge
sonunda), metin kutusu (düz paragraf olarak), simgeler (w:sym).
Taşınmayanlar: DİKEY hücre birleştirme (boş hücre kalır), hücre içi tablo
(düzleşir), hücre gölgesi, EMF/WMF görsel, şekil/grafik/SmartArt, paragraf
kenarlığı (üstbilgi/altbilgide alt çizgi satırıyla taklit edilir), görselin
sayfadaki serbest konumu.
Ofsetler Editör Java olduğu için UTF-16 birimiyle sayılır (emoji vb. 2 birim).

Kullanım:
    python3 araclar/docx_udf.py girdi.docx -o cikti.udf
"""

import argparse
import base64
import io
import os
import posixpath
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import zipfile
from xml.etree import ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
V = 'urn:schemas-microsoft-com:vml'
NS = {'w': W}


def w(tag):
    return f'{{{W}}}{tag}'


class DonusumHatasi(Exception):
    """Kullanıcıya gösterilecek, anlaşılır dönüştürme hatası."""


ALIGN = {'left': 0, 'start': 0, 'center': 1, 'right': 2, 'end': 2,
         'both': 3, 'justify': 3, 'distribute': 3}
# Word sekme türü -> javax.swing.text.TabStop sabiti; lider -> LEAD_* sabiti
TAB_ALIGN = {'left': 0, 'start': 0, 'num': 0, 'right': 1, 'end': 1, 'center': 2,
             'decimal': 4, 'bar': 5}
TAB_LEADER = {'dot': 1, 'middleDot': 1, 'hyphen': 2, 'underscore': 3, 'heavy': 4}
TAB_EDGE = 4.0                        # sağ kenardaki hizalı sekmeyi bu kadar punto içeri al
DEFAULT_FONT = 'Times New Roman'      # UDF hvl-default stili; farklı olan run'a family yazılır
IMG_CHAR = '¸'                        # Editör'ün görsel yer tutucusu (CDATA'da 1 karakter)
# Hücre paragrafı için Editör'ün kendi kullandığı iç boşluk (örnek UDF'lerden)
CELL_LEFT, CELL_RIGHT = 3.0, 1.0

HIGHLIGHT = {'yellow': 'FFFF00', 'green': '00FF00', 'cyan': '00FFFF', 'magenta': 'FF00FF',
             'blue': '0000FF', 'red': 'FF0000', 'darkBlue': '000080', 'darkCyan': '008080',
             'darkGreen': '008000', 'darkMagenta': '800080', 'darkRed': '800000',
             'darkYellow': '808000', 'darkGray': '808080', 'lightGray': 'C0C0C0',
             'black': '000000', 'white': 'FFFFFF'}
# w:sym (simge yazı tipleri) -> Unicode
SYM = {('wingdings', 0xF0FC): '✓', ('wingdings', 0xF0FB): '✗', ('wingdings', 0xF0FE): '☑',
       ('wingdings', 0xF0FD): '☒', ('wingdings', 0xF0A8): '☐', ('wingdings', 0xF06F): '☐',
       ('wingdings', 0xF0A7): '▪', ('wingdings', 0xF0D8): '➢', ('wingdings', 0xF0E0): '→',
       ('wingdings', 0xF0DF): '←', ('wingdings', 0xF028): '☎', ('wingdings', 0xF02A): '✉',
       ('wingdings 2', 0xF050): '✓', ('wingdings 2', 0xF052): '☑', ('wingdings 2', 0xF0A3): '☐',
       ('symbol', 0xF0B7): '•', ('symbol', 0xF0B0): '°', ('symbol', 0xF0B1): '±',
       ('symbol', 0xF0A3): '≤', ('symbol', 0xF0B3): '≥', ('symbol', 0xF0B4): '×',
       ('symbol', 0xF0AE): '→', ('symbol', 0xF0D7): '⋅'}
# madde işareti karakteri -> UDF BulletType (derlemde görülen dört tür)
BULLET_ARROW = {0xF0D8, 0xF0E0, 0x27A2, 0x2192, 0x25BA, 0x2794, 0x21D2, 0xF0F0}
BULLET_DIAMOND = {0xF075, 0xF076, 0x25C6, 0x2666, 0x2756, 0x25C7, 0xF077}
BULLET_RECT = {0xF0A7, 0x25AA, 0x25A0, 0x25FC, 0xF06E}        # dolu kare (Word'ün 3. düzey işareti)
BULLET_RECT_D = {0xF06F, 0x25AB, 0x25A1, 0x25FB, 0x6F, 0xF071}  # içi boş kare / Word'ün 'o' işareti
BULLET_TRIANGLE = {0x25B6, 0x25B8, 0x25B2, 0x2023}
_BAD = re.compile('[\x00-\x08\x0b-\x1f\ufffe\uffff]')


def u16(s):
    """Editör (Java) ofsetleri UTF-16 birimiyle sayar; BMP dışı karakter 2 birimdir."""
    return len(s.encode('utf-16-le')) // 2


# Editör (JDOM 1.x) BMP dışı karakter içeren UDF'yi AÇAMAZ ("Error in building"; 20.09.2026'da
# Editör motoruyla sınandı). Emoji vb. BMP içi bir karşılıkla, yoksa □ ile değiştirilir.
_NONBMP = re.compile('[\U00010000-\U0010FFFF]')
_NONBMP_MAP = {'\U0001F4DE': '\u260e', '\U0001F4E7': '\u2709', '\U0001F4E9': '\u2709',
               '\U0001F449': '\u261e', '\U0001F538': '\u25c6', '\U0001F539': '\u25c6', '\U0001F4CC': '\u2022',
               '\U0001F534': '\u25cf', '\U0001F535': '\u25cf', '\U0001F7E2': '\u25cf', '\U0001F4C5': '\u25a1'}
NONBMP_SEEN = set()


def _nonbmp_sub(m):
    NONBMP_SEEN.add(m.group(0))
    return _NONBMP_MAP.get(m.group(0), '\u25a1')


def clean(s):
    return _NONBMP.sub(_nonbmp_sub, _BAD.sub('', s.replace('\n', ' ')))


# Editör motoru subscript="true" dilimini yukarıda (üst simge gibi) çiziyor (20.09.2026 önizleme).
# Karşılığı olan karakterler Unicode alt simgeye çevrilir; olmayanlarda öznitelik bırakılır.
SUBSCRIPT = str.maketrans('0123456789+-=()aehklmnopstx',
                          '\u2080\u2081\u2082\u2083\u2084\u2085\u2086\u2087\u2088\u2089\u208a\u208b\u208c\u208d\u208e'
                          '\u2090\u2091\u2095\u2096\u2097\u2098\u2099\u2092\u209a\u209b\u209c\u2093')


def tr_upper(s):
    return s.replace('i', 'İ').replace('ı', 'I').upper()


def _on(el):
    """<w:b/>, <w:b w:val="1"/> -> True ; <w:b w:val="0"/> -> False"""
    if el is None:
        return None
    v = el.get(w('val'))
    return v not in ('0', 'false', 'none')


def _merge(dst, src):
    """src'yi dst üzerine yazar; 'ind' öznitelik bazında, 'tabs' ekleyerek birleşir."""
    for k, v in src.items():
        if k == 'ind':
            d = dict(dst.get('ind', {}))
            d.update(v)
            dst['ind'] = d
        elif k == 'tabs':
            dst['tabs'] = list(dst.get('tabs', [])) + list(v)
        else:
            dst[k] = v
    return dst


def _ind(el):
    d = {}
    if el is None:
        return d
    for k, alt in (('left', 'start'), ('right', 'end'), ('hanging', None),
                   ('firstLine', None)):
        v = el.get(w(k))
        if v is None and alt:
            v = el.get(w(alt))
        if v is not None:
            try:
                d[k] = float(v) / 20          # twip -> punto
            except ValueError:
                pass
    return d


def _ppr_props(ppr):
    """<w:pPr> -> {jc, before, after, line, ind, numId, ilvl, tabs, border_top}"""
    d = {}
    if ppr is None:
        return d
    jc = ppr.find(w('jc'))
    if jc is not None and jc.get(w('val')) in ALIGN:
        d['jc'] = ALIGN[jc.get(w('val'))]
    sp = ppr.find(w('spacing'))
    if sp is not None:
        for k in ('before', 'after'):
            if sp.get(w(k)) is not None:
                d[k] = float(sp.get(w(k))) / 20
        if sp.get(w('line')) is not None:
            rule = sp.get(w('lineRule')) or 'auto'
            # UDF LineSpacing = satır yüksekliğine eklenen oran (1.15 -> 0.15)
            d['line'] = max(0.0, float(sp.get(w('line'))) / 240 - 1) \
                if rule == 'auto' else 0.0
    ind = ppr.find(w('ind'))
    if ind is not None:
        d['ind'] = _ind(ind)
    npr = ppr.find(w('numPr'))
    if npr is not None:
        nid = npr.find(w('numId'))
        lvl = npr.find(w('ilvl'))
        if nid is not None:
            d['numId'] = nid.get(w('val'))
        if lvl is not None:
            d['ilvl'] = lvl.get(w('val'))
    tabs = ppr.find(w('tabs'))
    if tabs is not None:
        d['tabs'] = [(float(t.get(w('pos')) or 0) / 20, t.get(w('val')) or 'left',
                      t.get(w('leader')) or 'none') for t in tabs.findall(w('tab'))
                     if t.get(w('pos')) is not None]
    if _on(ppr.find(w('pageBreakBefore'))):
        d['pb_before'] = True
    top = ppr.find(f'{w("pBdr")}/{w("top")}')
    if top is not None and (top.get(w('val')) or 'none') not in ('none', 'nil'):
        c = (top.get(w('color')) or '').upper()
        d['border_top'] = c if re.fullmatch(r'[0-9A-F]{6}', c) else '000000'
    return d


def _rpr_props(rpr, theme=None):
    """<w:rPr> -> {bold, italic, underline, size, color, font}"""
    d = {}
    if rpr is None:
        return d
    for tag, key in (('b', 'bold'), ('i', 'italic'), ('u', 'underline')):
        v = _on(rpr.find(w(tag)))
        if v is not None:
            d[key] = v
    sz = rpr.find(w('sz'))
    if sz is not None:
        try:
            d['size'] = int(sz.get(w('val'))) / 2
        except (TypeError, ValueError):
            pass
    col = rpr.find(w('color'))
    if col is not None:
        v = (col.get(w('val')) or '').strip().upper()
        d['color'] = v if re.fullmatch(r'[0-9A-F]{6}', v) and v != '000000' else None
    va = rpr.find(w('vertAlign'))
    if va is not None:
        d['super'] = va.get(w('val')) == 'superscript'
        d['sub'] = va.get(w('val')) == 'subscript'
    st = _on(rpr.find(w('strike')))
    ds = _on(rpr.find(w('dstrike')))
    if st is not None or ds is not None:
        d['strike'] = bool(st or ds)
    for tag in ('caps', 'smallCaps'):
        v = _on(rpr.find(w(tag)))
        if v is not None:
            d['caps'] = v or d.get('caps', False)
    if _on(rpr.find(w('vanish'))):
        d['hidden'] = True
    hl = rpr.find(w('highlight'))
    if hl is not None:
        d['bg'] = HIGHLIGHT.get(hl.get(w('val')))
    shd = rpr.find(w('shd'))
    if shd is not None and 'bg' not in d:
        f_ = (shd.get(w('fill')) or '').upper()
        d['bg'] = f_ if re.fullmatch(r'[0-9A-F]{6}', f_) and f_ != 'FFFFFF' else None
    rf = rpr.find(w('rFonts'))
    if rf is not None:
        f = rf.get(w('ascii')) or rf.get(w('hAnsi'))
        if not f and theme:
            th = rf.get(w('asciiTheme')) or rf.get(w('hAnsiTheme')) or ''
            f = theme.get('major' if th.startswith('major') else 'minor')
        if f:
            d['font'] = f
    return d


def _borders(el):
    out = {}
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV', 'start', 'end'):
        b = el.find(w(side))
        if b is not None:
            out[{'start': 'left', 'end': 'right'}.get(side, side)] = b.get(w('val')) or 'none'
    return out


def theme_fonts(z):
    try:
        t = z.read('word/theme/theme1.xml').decode('utf-8')
    except KeyError:
        return {}
    out = {}
    for k in ('minor', 'major'):
        m = re.search(rf'<a:{k}Font>\s*<a:latin typeface="([^"]*)"', t)
        if m:
            out[k] = m.group(1)
    return out


def style_map(styles_xml, theme):
    """styleId -> biçim/paragraf özellikleri (basedOn zinciri için)."""
    out = {}
    root = ET.fromstring(styles_xml)
    for st in root.findall(w('style')):
        sid = st.get(w('styleId'))
        d = {}
        d.update(_rpr_props(st.find(w('rPr')), theme))
        d.update(_ppr_props(st.find(w('pPr'))))
        tb = st.find(f'{w("tblPr")}/{w("tblBorders")}')
        if tb is not None:
            d['tblBorders'] = _borders(tb)
        bo = st.find(w('basedOn'))
        d['basedOn'] = bo.get(w('val')) if bo is not None else None
        out[sid] = d
        if st.get(w('type')) == 'paragraph' and st.get(w('default')) in ('1', 'true'):
            out['__default_p__'] = sid
    dd = {}
    rd = root.find(f'{w("docDefaults")}/{w("rPrDefault")}/{w("rPr")}')
    pd = root.find(f'{w("docDefaults")}/{w("pPrDefault")}/{w("pPr")}')
    dd.update(_rpr_props(rd, theme))
    dd.update(_ppr_props(pd))
    out['__docdefault__'] = dd
    return out


def resolve(sid, smap, seen=None):
    """Stil zincirini basedOn üzerinden çözer."""
    seen = seen or set()
    if not sid or sid in seen or sid not in smap:
        return {}
    seen.add(sid)
    d = dict(resolve(smap[sid].get('basedOn'), smap, seen))
    _merge(d, {k: v for k, v in smap[sid].items() if k != 'basedOn'})
    return d


ROMAN = [(1000, 'm'), (900, 'cm'), (500, 'd'), (400, 'cd'), (100, 'c'), (90, 'xc'),
         (50, 'l'), (40, 'xl'), (10, 'x'), (9, 'ix'), (5, 'v'), (4, 'iv'), (1, 'i')]


def _roman(n):
    out = ''
    for v, s in ROMAN:
        while n >= v:
            out += s
            n -= v
    return out


def numbering_map(z):
    """numId -> {ilvl: (numFmt, lvlText, start, ind)}; startOverride dikkate alınır."""
    try:
        root = ET.fromstring(z.read('word/numbering.xml'))
    except KeyError:
        return {}
    abst = {}
    for a in root.findall(w('abstractNum')):
        aid = a.get(w('abstractNumId'))
        lv = {}
        for l in a.findall(w('lvl')):
            i = l.get(w('ilvl'))
            f = l.find(w('numFmt'))
            x = l.find(w('lvlText'))
            s = l.find(w('start'))
            ind = _ind(l.find(f'{w("pPr")}/{w("ind")}'))
            ind = dict(ind, _bullet=(x.get(w('val')) if x is not None else ''))
            lv[i] = (f.get(w('val')) if f is not None else 'decimal',
                     x.get(w('val')) if x is not None else '%1.',
                     int(s.get(w('val'))) if s is not None else 1,
                     ind)
        abst[aid] = lv
    out = {}
    for nm in root.findall(w('num')):
        nid = nm.get(w('numId'))
        a = nm.find(w('abstractNumId'))
        lv = dict(abst.get(a.get(w('val')), {})) if a is not None else {}
        for ov in nm.findall(w('lvlOverride')):
            i = ov.get(w('ilvl'))
            so = ov.find(w('startOverride'))
            if so is not None and i in lv:
                f, x, _, ind = lv[i]
                lv[i] = (f, x, int(so.get(w('val'))), ind)
        out[nid] = lv
    return out


# Word numFmt -> Editör numara ailesi. Tür adı = NUMBER_TYPE_<aile>_<ek>; ek lvlText'ten gelir:
# "%1." DOT, "%1)" PARANTHESE, "(%1)" D_PARANTHESE, "%1-" TRE (21.09.2026: Editör sınıf sabitlerinden
# okundu, 15 türün hepsi motorla çizdirildi).
UDF_NUM = {
    'decimal':     'NUMBER',
    'upperLetter': 'CHAR_BIG',
    'lowerLetter': 'CHAR_SMALL',
    'lowerRoman':  'ROMAN_SMALL',
    'upperRoman':  'ROMAN_BIG',
}


def _num_suffix(tmpl, ilvl):
    t = (tmpl or '').strip()
    ph = f'%{int(ilvl) + 1}'
    if t.startswith('(') and t.endswith(ph + ')'):
        return 'D_PARANTHESE'
    if t.endswith(ph + ')'):
        return 'PARANTHESE'
    if t.endswith(ph + '-') or t.endswith(ph + ' -'):
        return 'TRE'
    return 'DOT'


def list_spec(numid, ilvl, nmap, counters, warn=None):
    """Word numaralandırmasını UDF liste özniteliklerine çevirir.

    Döner: (attrs, literal, ind). attrs UDF paragraf öznitelikleri; literal UDF'de karşılığı
    olmayan biçimler için düz metin etiket; ind numbering.xml'deki düzey girintisidir.

    Editör alt düzeylerde (ListLevel >= 2) paragrafın NumberType/BulletType'ına BAKMAZ; o düzeydeki
    paragrafın kendi SecListTypeLevel<N> özniteliğini, yoksa varsayılanı (a. / i. / (1) ...) çizer
    (21.09.2026'da motorla sınandı). Bu yüzden alt düzey paragraflara o öznitelik de yazılır.
    Liste 1'den başlamıyorsa ilk kaleme NumberSetted = başlangıç - 1 konur (Editör n+1'den sayar)."""
    lv = nmap.get(numid, {}).get(ilvl)
    if not lv:
        return '', None, {}
    fmt, tmpl, start, ind = lv
    if fmt == 'none':
        return '', None, ind
    lid, llv = int(numid), int(ilvl) + 1
    if fmt == 'bullet':
        ch = ord(ind.get('_bullet', '')[:1] or '\u2022')
        bt = 'ARROW' if ch in BULLET_ARROW else 'DIAMOND' if ch in BULLET_DIAMOND else \
            'TRIANGLE' if ch in BULLET_TRIANGLE else 'RECTANGLE' if ch in BULLET_RECT else \
            'RECTANGLE_D' if ch in BULLET_RECT_D else 'ELLIPSE'
        sec = f' SecListTypeLevel{llv}="BULLET_TYPE_{bt}"' if 2 <= llv <= 6 else ''
        return (f' Bulleted="true" BulletType="BULLET_TYPE_{bt}"{sec}'
                f' ListId="{lid}" ListLevel="{llv}"'), None, ind
    if fmt in UDF_NUM:
        typ = f'NUMBER_TYPE_{UDF_NUM[fmt]}_{_num_suffix(tmpl, ilvl)}'
        if warn is not None and len(re.findall(r'%\d', tmpl or '')) > 1:
            warn.add('çok basamaklı numaralandırma ("1.1." gibi) Editör\'de yok; yalnız son basamak numaralandı')
        sec = f' SecListTypeLevel{llv}="{typ}"' if 2 <= llv <= 6 else ''
        say = counters.setdefault(('say', numid), {})          # sığ kalem gelince derin düzey yeniden başlar
        for k in [k for k in say if k > llv]:
            del say[k]
        say[llv] = say.get(llv, (start or 1) - 1) + 1
        if warn is not None and fmt in ('upperLetter', 'lowerLetter') and say[llv] >= 4:
            warn.add('harfli listede Editör Türk alfabesiyle sayar (…c, ç, d…): Word\'deki 4. ve sonraki '
                     'harfler (d, e…) Editör\'de bir harf kayar')
        setted = ''
        key = ('ilk', numid, ilvl)
        if key not in counters:
            counters[key] = True
            if start and start > 1:
                setted = f' NumberSetted="{start - 1}"'
        return (f' Numbered="true" NumberType="{typ}"{sec}{setted}'
                f' ListId="{lid}" ListLevel="{llv}"'), None, ind
    # karşılığı olmayan biçim (ör. sıra sözcükleri) — düz metne düş
    key = (numid, ilvl)
    counters[key] = counters.get(key, start - 1) + 1
    n = counters[key]
    val = _roman(n).upper() if fmt == 'upperRoman' else _roman(n) if fmt == 'lowerRoman' else str(n)
    return '', re.sub(r'%\d', val, tmpl), ind


FMT_KEYS = ('bold', 'italic', 'underline', 'size', 'color', 'font',
            'super', 'sub', 'strike', 'caps', 'bg', 'hidden')


def part_rels(z, part):
    """'word/header1.xml' -> {rId: hedef yolu (word/ altında)}"""
    d, n = part.rsplit('/', 1)
    try:
        root = ET.fromstring(z.read(f'{d}/_rels/{n}.rels'))
    except KeyError:
        return {}
    out = {}
    for rel in root:
        if rel.get('TargetMode') == 'External':
            continue
        t = rel.get('Target') or ''
        # '../media/x.png', './header1.xml' gibi göreli hedefler normalleştirilir (bazı üreticiler yazıyor;
        # normalleştirilmeyince görsel/üstbilgi uyarısız düşüyordu — 18.10.2026 Argon denetimi)
        out[rel.get('Id')] = posixpath.normpath(t.lstrip('/') if t.startswith('/') else posixpath.join(d, t))
    return out


def _img_kind(data):
    """'png' | 'raster' (JPEG/GIF/BMP/TIFF) | 'pdf' | 'emf' | 'wmf' | None"""
    if data[:4] == b'\x89PNG':
        return 'png'
    if data[:3] in (b'\xff\xd8\xff', b'GIF') or data[:2] == b'BM' or data[:4] in (b'II*\x00', b'MM\x00*'):
        return 'raster'
    if data[:5] == b'%PDF-':
        return 'pdf'                              # Mac Word PDF'i (ör. Canva anteti) .wmf adıyla saklayabiliyor
    if data[:4] == b'\x01\x00\x00\x00' and data[40:44] == b' EMF':
        return 'emf'
    if data[:4] == b'\xd7\xcd\xc6\x9a' or data[:4] in (b'\x01\x00\x09\x00', b'\x02\x00\x09\x00'):
        return 'wmf'
    return None


PDF_PX = 1684                                     # vektör görsel çizim boyu (uzun kenar, piksel; A4'te ~144 dpi)
_NO_WINDOW = 0x08000000 if os.name == 'nt' else 0  # Windows: PowerShell konsol penceresi açılmasın

# Windows 10+ yerleşik PDF çizicisi (Windows.Data.Pdf). Mac'te sınanamadı; başarısız olursa görsel atlanır + uyarı.
_PS_PDF = r"""
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$m=[System.WindowsRuntimeSystemExtensions].GetMethods()
$asOp=($m|?{$_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'})[0]
$asAct=($m|?{$_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction'})[0]
function W($o,$t){$k=$asOp.MakeGenericMethod($t).Invoke($null,@($o));$k.Wait(-1)|Out-Null;$k.Result}
function A($o){$k=$asAct.Invoke($null,@($o));$k.Wait(-1)|Out-Null}
[Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime]|Out-Null
[Windows.Data.Pdf.PdfDocument,Windows.Data.Pdf,ContentType=WindowsRuntime]|Out-Null
[Windows.Storage.Streams.IRandomAccessStream,Windows.Storage.Streams,ContentType=WindowsRuntime]|Out-Null
$f=W ([Windows.Storage.StorageFile]::GetFileFromPathAsync('%SRC%')) ([Windows.Storage.StorageFile])
$d=W ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($f)) ([Windows.Data.Pdf.PdfDocument])
$pg=$d.GetPage(0)
$dir=W ([Windows.Storage.StorageFolder]::GetFolderFromPathAsync('%DIR%')) ([Windows.Storage.StorageFolder])
$o=W ($dir.CreateFileAsync('out.png',[Windows.Storage.CreationCollisionOption]::ReplaceExisting)) ([Windows.Storage.StorageFile])
$st=W ($o.OpenAsync([Windows.Storage.FileAccessMode]::ReadWrite)) ([Windows.Storage.Streams.IRandomAccessStream])
$op=New-Object Windows.Data.Pdf.PdfPageRenderOptions
if($pg.Size.Height -ge $pg.Size.Width){$op.DestinationHeight=%PX%}else{$op.DestinationWidth=%PX%}
A ($pg.RenderToStreamAsync($st,$op))
$st.Dispose()
"""


def _convert_png(data):
    """JPEG/GIF/BMP/TIFF/PDF/EMF/WMF -> PNG. Sırayla: Pillow (EMF/WMF yalnız Windows'ta), macOS sips,
    Windows PowerShell (System.Drawing; PDF için Windows.Data.Pdf)."""
    kind = _img_kind(data)
    if kind != 'pdf':
        try:
            from PIL import Image
            im = Image.open(io.BytesIO(data))
            if kind in ('emf', 'wmf'):
                im.load(dpi=144)                  # vektör: çizim çözünürlüğü (Pillow yalnız Windows'ta çizer)
            im = im.convert('RGBA' if im.mode in ('RGBA', 'LA', 'P') or kind in ('emf', 'wmf') else 'RGB')
            buf = io.BytesIO()
            im.save(buf, 'PNG')
            return buf.getvalue()
        except Exception:
            pass
    ext = {'pdf': '.pdf', 'emf': '.emf', 'wmf': '.wmf'}.get(kind) or (
        '.jpg' if data[:3] == b'\xff\xd8\xff' else '.gif' if data[:3] == b'GIF' else '.bmp' if data[:2] == b'BM' else '.tif')
    d = tempfile.mkdtemp()
    try:
        src, dst = os.path.join(d, 'in' + ext), os.path.join(d, 'out.png')
        with open(src, 'wb') as f:
            f.write(data)
        if shutil.which('sips'):
            cmd = ['sips', '-s', 'format', 'png'] + (['-Z', str(PDF_PX)] if kind == 'pdf' else []) + [src, '--out', dst]
            subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=60)
        elif shutil.which('powershell'):
            if kind == 'pdf':
                ps = _PS_PDF.replace('%SRC%', src.replace("'", "''")).replace('%DIR%', d.replace("'", "''")) \
                    .replace('%PX%', str(PDF_PX))
            else:
                ps = ("Add-Type -AssemblyName System.Drawing; "
                      "$i=[System.Drawing.Image]::FromFile($env:UDF_SRC); "
                      "$i.Save($env:UDF_DST,[System.Drawing.Imaging.ImageFormat]::Png); $i.Dispose()")
            subprocess.run(['powershell', '-NoProfile', '-NonInteractive', '-Command', ps],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=60,
                           creationflags=_NO_WINDOW, env=dict(os.environ, UDF_SRC=src, UDF_DST=dst))
        if os.path.exists(dst):
            with open(dst, 'rb') as f:
                out = f.read()
            if out[:4] == b'\x89PNG':
                return out
    except Exception:
        pass
    finally:
        shutil.rmtree(d, ignore_errors=True)
    return None


GORSEL_DPI = 200                                  # sayfadaki boyutuna göre bu çözünürlüğe indirilir (baskıda fark görülmez)
UYAP_SINIR_MB = 10                                # UYAP'a yüklenen dosya sınırı (meslektaşların bildirdiği)
KUCULTULEN = [0, 0]                               # (küçültülen görsel sayısı, kazanılan bayt) — cevir() sıfırlar


def kucult(png, w_pt, h_pt):
    """PNG'yi sayfada kaplayacağı boyuta (GORSEL_DPI) indirir; gerekmiyorsa ya da Pillow yoksa aynen döner.
    UDF yalnız PNG taşıdığından kameradan gelen 4000 piksellik fotoğraf JPEG'deki boyutunun birkaç katına
    çıkıyordu (06.10.2026: 6 fotoğraflı 13,6 MB'lık belge 51,6 MB UDF oldu, UYAP'a yüklenemedi)."""
    if not w_pt or not h_pt:
        return png
    try:
        from PIL import Image
        im = Image.open(io.BytesIO(png))
        hedef = int(w_pt / 72.0 * GORSEL_DPI + 0.5)
        if im.width <= hedef * 1.15:              # zaten yeterince küçük: yeniden kodlama yok
            return png
        boy = max(1, int(im.height * hedef / im.width + 0.5))
        if im.mode not in ('RGB', 'RGBA', 'L', 'LA'):
            im = im.convert('RGBA' if 'transparency' in im.info or im.mode in ('P', 'PA') else 'RGB')
        im = im.resize((hedef, boy), Image.LANCZOS)
        buf = io.BytesIO()
        im.save(buf, 'PNG', optimize=True)
        out = buf.getvalue()
        if len(out) < len(png):
            KUCULTULEN[0] += 1
            KUCULTULEN[1] += len(png) - len(out)
            return out
    except Exception:
        pass
    return png


def to_png(data, name, warn):
    """UDF yalnız PNG taşır (derlemdeki 417 görselin tamamı PNG)."""
    kind = _img_kind(data)
    if kind == 'png':
        return data
    if kind is None:
        warn.add(f'desteklenmeyen görsel biçimi atlandı: {name}')
        return None
    out = _convert_png(data)
    if out is None:
        if kind == 'raster':
            warn.add(f'görsel PNG\'ye çevrilemedi, atlandı (Pillow kurulu değil): {name}')
        else:
            warn.add(f'desteklenmeyen görsel biçimi atlandı ({kind.upper()} çizilemedi): {name}')
    return out


def _len_pt(v):
    m = re.fullmatch(r'\s*(-?[\d.]+)\s*(pt|in|cm|mm|px)?\s*', v or '')      # VML konumu eksi olabilir
    if not m:
        return None
    k = {'pt': 1.0, 'in': 72.0, 'cm': 28.3465, 'mm': 2.83465, 'px': 0.75, None: 1.0}[m.group(2)]
    return float(m.group(1)) * k


def _anchor_info(node, ctx, w_pt):
    """Serbest konumlu (wp:anchor) çizimin kaydırma türü ve yatay tarafı."""
    an = node.find(f'{{{WP}}}anchor')
    if an is None:
        return None
    wrap = next((ch.tag.split('}')[1] for ch in an if ch.tag.split('}')[1].startswith('wrap')), 'wrapNone')
    side = None
    ph = an.find(f'{{{WP}}}positionH')
    if ph is not None:
        al, po = ph.find(f'{{{WP}}}align'), ph.find(f'{{{WP}}}posOffset')
        if al is not None and al.text:
            side = {'left': 'left', 'inside': 'left', 'right': 'right', 'outside': 'right',
                    'center': 'center'}.get(al.text.strip())
        elif po is not None and po.text:
            try:
                x = float(po.text) / 12700
            except ValueError:
                x = 0.0
            if ph.get('relativeFrom') == 'page':
                x -= ctx.get('margin_left', 70.0)
            tw = ctx.get('text_w', 450.0)
            mid = x + w_pt / 2
            side = 'left' if mid < tw * 0.4 else 'right' if mid > tw * 0.6 else 'center'
    try:
        gap = max(float(an.get('distL') or 0), float(an.get('distR') or 0)) / 12700
    except ValueError:
        gap = 0.0
    return {'wrap': wrap, 'side': side or 'left', 'gap': min(max(gap or 9.0, 4.0), 24.0)}


def _page_pos(an, ctx, w_pt, h_pt):
    """Serbest konumlu çizimin sayfadaki sol-üst köşesi (punto). Paragrafa/satıra bağlı dikey konumda
    paragrafın üstü, üstbilgi için sayfa kenarından üstbilgi uzaklığı (pgMar w:header) alınır."""
    pw, ph = ctx.get('page', (595.3, 841.9))
    ml, mr, mt, mb = ctx.get('margins', (70.9, 70.9, 70.9, 70.9))

    def axis(tag, size, page, m0, m1, para0):
        el = an.find(f'{{{WP}}}{tag}')
        if el is None:
            return 0.0
        rel = el.get('relativeFrom') or 'page'
        if rel == 'page':
            lo, span = 0.0, page
        elif rel in ('leftMargin', 'insideMargin', 'topMargin'):
            lo, span = 0.0, m0
        elif rel in ('rightMargin', 'outsideMargin', 'bottomMargin'):
            lo, span = page - m1, m1
        elif rel in ('paragraph', 'line'):
            lo, span = para0, page - para0 - m1
        else:                                     # margin, column, character, text
            lo, span = m0, page - m0 - m1
        al, po = el.find(f'{{{WP}}}align'), el.find(f'{{{WP}}}posOffset')
        if al is not None and al.text:
            a = al.text.strip()
            return lo + ((span - size) / 2 if a == 'center' else span - size if a in ('right', 'outside', 'bottom') else 0.0)
        try:
            return lo + float(po.text) / 12700 if po is not None and po.text else lo
        except ValueError:
            return lo
    return (axis('positionH', w_pt, pw, ml, mr, ml),
            axis('positionV', h_pt, ph, mt, mb, ctx.get('hdr_dist', 35.4)))


def _vml_style(sh):
    return {k.strip().lower(): v.strip() for k, v in
            (kv.split(':', 1) for kv in (sh.get('style') or '').split(';') if ':' in kv)}


def _vml_pos(sh, ctx, w_pt, h_pt):
    """Eski tip (VML) konumlu görselin sayfadaki sol-üst köşesi. Word'ün "filigran resmi" bu biçimdedir:
    mso-position-horizontal/vertical (left/center/right, top/center/bottom) + *-relative (page/margin/text)."""
    st = _vml_style(sh)
    pw, ph = ctx.get('page', (595.3, 841.9))
    ml, mr, mt, mb = ctx.get('margins', (70.9, 70.9, 70.9, 70.9))

    def axis(al_key, rel_key, off_key, size, page, m0, m1, para0):
        rel = st.get(rel_key, 'text')
        if rel == 'page':
            lo, span = 0.0, page
        elif rel in ('left-margin-area', 'top-margin-area', 'inner-margin-area'):
            lo, span = 0.0, m0
        elif rel in ('right-margin-area', 'bottom-margin-area', 'outer-margin-area'):
            lo, span = page - m1, m1
        elif rel in ('line', 'paragraph') or (rel == 'text' and al_key.endswith('vertical')):
            lo, span = para0, page - para0 - m1
        else:                                     # margin, text (yatay), char
            lo, span = m0, page - m0 - m1
        a = st.get(al_key, 'absolute')
        if a in ('center', 'left', 'right', 'top', 'bottom', 'inside', 'outside'):
            return lo + ((span - size) / 2 if a == 'center' else span - size if a in ('right', 'bottom', 'outside') else 0.0)
        return lo + (_len_pt(st.get(off_key)) or 0.0)
    return (axis('mso-position-horizontal', 'mso-position-horizontal-relative', 'margin-left', w_pt, pw, ml, mr, ml),
            axis('mso-position-vertical', 'mso-position-vertical-relative', 'margin-top', h_pt, ph, mt, mb,
                 ctx.get('hdr_dist', 35.4)))


def _image_run(node, ctx, warn):
    """<w:drawing> / <w:pict> -> {'image': {'data', 'w', 'h'}, ['anchor': ...]} ya da None."""
    rid, w_pt, h_pt, anchor = None, None, None, None
    if node.tag == w('drawing'):
        blip = node.find(f'.//{{{A}}}blip')
        if blip is None:
            return None
        rid = blip.get(f'{{{R}}}embed') or blip.get(f'{{{R}}}link')
        ext = node.find(f'.//{{{WP}}}extent')
        if ext is not None and ext.get('cx'):
            w_pt, h_pt = float(ext.get('cx')) / 12700, float(ext.get('cy')) / 12700
        anchor = _anchor_info(node, ctx, w_pt or 0.0)
    else:
        im = node.find(f'.//{{{V}}}imagedata')
        if im is None:
            return None
        rid = im.get(f'{{{R}}}id')
        vshape = None
        for sh in node.iter(f'{{{V}}}shape'):
            st = _vml_style(sh)
            w_pt, h_pt = _len_pt(st.get('width')), _len_pt(st.get('height'))
            vshape = sh
            break
        if vshape is not None and vshape is ctx.get('bg_pick'):
            anchor = {'wrap': 'wrapNone', 'side': 'center', 'gap': 9.0}     # aşağıda arka plana yönlenir
    target = ctx['rels'].get(rid)
    if not target:
        return None
    try:
        data = ctx['zip'].read(target)
    except KeyError:
        if warn is not None:
            warn.add('belgede adı geçen bir görsel dosyanın içinde bulunamadı; atlandı')
        return None
    data = to_png(data, target, warn)
    if data is None:
        return None
    if not w_pt or not h_pt:
        px = struct.unpack('>II', data[16:24])
        w_pt, h_pt = px[0] * 0.75, px[1] * 0.75               # 96 dpi varsayımı
    data = kucult(data, w_pt, h_pt)
    out = {'image': {'data': base64.encodebytes(data).decode('ascii'), 'w': w_pt, 'h': h_pt}}
    if anchor:
        out['anchor'] = anchor
        an = node.find(f'{{{WP}}}anchor')
        if an is None and node.tag != w('drawing'):
            an = next(node.iter(f'{{{V}}}shape'), None)
        if an is not None and an is ctx.get('bg_pick'):
            # Üstbilgide metni kaydırmayan (wrapNone) serbest görsel = antet / sayfa zemini. Satır içine
            # alınırsa üstbilgi sayfayı doldurur (30.09.2026: 5 sayfalık antetli dilekçe 27 sayfa oldu).
            # Editör'de karşılığı sayfa arka planıdır (bgImage): kenar paylarıyla verilen dikdörtgene gerilir.
            x, y = (_vml_pos(an, ctx, w_pt, h_pt) if an.tag == f'{{{V}}}shape' else _page_pos(an, ctx, w_pt, h_pt))
            ctx['bg_found'] = {'data': ''.join(out['image']['data'].split()), 'left': x, 'top': y, 'w': w_pt, 'h': h_pt}
            return None
    return out


def _run_nodes(r):
    """Run'ın çocukları; mc:AlternateContent tek dala indirgenir (metin/görsel iki kez gelmesin)."""
    for node in r:
        if node.tag == f'{{{MC}}}AlternateContent':
            ch, fb = node.find(f'{{{MC}}}Choice'), node.find(f'{{{MC}}}Fallback')
            ok = ch is not None and (ch.find(f'.//{{{A}}}blip') is not None
                                     or ch.find(f'.//{w("txbxContent")}') is not None)
            pick = ch if ok else (fb if fb is not None else ch)
            if pick is not None:
                for sub in pick:
                    yield sub
        else:
            yield node


PAGE_FIELDS = ('PAGE', 'NUMPAGES', 'SECTIONPAGES')


def _is_page_field(instr):
    """Alan yönergesinin ilk sözcüğü sayfa numarası alanı mı? ('PAGEREF' ve 'TOC' DEĞİL:
    içindekiler tablosunun sayfa numaraları metindir, atılırsa metin kaybolur — 21.09.2026.)"""
    tok = (instr or '').strip().split()
    return bool(tok) and tok[0].upper() in PAGE_FIELDS


def _field_instrs(el):
    """Öğedeki alan yönergeleri: karmaşık alanda begin..separate arasındaki bütün instrText parçaları
    birleştirilir (Word yönergeyi " PA" + "GE " diye bölebiliyor), basit alanda w:instr."""
    out, stack = [], []                           # stack: açık alanların yönerge parçaları (None = okundu)
    for n in el.iter():
        if n.tag == w('fldChar'):
            t = n.get(w('fldCharType'))
            if t == 'begin':
                stack.append([])
            elif stack:
                if stack[-1] is not None:         # ilk separate ya da end: yönerge tamamlandı
                    out.append(''.join(stack[-1]))
                    stack[-1] = None
                if t == 'end':
                    stack.pop()
        elif n.tag == w('instrText'):
            if not stack:                         # alan önceki paragrafta başlamış
                stack.append([])
            if stack[-1] is not None:
                stack[-1].append(n.text or '')
        elif n.tag == w('fldSimple'):
            out.append(n.get(w('instr')) or '')
    out += [''.join(x) for x in stack if x]
    return out


_PAGE_WORDS = {'sayfa', 'page', 'of', 'sf', 's', 'syf', 'toplam', 'pg', 'p'}


def _page_pattern_only(rest):
    """Sayfa alanı çıkınca kalan metin yalnız "Sayfa  / ", "- -", "Page  of " gibi bir kalıp mı?"""
    toks = re.findall(r'[^\W\d_]+', rest.lower())
    return len(rest.strip()) <= 24 and all(t in _PAGE_WORDS for t in toks)


def _iter_runs(p):
    """Paragrafın kendi run'ları. Atlananlar: metin kutusu içindekiler (ayrı paragraf
    olur), taşıma kaynağı (w:moveFrom), PAGE/NUMPAGES alan önbelleği."""
    skip = set()
    for tag in ('txbxContent', 'moveFrom'):
        for el in p.iter(w(tag)):
            skip.update(id(r) for r in el.iter(w('r')))
    for fs in p.iter(w('fldSimple')):
        if _is_page_field(fs.get(w('instr'))):
            skip.update(id(r) for r in fs.iter(w('r')))
    stack = []                                    # iç içe alanlar (içindekiler > PAGEREF): düzey başına gizleme
    instr = []                                    # her düzeyin şimdiye kadarki yönergesi
    for r in p.iter(w('r')):
        if id(r) in skip:
            continue
        fc = r.find(w('fldChar'))
        if fc is not None:
            t = fc.get(w('fldCharType'))
            if t == 'begin':
                stack.append(False)
                instr.append('')
            elif t == 'end' and stack:
                stack.pop()
                instr.pop()
            continue
        it = r.find(w('instrText'))
        if it is not None:
            if stack:
                instr[-1] += it.text or ''
                if _is_page_field(instr[-1]):
                    stack[-1] = True
            continue
        if any(stack):
            continue
        yield r


def _empty_para(in_cell=False):
    return {'type': 'p', 'align': 0, 'runs': [], 'indent': {}, 'list': '', 'literal': False,
            'in_cell': in_cell, 'before': None, 'after': None, 'line': 0.0 if not in_cell else None,
            'tabs': None, 'border_top': None, 'avail': None}


def parse_paragraph(p, ctx, in_cell=False, extra=None, warn=None):
    """<w:p> -> blok listesi: paragraflar ve {'type': 'pb'} sayfa sonu işaretleri.
    w:br satır sonu paragrafı böler; metin kutusu içeriği paragrafın ardına eklenir."""
    warn = warn if warn is not None else set()
    smap, nmap, counters = ctx['smap'], ctx['nmap'], ctx['counters']
    ppr = p.find(w('pPr'))
    own = _ppr_props(ppr)
    ps = ppr.find(w('pStyle')) if ppr is not None else None
    pstyle = ps.get(w('val')) if ps is not None else smap.get('__default_p__')
    base = _merge(dict(smap.get('__docdefault__', {})), resolve(pstyle, smap))
    if extra:
        _merge(base, extra)                       # tablo stilinin pPr'si
    eff = _merge(dict(base), own)

    align = eff.get('jc')
    if align is None:
        align = 0                                 # Word varsayılanı: sola yaslı

    listattr, label, lvl_ind = '', None, {}
    nid = eff.get('numId')
    if nid and nid != '0':
        listattr, label, lvl_ind = list_spec(nid, eff.get('ilvl', '0'), nmap, counters, warn)
    lvl_ind = {k: v for k, v in lvl_ind.items() if not k.startswith('_')}
    # Öncelik: paragrafın kendi w:ind > numbering düzeyi > stil
    indent = _merge(_merge(dict(base.get('ind', {})), lvl_ind), own.get('ind', {}))

    lines = [[]]                                  # öğe: run listesi ya da 'PB'
    boxes = []
    float_img, block_imgs = None, []
    has_pagefld = any(_is_page_field(x) for x in _field_instrs(p))
    for r in _iter_runs(p):
        rpr = r.find(w('rPr'))
        fmt = {k: base.get(k) for k in FMT_KEYS if base.get(k)}
        if rpr is not None:
            rs = rpr.find(w('rStyle'))
            if rs is not None:
                for k, v in resolve(rs.get(w('val')), smap).items():
                    if k in FMT_KEYS:
                        fmt[k] = v
            for k, v in _rpr_props(rpr, ctx.get('theme')).items():
                if v:
                    fmt[k] = v
                else:
                    fmt.pop(k, None)
        if fmt.pop('hidden', None):
            continue                              # gizli metin (w:vanish)
        caps = fmt.pop('caps', None)
        buf = ['']

        def flush():
            if buf[0]:
                t = clean(buf[0])
                t = tr_upper(t) if caps else t
                f2 = fmt
                if fmt.get('sub') and t.translate(SUBSCRIPT) != t and \
                        all(ch.isspace() or ch.translate(SUBSCRIPT) != ch for ch in t):
                    t, f2 = t.translate(SUBSCRIPT), {k: v for k, v in fmt.items() if k != 'sub'}
                lines[-1].append((t, f2))
                buf[0] = ''

        for node in _run_nodes(r):
            tag = node.tag.split('}')[1]
            if tag == 't':
                buf[0] += node.text or ''
            elif tag == 'tab':
                buf[0] += '\t'
            elif tag == 'noBreakHyphen':
                buf[0] += '-'
            elif tag == 'sym':
                font = (node.get(w('font')) or '').lower()
                try:
                    code = int(node.get(w('char')) or '0', 16)
                except ValueError:
                    code = 0
                ch = SYM.get((font, code)) or SYM.get((font, code | 0xF000))
                if ch is None:
                    ch = '□'
                    warn.add(f'simgenin Unicode karşılığı bilinmiyor ({font} {code:04X}); □ yazıldı')
                buf[0] += ch
            elif tag in ('br', 'cr'):
                kind = node.get(w('type'))
                flush()
                if kind == 'page':
                    if in_cell:
                        warn.add('tablo hücresindeki sayfa sonu atlandı')
                    else:
                        lines.append('PB')
                        lines.append([])
                elif kind != 'column':
                    lines.append([])
            elif tag in ('footnoteReference', 'endnoteReference'):
                kind = 'footnote' if tag.startswith('foot') else 'endnote'
                el = ctx.get('noteparts', {}).get(kind, {}).get(node.get(w('id')))
                if el is not None:
                    n = sum(1 for k, _, _ in ctx['notes'] if k == kind) + 1
                    lab = str(n) if kind == 'footnote' else _roman(n)
                    ctx['notes'].append((kind, el, lab))
                    flush()
                    lines[-1].append((lab, dict(fmt, super=True)))
            elif tag in ('footnoteRef', 'endnoteRef'):
                if ctx.get('cur_note'):
                    flush()
                    lines[-1].append((ctx['cur_note'], dict(fmt, super=True)))
            elif tag in ('drawing', 'pict'):
                tb = node.find(f'.//{w("txbxContent")}')
                if tb is not None:
                    boxes.extend(tb.findall(w('p')))
                    warn.add('metin kutusu düz paragrafa çevrildi (kutunun konumu ve çerçevesi taşınmaz)')
                    continue
                img = _image_run(node, ctx, warn)
                if img:
                    an = img.pop('anchor', None)
                    layout = an and not in_cell and ctx.get('body')
                    if layout and ctx.get('yan_gorsel') and float_img is None and \
                            an['side'] in ('left', 'right') and \
                            an['wrap'] in ('wrapSquare', 'wrapTight', 'wrapThrough'):
                        float_img = dict(img['image'], side=an['side'], gap=an['gap'])
                    elif layout and an['wrap'] == 'wrapTopAndBottom':
                        block_imgs.append((img, an['side']))       # üst-alt kaydırma: kendi satırında
                    else:
                        if an:
                            warn.add('serbest konumlu görsel satır içine alındı (sayfadaki konum ve metin '
                                     'kaydırma taşınmaz)')
                        flush()
                        lines[-1].append((IMG_CHAR, img))
                elif node.find(f'.//{{{A}}}blip') is None and node.find(f'.//{{{V}}}imagedata') is None:
                    warn.add('şekil/grafik/SmartArt taşınmadı')
        flush()

    has_pb = 'PB' in lines
    segs = [x for x in lines if x == 'PB' or x or not has_pb]      # PB komşusu boş parçaları at
    out = []
    if eff.get('pb_before') and not in_cell:
        out.append({'type': 'pb'})
    plist = [x for x in segs if x != 'PB']
    for runs in segs:
        if runs == 'PB':
            out.append({'type': 'pb'})
            continue
        li = plist.index(runs) if runs in plist else 0
        first, last = runs is plist[0], runs is plist[-1]
        merged = []
        for txt, fmt in runs:                     # aynı biçimdeki bitişik parçaları birleştir
            if merged and merged[-1][1] == fmt and 'image' not in fmt:
                merged[-1][0] += txt
            else:
                merged.append([txt, dict(fmt)])
        ind = dict(indent)
        d = {'type': 'p', 'align': align, 'runs': merged, 'indent': ind,
             'list': listattr if first else '', 'literal': False, 'in_cell': in_cell,
             'before': eff.get('before') if first else None,
             'after': eff.get('after') if last else None,
             'line': eff.get('line'), 'tabs': eff.get('tabs'),
             'border_top': own.get('border_top') if first else None,
             'pagefld': has_pagefld, 'text_w': ctx.get('text_w'), 'base_size': base.get('size'),
             'avail': max(60.0, ctx.get('text_w', 450.0) - (ind.get('left') or 0) - (ind.get('right') or 0))}
        if first and label:
            d['literal'] = True
            if merged:
                merged.insert(0, [label + '\t', {k: v for k, v in merged[0][1].items() if k != 'image'}])
            else:
                merged.append([label + '\t', {}])
        if not first:
            # Word'de satır sonundan sonraki metin devam satırıdır: "left" hizasında başlar.
            # Ayrı paragrafa dönüştüğü için asılı/ilk satır girintisi kaldırılır (LeftIndent = left).
            d['indent'].pop('firstLine', None)
            d['indent'].pop('hanging', None)
        out.append(d)
    firstp = next((x for x in out if x['type'] == 'p'), None)
    if float_img is not None:
        if firstp is None:
            firstp = _empty_para()
            out.append(firstp)
        firstp['float'] = float_img
    for img, side in reversed(block_imgs):
        ip = _empty_para()
        ip.update(align={'left': 0, 'center': 1, 'right': 2}.get(side, 1), runs=[[IMG_CHAR, img]],
                  avail=ctx.get('text_w'), after=6.0)
        out.insert(out.index(firstp) if firstp in out else 0, ip)
    if ppr is not None and ppr.find(w('sectPr')) is not None and not in_cell:
        t = ppr.find(f'{w("sectPr")}/{w("type")}')
        if t is None or t.get(w('val')) not in ('continuous',):
            out.append({'type': 'pb'})            # bölüm sonu (sonraki sayfa)
    for bp in boxes:
        out.extend(b for b in parse_paragraph(bp, ctx, in_cell=in_cell, warn=warn) if b['type'] == 'p')
    return out


def iter_blocks(container):
    """Belge sırasıyla <w:p> ve <w:tbl> öğelerini verir (sdt vb. sarmalayıcılara iner)."""
    for ch in container:
        tag = ch.tag.split('}')[1]
        if tag in ('p', 'tbl'):
            yield ch
        elif tag in ('sectPr', 'pPr', 'tblPr', 'tblGrid', 'trPr', 'tcPr'):
            continue
        else:
            yield from iter_blocks(ch)


def cell_paragraphs(tc):
    """Hücrenin paragrafları; iç içe tablo varsa düzleşerek sıraya girer."""
    for b in iter_blocks(tc):
        if b.tag == w('p'):
            yield b
        else:
            for tr in b.findall(w('tr')):
                for c in tr.findall(w('tc')):
                    yield from cell_paragraphs(c)


def _lum(hex6):
    r, g, b = (int(hex6[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def _border_kind(borders):
    vis = lambda s: borders.get(s) not in (None, 'none', 'nil')
    if vis('insideH') or vis('insideV'):
        return 'borderCell'
    if any(vis(s) for s in ('top', 'left', 'bottom', 'right')):
        return 'borderTable'
    return 'borderNone'


# ---- metin ölçümü (tablo satır yüksekliği / dolgu bandı tahmini için) ----
# Times New Roman genişlikleri, 1/1000 em (AFM). Diğer yazı tipleri katsayıyla yaklaşıklanır.
_TNR = {' ': 250, '!': 333, '"': 408, '#': 500, '$': 500, '%': 833, '&': 778, "'": 180, '(': 333, ')': 333,
        '*': 500, '+': 564, ',': 250, '-': 333, '.': 250, '/': 278, ':': 278, ';': 278, '<': 564, '=': 564,
        '>': 564, '?': 444, '@': 921, '[': 333, ']': 333, '_': 500, '—': 1000, '–': 500, '’': 333, '‘': 333,
        '“': 444, '”': 444, '•': 350, '₺': 500, '€': 500, '§': 500, '°': 400, ' ': 250, '\t': 1000,
        'a': 444, 'b': 500, 'c': 444, 'd': 500, 'e': 444, 'f': 333, 'g': 500, 'h': 500, 'i': 278, 'j': 278,
        'k': 500, 'l': 278, 'm': 778, 'n': 500, 'o': 500, 'p': 500, 'q': 500, 'r': 333, 's': 389, 't': 278,
        'u': 500, 'v': 500, 'w': 722, 'x': 500, 'y': 500, 'z': 444, 'ç': 444, 'ğ': 500, 'ı': 278, 'ö': 500,
        'ş': 389, 'ü': 500, 'â': 444, 'î': 278, 'û': 500,
        'A': 722, 'B': 667, 'C': 667, 'D': 722, 'E': 611, 'F': 556, 'G': 722, 'H': 722, 'I': 333, 'J': 389,
        'K': 722, 'L': 611, 'M': 889, 'N': 722, 'O': 722, 'P': 556, 'Q': 722, 'R': 667, 'S': 556, 'T': 611,
        'U': 722, 'V': 722, 'W': 944, 'X': 722, 'Y': 722, 'Z': 611, 'Ç': 667, 'Ğ': 722, 'İ': 333, 'Ö': 722,
        'Ş': 556, 'Ü': 722, 'Â': 722, 'Î': 333, 'Û': 722}
_FONT_K = (('narrow', 0.86), ('arial', 1.08), ('helvet', 1.08), ('calibri', 0.98), ('verdana', 1.22),
           ('tahoma', 1.10), ('cambria', 1.05), ('courier', 1.20), ('garamond', 0.92), ('georgia', 1.10))
LINE_K = 1.15                                     # Editör'de satır yüksekliği ≈ 1,15 x punto (ölçüldü)


def text_width(t, fmt):
    size = fmt.get('size') or 12
    font = (fmt.get('font') or '').lower()
    k = next((v for key, v in _FONT_K if key in font), 1.0)
    wd = sum(_TNR.get(ch, 500 if ch.isdigit() else 560) for ch in t)
    return wd * size / 1000.0 * k * (1.05 if fmt.get('bold') else 1.0)


def para_metrics(p, avail):
    """(tahmini yükseklik, satır sayısı, metin genişliği) — punto."""
    width, size, img_h = 0.0, 0.0, 0.0
    for t, f in p['runs']:
        if 'image' in f:
            iw, ih = f['image']['w'], f['image']['h']
            if iw > avail > 0:
                ih = ih * avail / iw
            img_h = max(img_h, ih)
        else:
            width += text_width(t, f)
            size = max(size, f.get('size') or 12)
    size = size or p.get('base_size') or 12
    lines = max(1, int(-(-width // max(avail, 10.0)))) if width > 0 else 1
    if p.get('nowrap_band'):
        lines = 1                                 # dolgu bandı tek satıra sığacak kadar doldurulur
    h = max(lines * size * LINE_K * (1 + (p.get('line') if p.get('line') is not None else 0.15)), img_h)
    return h + (p.get('before') or 0) + (p.get('after') or 0), lines, width


def cell_height(cell):
    if cell.get('table'):
        return sum(cell['table']['row_h']) + 2 * len(cell['table']['row_h'])
    return sum(para_metrics(p, p.get('avail') or 100)[0] for p in cell['paras'])


def _cell_margins(tcpr, tblmar):
    """Hücre iç boşlukları (üst, sol, alt, sağ) punto: tcMar > tblCellMar > Word varsayılanı."""
    out = dict(tblmar)
    mar = tcpr.find(w('tcMar')) if tcpr is not None else None
    if mar is not None:
        for side, alt in (('top', None), ('left', 'start'), ('bottom', None), ('right', 'end')):
            el = mar.find(w(side))
            if el is None and alt:
                el = mar.find(w(alt))
            if el is not None and el.get(w('w')):
                out[side] = float(el.get(w('w'))) / 20
    return out


def _pad_fill(cell, mode):
    """Hücre dolgusu UDF'de yok (Editör'ün hücre görünümü yalnız kenarlık/genişlik tanır; yazı
    düzeyinde background çalışır). 'bant': yazıya hücre rengi zemin verilir ve tek satırlık
    paragraflar aynı zeminli bölünmez boşlukla hücre genişliğine uzatılır (20.09.2026 önizleme)."""
    fill = cell.get('fill')
    if not fill or mode == 'yok':
        return False
    dark = _lum(fill) < 0.5
    used = False
    for p in cell['paras']:
        texts = [r for r in p['runs'] if 'image' not in r[1]]
        if mode == 'yazi':                        # yalnız okunabilirlik: koyu zemin + açık yazı
            for r in texts:
                c_ = r[1].get('color')
                if dark and c_ and _lum(c_) > 0.7 and not r[1].get('bg'):
                    r[1]['bg'] = fill
                    used = True
            continue
        for r in texts:
            if not r[1].get('bg'):
                r[1]['bg'] = fill
        avail = p.get('avail') or 100
        _, lines, width = para_metrics(p, avail)
        if lines > 1 or width > 0.92 * avail or any('image' in r[1] for r in p['runs']):
            used = used or bool(texts)
            continue                              # sarılan paragraf: yalnız yazının arkası boyanır
        base = dict(texts[0][1]) if texts else {'size': p.get('base_size') or 12}
        padf = {k: v for k, v in base.items() if k in ('size', 'font')}
        padf['bg'] = fill
        nb = text_width(' ', padf) or 3.0
        tnr = not base.get('font') or 'times' in base['font'].lower()
        room = max(0.0, (0.94 if tnr else 0.86) * (avail + p['cell_pad'][0] + p['cell_pad'][1] - 2) - width)
        # (TNR genişlikleri AFM'den bilinir; diğer yazı tiplerinde tahmin kaba olduğundan pay bırakılır)
        n = int(room / nb)
        if n < 2:
            continue
        edge = min(2, n // 2)                     # yazı bandın kenarına yapışmasın
        if p['align'] == 1:
            left = right = n // 2
        elif p['align'] == 2:
            left, right = n - edge, edge
        else:
            left, right = edge, n - edge
        if left > 0:
            p['runs'].insert(0, [' ' * left, dict(padf)])
        if right > 0:
            p['runs'].append([' ' * right, dict(padf)])
        p['cell_pad'] = (0.5, 0.5)                # bant kenarlığa kadar uzansın
        p['nowrap_band'] = True
        used = True
    return used


def _valign(cells, row_h):
    """Dikey hizalama özniteliği yok: kısa hücrenin ilk paragrafına üst boşluk eklenir."""
    for c in cells:
        if c.get('table') or not c['paras']:
            continue
        gap = row_h - cell_height(c)
        if gap <= 0.5:
            continue
        va = c.get('valign') or 'top'
        if va == 'center':
            c['paras'][0]['before'] = (c['paras'][0].get('before') or 0) + gap / 2
            c['paras'][-1]['after'] = (c['paras'][-1].get('after') or 0) + gap / 2
        elif va == 'bottom':
            c['paras'][0]['before'] = (c['paras'][0].get('before') or 0) + gap
        else:
            c['paras'][-1]['after'] = (c['paras'][-1].get('after') or 0) + gap


def parse_table(tbl, ctx, warn):
    smap = ctx['smap']
    fill_mode = ctx.get('fill_mode', 'bant')
    tblpr = tbl.find(w('tblPr'))
    tstyle = {}
    if tblpr is not None and tblpr.find(w('tblStyle')) is not None:
        tstyle = resolve(tblpr.find(w('tblStyle')).get(w('val')), smap)
    borders = dict(tstyle.get('tblBorders', {}))
    if tblpr is not None and tblpr.find(w('tblBorders')) is not None:
        borders.update(_borders(tblpr.find(w('tblBorders'))))
    # tablo stilinin paragraf/karakter özellikleri hücre paragraflarına taban olur
    extra = {k: v for k, v in tstyle.items()
             if k in ('before', 'after', 'line', 'jc') + FMT_KEYS}
    tblmar = {'top': 0.0, 'left': 5.4, 'bottom': 0.0, 'right': 5.4}        # Word varsayılanı
    tm = tblpr.find(w('tblCellMar')) if tblpr is not None else None
    if tm is not None:
        for side, alt in (('top', None), ('left', 'start'), ('bottom', None), ('right', 'end')):
            el = tm.find(w(side))
            if el is None and alt:
                el = tm.find(w(alt))
            if el is not None and el.get(w('w')):
                tblmar[side] = float(el.get(w('w'))) / 20

    grid = [float(g.get(w('w')) or 0) / 20
            for g in tbl.findall(f'{w("tblGrid")}/{w("gridCol")}')]
    raw = []
    cell_border = []                              # hücre düzeyinde kenarlık: True=çizgisiz
    cell_drawn = []                               # hücre kendi kenarlığını çiziyor mu (en az iki kenar)
    for tr in tbl.findall(w('tr')):
        min_h = 0.0
        th = tr.find(f'{w("trPr")}/{w("trHeight")}')
        if th is not None and th.get(w('val')):
            min_h = float(th.get(w('val'))) / 20
        cells = []
        for tc in tr.findall(w('tc')):
            tcpr = tc.find(w('tcPr'))
            span, fill, vmerge, valign = 1, None, None, 'top'
            tcb = tcpr.find(w('tcBorders')) if tcpr is not None else None
            sides = _borders(tcb) if tcb is not None else {}
            cell_border.append(bool(sides) and all(sides.get(k) in ('none', 'nil')
                                                   for k in ('top', 'left', 'bottom', 'right')))
            cell_drawn.append(sum(1 for k in ('top', 'left', 'bottom', 'right')
                                  if sides.get(k) not in (None, 'none', 'nil')) >= 2)
            if tcpr is not None:
                gs = tcpr.find(w('gridSpan'))
                if gs is not None:
                    span = max(1, int(gs.get(w('val')) or 1))
                vm = tcpr.find(w('vMerge'))
                if vm is not None:
                    vmerge = 'restart' if vm.get(w('val')) == 'restart' else 'continue'
                va = tcpr.find(w('vAlign'))
                if va is not None and va.get(w('val')) in ('center', 'bottom'):
                    valign = va.get(w('val'))
                shd = tcpr.find(w('shd'))
                if shd is not None:
                    f_ = (shd.get(w('fill')) or '').upper()
                    if re.fullmatch(r'[0-9A-F]{6}', f_) and f_ != 'FFFFFF':
                        fill = f_
            paras = []
            for p in cell_paragraphs(tc):
                paras.extend(b for b in parse_paragraph(p, ctx, in_cell=True, extra=extra, warn=warn)
                             if b['type'] == 'p')
            if tc.find(f'.//{w("tbl")}') is not None:
                warn.add('hücre içindeki tablo düz paragraflara çevrildi')
            cells.append({'paras': paras or [_empty_para(True)], 'span': span, 'vmerge': vmerge,
                          'valign': valign, 'fill': fill, 'mar': _cell_margins(tcpr, tblmar)})
        raw.append({'cells': cells, 'min_h': min_h})

    ncol = len(grid) or max((sum(c['span'] for c in r['cells']) for r in raw), default=0)
    if not grid:
        grid = [1.0] * ncol
    # Editör: columnSpans oranları, toplam = 100 x sütun sayısı
    total = sum(grid) or 1
    spans = [int(round(g / total * 100 * ncol)) for g in grid] if ncol else []
    if spans:
        spans = [max(1, x) for x in spans]        # çok dar sütun 0'a yuvarlanmasın
        spans[spans.index(max(spans))] += 100 * ncol - sum(spans)   # fark en geniş sütuna
    text_w = ctx.get('text_w', 450.0)
    filled = False
    for r in raw:
        used = sum(c['span'] for c in r['cells'])
        while used < ncol:                        # eksik sütunları boş hücreyle tamamla
            r['cells'].append({'paras': [_empty_para(True)], 'span': 1, 'vmerge': None,
                               'valign': 'top', 'fill': None, 'mar': dict(tblmar)})
            used += 1
        gi, keep = 0, []
        for c in r['cells']:
            if gi >= ncol:
                warn.add('satırda sütun sayısından fazla hücre; fazlası atıldı')
                break
            n = min(c['span'], ncol - gi)
            c['col'], c['span'] = gi, n
            c['weight'] = sum(spans[gi:gi + n])
            keep.append(c)
            gi += n
            m = c['mar']
            cw = text_w * c['weight'] / (100.0 * ncol)
            for p in c['paras']:
                p['cell_pad'] = (max(1.0, m['left']), max(1.0, m['right']))
                ind = p.get('indent') or {}
                p['avail'] = max(20.0, cw - p['cell_pad'][0] - p['cell_pad'][1] - 2
                                 - (ind.get('left') or 0) - (ind.get('right') or 0))
            c['paras'][0]['before'] = (c['paras'][0].get('before') or 0) + max(1.0, m['top'])
            c['paras'][-1]['after'] = (c['paras'][-1].get('after') or 0) + max(1.0, m['bottom'])
            if c['vmerge'] != 'continue':
                filled = _pad_fill(c, fill_mode) or filled
        r['cells'] = keep
    if filled:
        warn.add('hücre dolgusu UDF\'de yok: yazının arkası hücre rengiyle boyandı' +
                 (' ve tek satırlık hücrelerde bant satır boyunca uzatıldı' if fill_mode == 'bant' else ''))
    elif any(c.get('fill') for r in raw for c in r['cells']):
        warn.add('hücre gölgelendirmesi (dolgu rengi) UDF\'ye taşınmadı')

    rows = _nest_vmerge(raw, spans, ncol, warn)
    for r in rows:                                # dikey hizalama + Word'ün en az satır yüksekliği
        h = max([cell_height(c) for c in r['cells']] + [r.get('min_h') or 0])
        _valign(r['cells'], h)
        r['colspans'] = [c['weight'] for c in r['cells']] if len(r['cells']) != ncol else None
    kind = _border_kind(borders)
    if cell_border and all(cell_border):          # bütün hücreler "kenarlık yok": Word tabloyu çizgisiz gösterir
        kind = 'borderNone'
    elif kind != 'borderCell' and cell_drawn and sum(cell_drawn) * 2 >= len(cell_drawn):
        # Kenarlığı tabloda değil hücrelerde tanımlı tablo (web sayfasından / yapay zekâ sohbetinden
        # yapıştırılan HTML tabloları böyle gelir): Word her hücreyi kutu içinde çizer (06.10.2026: çizgisiz çıkıyordu)
        kind = 'borderCell'
    for r in rows:                                # iç tablolar dış tablonun kenarlık türünü alır
        for c in r['cells']:
            if c.get('table'):
                c['table']['border'] = kind
    return {'type': 'tbl', 'ncol': ncol, 'spans': spans, 'rows': rows, 'border': kind}


def _nest_vmerge(raw, spans, ncol, warn):
    """Dikey birleştirme UDF modelinde yok. "Temiz" gruplar (gruptaki bütün birleşik hücreler
    grubun tüm satırlarını kapsıyorsa) tek dış satıra çevrilir: birleşik sütun = dış hücre,
    aradaki sütun blokları = hücre içinde iç tablo. Editör iç içe tabloyu çiziyor (20.09.2026)."""
    n = len(raw)
    regions = []                                  # (r0, r1, col, span)
    for ri, r in enumerate(raw):
        for c in r['cells']:
            if c['vmerge'] == 'restart':
                r1 = ri
                while r1 + 1 < n and any(d['vmerge'] == 'continue' and d['col'] == c['col'] and
                                         d['span'] == c['span'] for d in raw[r1 + 1]['cells']):
                    r1 += 1
                if r1 > ri:
                    regions.append((ri, r1, c['col'], c['span']))
    out, ri = [], 0
    while ri < n:
        grp = [g for g in regions if g[0] <= ri <= g[1]]
        if not grp:
            if any(c['vmerge'] == 'continue' for c in raw[ri]['cells']):
                warn.add('dikey birleştirilmiş hücre taşınamadı; alt hücre boş kaldı')
            out.append(raw[ri])
            ri += 1
            continue
        g0, g1 = min(g[0] for g in grp), max(g[1] for g in grp)
        changed = True
        while changed:                            # örtüşen bölgeleri tek grupta topla
            changed = False
            for g in regions:
                if g[0] <= g1 and g[1] >= g0 and (g[0] < g0 or g[1] > g1):
                    g0, g1, changed = min(g0, g[0]), max(g1, g[1]), True
        grp = [g for g in regions if g0 <= g[0] and g[1] <= g1]
        if any((g[0], g[1]) != (g0, g1) for g in grp):
            warn.add('karmaşık dikey birleştirme taşınamadı; alt hücreler boş ve çizgili kaldı')
            out.extend(raw[g0:g1 + 1])
            ri = g1 + 1
            continue
        merged_cols = {}
        for g in grp:
            merged_cols[g[2]] = g
        cells, col = [], 0
        row_h = [0.0] * (g1 - g0 + 1)
        blocks = []
        while col < ncol:
            if col in merged_cols:
                g = merged_cols[col]
                mc = next(c for c in raw[g0]['cells'] if c['col'] == col)
                cells.append(mc)
                col += g[3]
                continue
            end = col
            while end < ncol and end not in merged_cols:
                end += 1
            brows = []
            for k, r in enumerate(raw[g0:g1 + 1]):
                bc = [c for c in r['cells'] if col <= c['col'] < end]
                brows.append({'cells': bc, 'min_h': r.get('min_h') or 0})
                row_h[k] = max([row_h[k], r.get('min_h') or 0] + [cell_height(c) for c in bc])
            sub = {'type': 'tbl', 'ncol': end - col, 'spans': spans[col:end], 'rows': brows, 'border': None}
            blk = {'table': sub, 'span': end - col, 'col': col, 'weight': sum(spans[col:end]),
                   'paras': [], 'valign': 'top', 'fill': None, 'vmerge': None}
            blocks.append(blk)
            cells.append(blk)
            col = end
        # birleşik hücre iç satırların toplamından uzunsa fark son iç satıra verilir
        tall = max([cell_height(c) for c in cells if not c.get('table')] + [0])
        if tall > sum(row_h) + 2 * len(row_h):
            row_h[-1] += tall - sum(row_h) - 2 * len(row_h)
        for blk in blocks:                        # bloklar arasında satırlar aynı hizada kalsın
            blk['table']['row_h'] = list(row_h)
            for k, br in enumerate(blk['table']['rows']):
                _valign(br['cells'], row_h[k])
                br['colspans'] = [c['weight'] for c in br['cells']] \
                    if len(br['cells']) != blk['table']['ncol'] else None
        out.append({'cells': cells, 'min_h': 0})
        ri = g1 + 1
    return out


def split_para(p, avail, nlines):
    """Paragrafı tahmini satır kırılımına göre nlines. satırın sonundan ikiye böler."""
    toks = []
    for t, f in p['runs']:
        if 'image' in f:
            toks.append((t, f, f['image']['w'], f['image']['w']))
            continue
        for m in re.finditer(r'\S+\s*|\s+', t):
            toks.append((m.group(0), f, text_width(m.group(0), f), text_width(m.group(0).rstrip(), f)))
    line, x, cut = 1, (p.get('indent') or {}).get('firstLine') or 0.0, None
    for k, (_, _, wf, wt) in enumerate(toks):
        if x + wt > avail and x > 0:
            line, x = line + 1, 0.0
            if line > nlines:
                cut = k
                break
        x += wf
    if not cut:
        return p, None

    def runs_of(ts):
        runs = []
        for txt, f, _, _ in ts:
            if runs and runs[-1][1] == f and 'image' not in f:
                runs[-1][0] += txt
            else:
                runs.append([txt, dict(f)])
        return runs
    r1, r2 = runs_of(toks[:cut]), runs_of(toks[cut:])
    if r1 and 'image' not in r1[-1][1]:
        r1[-1][0] = r1[-1][0].rstrip()
        if not r1[-1][0]:
            r1.pop()
    q1 = dict(p, runs=r1, after=None)
    q2 = dict(p, runs=r2, before=None, list='',
              indent={k: v for k, v in (p.get('indent') or {}).items() if k not in ('firstLine', 'hanging')})
    q2.pop('float', None)
    return q1, q2


def layout_floats(blocks, ctx, warn):
    """Metin kaydırmalı (kare/sıkı) serbest görsel: Editör'de görselin yanından metin akıtma yok
    (görsel hep satır içi). Kenarlıksız iki sütunlu tabloyla taklit edilir: bir hücrede görsel,
    diğerinde görselin yüksekliğine sığan metin; taşan paragraf tahmini satır sınırından bölünür
    ve kalanı tablonun altında tam genişlikte sürer."""
    text_w = ctx.get('text_w', 450.0)
    out, i = [], 0
    while i < len(blocks):
        b = blocks[i]
        fl = b.get('float') if b['type'] == 'p' else None
        if not fl:
            out.append(b)
            i += 1
            continue
        b = dict(b)
        b.pop('float')
        iw, ih = fl['w'], fl['h']
        if iw > 0.62 * text_w:                    # yanında metne yer yok: görsel kendi satırında
            ip = _empty_para()
            ip.update(align=0 if fl['side'] == 'left' else 2, avail=text_w, after=6.0,
                      runs=[[IMG_CHAR, {'image': {k: fl[k] for k in ('data', 'w', 'h')}}]])
            out.extend([ip, b])
            i += 1
            continue
        gap = fl['gap']
        side_w = text_w - iw - gap - 6            # 6: hücrelerin kendi iç payı
        paras, used, rest, j = [], 0.0, None, i
        while j < len(blocks):
            q = b if j == i else blocks[j]
            if q['type'] != 'p' or (j > i and q.get('float')):
                break
            ind = q.get('indent') or {}
            qa = max(40.0, side_w - (ind.get('left') or 0) - (ind.get('right') or 0))
            h, lines, _ = para_metrics(q, qa)
            if used + h <= ih + 4:
                paras.append(q)
                used += h
                j += 1
                continue
            size = max([f.get('size') or 12 for _, f in q['runs'] if 'image' not in f] or [q.get('base_size') or 12])
            line_h = size * LINE_K * (1 + (q.get('line') if q.get('line') is not None else 0.15))
            n = int((ih - used - (q.get('before') or 0)) // line_h)
            if n >= 2 and lines - n >= 2:
                q1, q2 = split_para(q, qa, n)
                if q2 is not None:
                    paras.append(q1)
                    rest = q2
                    j += 1
                    break
            if not paras or (ih - used) > 0.5 * h:
                paras.append(q)
                j += 1
            break
        if not paras:
            paras = [_empty_para()]
        left = fl['side'] == 'left'
        for q in paras:
            ind = q.get('indent') or {}
            q.update(in_cell=True, cell_pad=(gap, 1.0) if left else (1.0, gap),
                     avail=max(40.0, side_w - (ind.get('left') or 0) - (ind.get('right') or 0)))
        ip = _empty_para(True)
        ip.update(align=0 if left else 2, line=0.0, cell_pad=(0.0, 0.0), avail=iw + 2,
                  runs=[[IMG_CHAR, {'image': {k: fl[k] for k in ('data', 'w', 'h')}}]])
        wi = min(190, max(10, int(round(200.0 * (iw + gap / 2 + 3) / text_w))))
        ci = {'paras': [ip], 'span': 1, 'weight': wi, 'valign': 'top', 'fill': None, 'vmerge': None}
        ct = {'paras': paras, 'span': 1, 'weight': 200 - wi, 'valign': 'top', 'fill': None, 'vmerge': None}
        cells = [ci, ct] if left else [ct, ci]
        out.append({'type': 'tbl', 'ncol': 2, 'spans': [c['weight'] for c in cells], 'border': 'borderNone',
                    'rows': [{'cells': cells, 'colspans': None, 'min_h': 0}]})
        if rest is not None:
            out.append(rest)
        warn.add('metin kaydırmalı görsel kenarlıksız iki sütunlu tabloyla taklit edildi (Editör\'de görselin '
                 'yanından metin akıtma yok); metnin görselin altına geçtiği satır tahminidir')
        i = j
    return out


# Editör'ün mediaSizeName kodları (21.09.2026'da Editör motorunda sayfa ölçülerek bulundu)
MEDIA = ((1, 'A4', 595.3, 841.9), (2, 'A5', 419.5, 595.3), (3, 'Letter', 612.0, 792.0),
         (4, 'B5', 515.9, 728.5), (5, 'Legal', 612.0, 1008.0), (6, 'Executive', 522.0, 756.0))


def media_code(page_w, page_h):
    """(kod, ad) — sayfa ölçüsüne uyan Editör kâğıt kodu; uymuyorsa (None, None)."""
    a, b = sorted((page_w, page_h))
    for code, name, mw, mh in MEDIA:
        if abs(a - mw) <= 6 and abs(b - mh) <= 6:
            return code, name
    return None, None


def sect_info(body):
    """Son sectPr: kenar boşlukları (punto), sayfa boyutu, üst/altbilgi referansları."""
    sp = body.find(w('sectPr')) if body is not None else None
    info = {'margins': (56.7, 56.7, 70.9, 70.9), 'page_w': 595.3, 'page_h': 841.9,
            'titlePg': False, 'landscape': False, 'header': {}, 'footer': {}}
    if sp is None:
        return info
    pm = sp.find(w('pgMar'))
    if pm is not None:
        g = lambda k: float(pm.get(w(k)) or 0) / 20
        info['margins'] = (g('left'), g('right'), g('top'), g('bottom'))
        if pm.get(w('header')):
            info['hdr_dist'] = g('header')
    ps = sp.find(w('pgSz'))
    if ps is not None and ps.get(w('w')):
        info['page_w'] = float(ps.get(w('w'))) / 20
        info['page_h'] = float(ps.get(w('h')) or 16838) / 20
        info['landscape'] = ps.get(w('orient')) == 'landscape' or info['page_w'] > info['page_h']
    info['titlePg'] = sp.find(w('titlePg')) is not None
    # Word'de üst/altbilgi başvurusu olmayan bölüm, öncekinin üst/altbilgisini DEVRALIR. Editör tek
    # üst/altbilgi tuttuğundan son bölümün (devralınmış dâhil) başvuruları kullanılır.
    for s_ in [x for x in body.iter(w('sectPr')) if x is not sp] + [sp]:
        for kind in ('header', 'footer'):
            for ref in s_.findall(w(f'{kind}Reference')):
                info[kind][ref.get(w('type')) or 'default'] = ref.get(f'{{{R}}}id')
    return info


def _has_text(paras):
    return any(t.strip() or 'image' in f for p in paras for t, f in p['runs'])


def read_hf(z, ctx, info, doc_rels, warn):
    """{'header': [(paragraflar, sayfa_no_var_mı, aralık_özniteliği)], 'footer': [...]}"""
    out, bgs = {}, {}
    for kind in ('header', 'footer'):
        cands = {}
        for typ, rid in info[kind].items():
            part = doc_rels.get(rid)
            if not part:
                continue
            try:
                raw = z.read(part)
            except KeyError:
                continue
            root = ET.fromstring(raw)
            ctx['rels'] = part_rels(z, part)
            ctx['bg_pick'] = None
            if kind == 'header':                  # en büyük, metni kaydırmayan serbest görsel: antet adayı
                best = 0.0
                for an in root.iter(f'{{{WP}}}anchor'):
                    ext = an.find(f'{{{WP}}}extent')
                    if an.find(f'{{{WP}}}wrapNone') is None or ext is None or an.find(f'.//{{{A}}}blip') is None:
                        continue
                    try:
                        area = float(ext.get('cx')) * float(ext.get('cy'))
                    except (TypeError, ValueError):
                        continue
                    if area > best:
                        best, ctx['bg_pick'] = area, an
                for sh in root.iter(f'{{{V}}}shape'):     # eski tip: Word'ün "filigran resmi" (VML)
                    st = _vml_style(sh)
                    if st.get('position') != 'absolute' or sh.find(f'{{{V}}}imagedata') is None:
                        continue
                    area = (_len_pt(st.get('width')) or 0) * (_len_pt(st.get('height')) or 0) * 12700 * 12700
                    if area > best:
                        best, ctx['bg_pick'] = area, sh
            paras = []
            for b in iter_blocks(root):
                for p in ([b] if b.tag == w('p') else
                          [q for tr in b.findall(w('tr')) for c in tr.findall(w('tc'))
                           for q in cell_paragraphs(c)]):
                    paras.extend(x for x in parse_paragraph(p, ctx, warn=warn) if x['type'] == 'p')
            instrs = _field_instrs(root)
            has_page = any(_is_page_field(x) for x in instrs)
            # "Sayfa 3 / 12" gibi yalnız sayfa numarası kalıbından ibaret paragraf: numarayı Editör
            # kendisi çizer; alan çıkınca geriye kalan "Sayfa  / " artığı yazılmaz.
            if has_page:                          # pageNumber-spec bitleri: 8 taban, 32 orta, 64 sağ, 2048 toplam
                al = next((q['align'] for q in paras if q.get('pagefld')), 1)
                total = any(x.strip().split()[:1] in (['NUMPAGES'], ['SECTIONPAGES']) for x in instrs)
                has_page = 8 + {1: 32, 2: 64}.get(al, 0) + (2048 if total else 0)
            kept = []
            for q in paras:
                rest = ''.join(t for t, f in q['runs'] if 'image' not in f)
                if q.get('pagefld') and _page_pattern_only(rest):
                    continue
                if q.get('pagefld') and rest.strip():
                    warn.add(f'{"üstbilgideki" if kind == "header" else "altbilgideki"} sayfa numarasının '
                             'yanındaki metin korundu; numarayı Editör altbilgide ayrı satırda çizer')
                kept.append(q)
            paras = kept
            if ctx.get('bg_found'):
                bgs[typ] = ctx.pop('bg_found')
            ctx['bg_pick'] = None
            if _has_text(paras) or has_page:
                cands[typ] = (paras, has_page)
        # Editör tek üstbilgi/altbilgi öğesi tutar (birden çok yazılırsa sonuncusu kalır;
        # 20.09.2026 önizlemeyle sınandı). Aralık: stopPage="1" yalnız ilk sayfa,
        # startPage="2" ikinci sayfadan itibaren (derlemde 19 örnek).
        items = []
        text = lambda c: [t for p in c[0] for t, _ in p['runs']]
        first_c, def_c = cands.get('first'), cands.get('default')
        ad = 'üstbilgi' if kind == 'header' else 'altbilgi'
        if not info['titlePg']:
            if def_c:
                items = [def_c + ('',)]
        elif first_c and def_c and text(first_c) != text(def_c):
            items = [first_c + (' stopPage="1"',)]
            warn.add(f'ilk sayfa {ad}si korundu; devam sayfalarının farklı {ad}si taşınmadı '
                     f'(Editör tek {ad} tutar)')
        elif first_c and def_c:
            items = [def_c + ('',)]
        elif first_c:
            items = [first_c + (' stopPage="1"',)]
        elif def_c:
            items = [def_c + (' startPage="2"',)]
        out[kind] = items
    # Sayfa numarası yalnız üstbilgideyse kaybolmasın: UDF'de sayfa numarası altbilginin özniteliğidir
    hp = next((x[1] for x in out.get('header', []) if x[1]), None)
    if hp and not any(x[1] for x in out.get('footer', [])):
        out['footer'] = [(x[0], hp, x[2]) for x in out.get('footer', [])] or [([], hp, '')]
        warn.add('sayfa numarası Word\'de üstbilgideydi; UDF\'de altbilgiye yazıldı')
    # Editör'de tek sayfa arka planı vardır ve her sayfada görünür
    b_first, b_def = (bgs.get('first') if info['titlePg'] else None), bgs.get('default')
    bg = b_def or b_first
    if bg:
        if info['titlePg'] and b_first and b_def and b_first['data'] != b_def['data']:
            bg = b_first
            warn.add('antet (üstbilgideki arka plan görseli) sayfa arka planı olarak yazıldı; ilk sayfanınki '
                     'kullanıldı, devam sayfalarındaki farklı görsel taşınmadı (Editör tek arka plan tutar)')
        elif info['titlePg'] and not (b_first and b_def):
            warn.add('antet (üstbilgideki arka plan görseli) sayfa arka planı olarak yazıldı; Word\'de yalnız '
                     + ('ilk sayfadaydı' if b_first else 'devam sayfalarındaydı') + ', Editör\'de her sayfada görünür')
        else:
            warn.add('antet (üstbilgideki arka plan görseli) UDF\'de sayfa arka planı olarak yazıldı')
        info['bg'] = bg
    return out


def read_document(docx, warn, fill_mode='bant', yan_gorsel=False):
    z = zipfile.ZipFile(docx)
    names = set(z.namelist())
    if 'word/document.xml' not in names:
        raise ValueError('word/document.xml yok; bu bir Word (DOCX) belgesi değil')
    theme = theme_fonts(z)
    smap = style_map(z.read('word/styles.xml'), theme) if 'word/styles.xml' in names \
        else {'__docdefault__': {}}
    root = ET.fromstring(z.read('word/document.xml'))
    body = root.find(w('body'))
    info = sect_info(body)
    l, r, t, b = info['margins']
    ctx = {'smap': smap, 'theme': theme, 'nmap': numbering_map(z), 'counters': {}, 'zip': z,
           'rels': part_rels(z, 'word/document.xml'), 'notes': [], 'noteparts': {}, 'cur_note': None,
           'fill_mode': fill_mode, 'body': True, 'yan_gorsel': yan_gorsel, 'margin_left': l,
           'page': (info['page_w'], info['page_h']), 'margins': (l, r, t, b), 'hdr_dist': info.get('hdr_dist', 35.4),
           'text_w': max(100.0, info['page_w'] - l - r), 'text_h': max(100.0, info['page_h'] - t - b)}
    doc_rels = dict(ctx['rels'])
    for kind in ('footnote', 'endnote'):
        part = f'word/{kind}s.xml'
        if part in names:
            nroot = ET.fromstring(z.read(part))
            ctx['noteparts'][kind] = {n.get(w('id')): n for n in nroot.findall(w(kind))
                                      if n.get(w('type')) in (None, 'normal')}
    blocks = []
    for blk in (iter_blocks(body) if body is not None else ()):
        if blk.tag == w('p'):
            blocks.extend(parse_paragraph(blk, ctx, warn=warn))
        else:
            blocks.append(parse_table(blk, ctx, warn))
    ctx['body'] = False                           # dipnot ve üst/altbilgide yerleşim taklidi yapılmaz
    blocks = layout_floats(blocks, ctx, warn)
    if ctx['notes']:                              # dipnot/sonnot: belge sonuna numaralı paragraflar
        rule = _empty_para()
        rule['runs'] = [['_' * 24, {}]]
        blocks.append(rule)
        for kind, el, lab in list(ctx['notes']):
            ctx['cur_note'] = lab
            first = True
            for p in el.findall(w('p')):
                for x in parse_paragraph(p, ctx, warn=warn):
                    if x['type'] != 'p':
                        continue
                    rs = x['runs']
                    if first and len(rs) > 1 and rs[0][0] == lab and not rs[1][0][:1].isspace():
                        rs.insert(1, [' ', {k: v for k, v in rs[1][1].items() if k != 'image'}])
                    first = False
                    blocks.append(x)
        ctx['cur_note'] = None
        warn.add('dipnotlar belge sonuna numaralı paragraf olarak taşındı (UDF\'de sayfa altı dipnot yok)')
    fixed = []
    for blk in blocks:                            # tablolar arasına ve sondaki tablo/sayfa sonu ardına paragraf
        if blk['type'] == 'tbl' and fixed and fixed[-1]['type'] == 'tbl':
            fixed.append(_empty_para())
        fixed.append(blk)
    if not fixed or fixed[-1]['type'] != 'p':
        fixed.append(_empty_para())
    if body is not None and any('PAGEREF' in (it.text or '').upper() or (it.text or '').strip().upper().startswith('TOC')
                                for it in body.iter(w('instrText'))):
        warn.add('içindekiler tablosundaki sayfa numaraları Word\'ün sayfalamasına aittir ve sabit metin olarak '
                 'taşındı; UDF\'de sayfalar kayarsa numaralar tutmayabilir')
    pw, ph = sorted((info['page_w'], info['page_h']))
    info['media'], ad = media_code(pw, ph)
    if info['media'] is None:
        info['media'] = 1
        warn.add(f'sayfa boyutu ({pw * 0.3528:.0f}x{ph * 0.3528:.0f} mm) Editör\'ün tanıdığı boyutlardan değil; '
                 'UDF A4 yazıldı, kenar boşlukları korunur ama satır sonları değişebilir')
    elif ad == 'Letter':
        warn.add('sayfa boyutu Letter (216x279 mm) ve UDF de Letter yazıldı; A4 istiyorsanız Word\'de '
                 'sayfa boyutunu A4 yapıp yeniden çevirin')
    hf = read_hf(z, ctx, info, doc_rels, warn)
    info['text_h'] = ctx['text_h']
    return fixed, info, hf


def _java_color(hex6):
    v = 0xFF000000 | int(hex6, 16)
    return v - (1 << 32) if v >= (1 << 31) else v


def _xattr(s):
    return s.replace('&', '&amp;').replace('"', '&quot;').replace('<', '&lt;').replace('>', '&gt;')


def attrs(fmt):
    a = ''
    if fmt.get('bold'):
        a += ' bold="true"'
    if fmt.get('italic'):
        a += ' italic="true"'
    if fmt.get('underline'):
        a += ' underline="true"'
    if fmt.get('strike'):
        a += ' strikethrough="true"'
    if fmt.get('super'):
        a += ' superscript="true"'
    elif fmt.get('sub'):
        a += ' subscript="true"'
    if fmt.get('size') and fmt['size'] != 12:
        a += f' size="{int(fmt["size"])}"'
    if fmt.get('color'):
        a += f' foreground="{_java_color(fmt["color"])}"'
    if fmt.get('bg'):
        a += f' background="{_java_color(fmt["bg"])}"'
    if fmt.get('font') and fmt['font'] != DEFAULT_FONT:
        a += f' family="{_xattr(fmt["font"])}"'
    return a


def _num(v):
    """3 -> '3.0', 0.15 -> '0.15' (örnek UDF'lerdeki ondalık biçim)."""
    s = f'{v:.6g}'
    return s if '.' in s or 'e' in s else s + '.0'


def tabset(tabs, limit=None, shift=0.0):
    """Word sekme durakları -> UDF TabSet ("pos:hiza:lider,...").

    Word durak konumunu SAYFA KENARINDAN, Editör ise paragrafın SOL GİRİNTİSİNDEN ölçer
    (21.09.2026'da Editör motorunda ölçüldü: LeftIndent 50 artınca sekme sonrası metin de 50 kayıyor).
    Bu yüzden her konumdan paragrafın LeftIndent'i (shift) çıkarılır. Özel durakların ardından
    Word'ün varsayılan 36 puntoluk durakları (kenardan ölçülü) 648'e kadar sürer."""
    stops = {}
    for pos, val, leader in tabs or []:
        if val == 'clear':
            stops.pop(round(pos, 1), None)
            continue
        if limit and TAB_ALIGN.get(val, 0) != 0 and pos > limit - TAB_EDGE:
            pos = limit - TAB_EDGE                # kenara dayalı sağa/ortaya hizalı durak taşmasın
        stops[round(pos, 1)] = (TAB_ALIGN.get(val, 0), TAB_LEADER.get(leader, 0))
    if not stops and shift < 0.05:
        return ''                                 # stilin varsayılan durakları yeter
    last = max(stops) if stops else 0.0
    items = [f'{p - shift:.1f}:{a}:{l}' for p, (a, l) in sorted(stops.items()) if p - shift > 0.5]
    x = 36.0
    while x <= 648.0:
        if x > last + 0.05 and x - shift > 0.5:
            items.append(f'{x - shift:.1f}:0:0')
        x += 36.0
    return f' TabSet="{",".join(items)}"' if items else ''


def para_xml(p, off, max_h=None):
    """Paragraf -> (xml, metin). Metin '\\n' ile biter; off ve uzunluklar UTF-16 birimidir."""
    # "]]>" CDATA'yı kapatır; Editör bölünmüş CDATA'yı okuyamaz (20.09.2026'da motorla sınandı).
    # Dizi, araya boşluk konarak zararsızlaştırılır; run sınırına denk gelen hâl de kapsanır.
    tail = ''
    for run in p['runs']:
        if 'image' in run[1]:
            tail = ''
            continue
        joined = (tail + run[0]).replace(']]>', ']] >')
        run[0] = joined[len(tail):]
        tail = joined[-2:]
    body = ''.join(t for t, _ in p['runs'])
    parts = []
    cur = off
    for t, fmt in p['runs']:
        n = u16(t)
        if 'image' in fmt:
            im = fmt['image']
            iw, ih = im['w'], im['h']
            avail = p.get('avail')
            if avail and iw > avail:                          # alana sığdır (orantılı)
                ih, iw = ih * avail / iw, avail
            if max_h and ih > max_h:
                iw, ih = iw * max_h / ih, max_h
            parts.append(f'<image imageData="{im["data"]}" width="{iw:.1f}" '
                         f'height="{ih:.1f}" startOffset="{cur}" length="{n}" />')
        else:
            parts.append(f'<content{attrs(fmt)} resolver="defaultdefault" '
                         f'startOffset="{cur}" length="{n}" />')
        cur += n
    parts.append(f'<content startOffset="{cur}" length="1" />')   # satır sonu
    ind = p.get('indent') or {}
    ia = ''
    right = ind.get('right')
    if p.get('list') and not p.get('literal'):
        # Yerel liste: Editör numarayı/işareti LeftIndent'in SOLUNA çizer, metin LeftIndent'te
        # başlar ve devam satırları da orada hizalanır. Word'ün "left" değeri = metin konumu.
        # (Derlemde 1305 liste paragrafında Hanging yok; 20.09.2026 önizlemeyle doğrulandı.)
        left, hanging, first = (ind.get('left') or 0), None, None
    else:
        # UDF modeli: ilk satır = LeftIndent (+FirstLineIndent), devam satırları =
        # LeftIndent + Hanging. Word'de devam satırları = left, ilk satır = left - hanging.
        left = (ind.get('left') or 0) - (ind.get('hanging') or 0)
        hanging, first = ind.get('hanging'), ind.get('firstLine')
    tab_base = max(left, 0.0)                     # Editör sekmeyi buradan ölçer (hücre payı hariç)
    if p.get('in_cell'):                          # hücre iç boşluğu (Word tcMar/tblCellMar) girintiye eklenir
        pl, pr = p.get('cell_pad') or (CELL_LEFT, CELL_RIGHT)
        left += pl
        right = (right or 0) + pr
    if left > 0.05:
        ia += f' LeftIndent="{left:.1f}"'
    if right:
        ia += f' RightIndent="{right:.1f}"'
    if hanging:
        ia += f' Hanging="{hanging:.1f}"'
    if first:
        ia += f' FirstLineIndent="{first:.1f}"'
    if p.get('before'):
        ia += f' SpaceAbove="{_num(p["before"])}"'
    if p.get('after'):
        ia += f' SpaceBelow="{_num(p["after"])}"'     # w:spacing after (stil/docDefaults dâhil)
    if p.get('tabs') or ('\t' in body and tab_base > 0.05):
        ia += tabset(p.get('tabs'), (p.get('text_w') or 0) - (ind.get('right') or 0), tab_base)
    ia += p.get('list') or ''
    ls = p.get('line')
    ls = 0.15 if ls is None else ls               # eski davranış: 1.15 satır aralığı
    xml = (f'<paragraph Alignment="{p["align"]}"{ia} LineSpacing="{_num(ls)}" '
           f'resolver="defaultdefault">' + ''.join(parts) + '</paragraph>')
    return xml, body + '\n', cur + 1


def rule_para(color, text_w):
    """Paragraf üst kenarlığı yerine alt çizgi satırı (büronun Editör antetindeki gibi)."""
    n = max(10, int(text_w / 6.0))                # TNR 12 punto '_' ≈ 6 punto
    d = _empty_para()
    d.update(align=1, runs=[['_' * n, {'bold': True, 'color': color}]], line=0.0)
    return d


FOOTER_PAGE = (' pageNumber-spec="BSP32_2120" pageNumber-seperator="/" '
               'pageNumber-fontBold="false" pageNumber-fontItalic="false" '
               'pageNumber-fontFace="Times New Roman" pageNumber-fontSize="10" '
               'pageNumber-foreStr="" pageNumber-pageStartNumStr=""')
# Büronun Editör'de yaptığı antetli UDF'den: sayfa numarasız altbilgi
FOOTER_NOPAGE = (' pageNumber-spec="BSP32_0" pageNumber-color="-16777216" '
                 'pageNumber-fontFace="Arial Narrow" pageNumber-fontSize="11" '
                 'pageNumber-foreStr="" pageNumber-pageStartNumStr="" '
                 'pageNumber-seperator="/" pageNumber-fontBold="true" '
                 'pageNumber-fontItalic="false"')


def build(blocks, info, hf, page_mode='otomatik'):
    """page_mode: 'otomatik' = DOCX'e sadık (altbilgide PAGE alanı varsa ya da altbilgi
    hiç yoksa sayfa numarası), 'var' / 'yok' = zorla."""
    text = []
    els = []
    state = {'off': 0}
    l, r, t, bm = info['margins']
    text_w = info['page_w'] - l - r
    max_h = max(100.0, info.get('text_h', 700.0) - 24)

    def emit_paras(paras, with_rule):
        xs = []
        for p in paras:
            if with_rule and p.get('border_top'):
                xml, tx, state['off'] = para_xml(rule_para(p['border_top'], text_w), state['off'])
                xs.append(xml)
                text.append(tx)
            xml, tx, state['off'] = para_xml(p, state['off'], max_h)
            xs.append(xml)
            text.append(tx)
        return ''.join(xs)

    def table_xml(tb):
        rows = []
        for ri, rw in enumerate(tb['rows']):
            cells = []
            for c in rw['cells']:                 # hücre: paragraflar ya da (dikey birleştirme için) iç tablo
                cells.append('<cell>' + (table_xml(c['table']) if c.get('table')
                                         else emit_paras(c['paras'], False)) + '</cell>')
            cs = ''
            if rw.get('colspans'):
                cs = ' columnSpans="' + ','.join(str(x) for x in rw['colspans']) + '"'
            rows.append(f'<row rowName="row{ri + 1}" rowType="dataRow"{cs}>' + ''.join(cells) + '</row>')
        spans = ','.join(str(x) for x in tb['spans'])
        return (f'<table tableName="Sabit" columnCount="{tb["ncol"]}" columnSpans="{spans}" '
                f'border="{tb["border"]}">' + ''.join(rows) + '</table>')

    # üstbilgi(ler): CDATA'nın başında (Editör antetindeki düzen)
    for paras, _, rng in hf.get('header', []):
        els.append(f'<header{rng}>' + emit_paras(paras, True) + '</header>')

    for b in blocks:
        if b['type'] == 'p':
            els.append(emit_paras([b], False))
        elif b['type'] == 'pb':
            # Editör'ün insert-pageBreak eyleminin kurduğu yapı: tek '\n' taşıyan paragrafı saran öğe
            pb = _empty_para()
            els.append('<page-break>' + emit_paras([pb], False) + '</page-break>')
        else:
            els.append(table_xml(b))

    # altbilgi(ler): DOCX'te varsa metni; sayfa numarası alanı varsa ya da altbilgi HİÇ tanımlı değilse
    # (eski davranış) CDATA sonuna sayfa numarası için yer ayrılır. Word'de altbilgi tanımlı ama boşsa
    # Word sayfa numarası göstermez; UDF'ye de eklenmez (02.10.2026: boş altbilgili antetli belgeye
    # "1/11" ekleniyordu).
    footers = hf.get('footer', []) or ([] if info.get('footer') else [([], True, '')])
    for paras, has_page, rng in footers:
        if page_mode == 'var':
            has_page = True
        elif page_mode == 'yok':
            has_page = False
        fx = emit_paras(paras, True)
        if has_page:
            foot_off = state['off']
            text.append('  \n')
            state['off'] += 3
            fx += ('<paragraph Alignment="1" family="Times New Roman" size="12" description="Gövde">'
                   f'<content family="Times New Roman" size="12" description="Gövde" '
                   f'startOffset="{foot_off}" length="3" /></paragraph>')
        elif not fx:
            continue                              # ne metin ne sayfa numarası: altbilgi yazma
        spec = has_page if has_page and has_page is not True else 2120      # altbilgisiz belge: eski davranış
        els.append(f'<footer{rng}{FOOTER_PAGE.replace("BSP32_2120", f"BSP32_{spec}") if has_page else FOOTER_NOPAGE}>'
                   + fx + '</footer>')

    bg, bm_ = info.get('bg'), [0, 0, 0, 0]         # sol, üst, sağ, alt pay (punto)
    if bg:
        bm_ = [bg['left'], bg['top'], info['page_w'] - bg['left'] - bg['w'], info['page_h'] - bg['top'] - bg['h']]
        bm_ = [0 if v < 6.0 else int(round(v)) for v in bm_]      # tam sayfa antette küçük kaymalar yutulur
    bg_xml = (f'<bgImage bgImageSource="" bgImageData="{bg["data"] if bg else ""}" bgImageBottomMargin="{bm_[3]}" '
              f'bgImageUpMargin="{bm_[1]}" bgImageRigtMargin="{bm_[2]}" bgImageLeftMargin="{bm_[0]}" />')
    cdata = ''.join(text)
    safe = cdata                                   # ']]>' para_xml'de zaten zararsızlaştırıldı
    orient = '0' if info.get('landscape') else '1'            # java.awt.print.PageFormat: 0 yatay, 1 dikey
    xml = (
        '<?xml version="1.0" encoding="UTF-8" ?> \n\n'
        '<template format_id="1.8" >\n'
        f'<content><![CDATA[{safe}]]></content>\n'
        f'<properties><pageFormat mediaSizeName="{info.get("media", 1)}" leftMargin="{l}" rightMargin="{r}" '
        f'topMargin="{t}" bottomMargin="{bm}" paperOrientation="{orient}" headerFOffset="20.0" '
        'footerFOffset="35.43307065963745" />'
        + bg_xml + '</properties>\n'
        '<elements resolver="hvl-default" >\n' + ''.join(els) + '\n</elements>\n'
        '<styles><style name="default" description="Geçerli" family="Dialog" size="12" '
        'bold="false" italic="false" foreground="-13421773" />'
        '<style name="hvl-default" family="Times New Roman" size="12" description="Gövde" />'
        '<style name="defaultdefault" resolver="hvl-default" description="Gövde" '
        'TabSet="36.0:0:0,72.0:0:0,108.0:0:0,144.0:0:0,180.0:0:0,216.0:0:0,252.0:0:0,'
        '288.0:0:0,324.0:0:0,360.0:0:0,396.0:0:0,432.0:0:0,468.0:0:0,504.0:0:0,540.0:0:0,'
        '576.0:0:0,612.0:0:0,648.0:0:0" family="Times New Roman" size="12" /></styles>\n'
        '</template>\n')
    return xml, cdata


def check(xml, cdata):
    """Dilimler bitişik mi, CDATA'yı (UTF-16 birimiyle) tam kaplıyor mu; tablo satırları tutarlı mı."""
    root = ET.fromstring(xml.split('?> \n\n', 1)[1])
    total = u16(cdata)
    pos = 0
    for p in root.iter('paragraph'):
        for c in p:
            so, ln = c.get('startOffset'), c.get('length')
            if so is None:
                continue
            so, ln = int(so), int(ln)
            if so + ln > total:
                raise DonusumHatasi(f'offset taşması {so}+{ln} > {total}')
            if so != pos:
                raise DonusumHatasi(f'dilim boşluğu {pos} != {so}')
            pos = so + ln
    if pos != total:
        raise DonusumHatasi(f'CDATA tam kaplanmadı {pos} != {total}')
    for t in root.iter('table'):
        n = int(t.get('columnCount'))
        for r in t.findall('row'):
            want = len(r.get('columnSpans').split(',')) if r.get('columnSpans') else n
            if len(r.findall('cell')) != want:
                raise DonusumHatasi('tablo satırında hücre sayısı sütun tanımıyla uyuşmuyor')


def cevir(docx, cikti, sayfa_no='otomatik', hucre_dolgusu='bant', yan_gorsel=False):
    """DOCX'i UDF'ye çevirir; özet sözlüğü döndürür, sorun olursa DonusumHatasi atar.

    sayfa_no: otomatik | var | yok      hucre_dolgusu: bant | yazi | yok
    yan_gorsel: metin kaydırmalı görseli iki sütunlu tabloyla taklit et (deneysel, varsayılan kapalı)
    """
    warn = set()
    NONBMP_SEEN.clear()
    KUCULTULEN[:] = [0, 0]
    try:
        blocks, info, hf = read_document(docx, warn, hucre_dolgusu, yan_gorsel)
    except FileNotFoundError:
        raise DonusumHatasi(f'dosya bulunamadı: {docx}')
    except zipfile.BadZipFile:
        with open(docx, 'rb') as f:
            head = f.read(8)
        if head[:4] == b'\xd0\xcf\x11\xe0':
            raise DonusumHatasi('bu eski biçimli bir Word dosyası (.doc) ya da parola korumalı. '
                                "Word'de \"Farklı Kaydet → Word Belgesi (.docx)\" ile kaydedip yeniden deneyin.")
        raise DonusumHatasi('dosya bir DOCX (zip) arşivi değil ya da bozuk.')
    except (ValueError, ET.ParseError) as ex:
        raise DonusumHatasi(f'belge okunamadı: {ex}')
    if NONBMP_SEEN:
        warn.add('Editör\'ün açamadığı karakterler (emoji vb.) değiştirildi: '
                 + ' '.join(f'{c}→{_NONBMP_MAP.get(c, chr(0x25a1))}' for c in sorted(NONBMP_SEEN)))
    xml, cdata = build(blocks, info, hf, sayfa_no)
    check(xml, cdata)
    with zipfile.ZipFile(cikti, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('content.xml', xml.encode('utf-8'))
    if KUCULTULEN[0]:
        warn.add(f'{KUCULTULEN[0]} görsel sayfadaki boyutuna göre küçültüldü ({GORSEL_DPI} dpi; baskıda fark '
                 f'görülmez), dosya {KUCULTULEN[1] / 1048576:.1f} MB küçüldü')
    mb = os.path.getsize(cikti) / 1048576
    if mb > UYAP_SINIR_MB:
        warn.add(f'UDF dosyası {mb:.1f} MB; UYAP {UYAP_SINIR_MB} MB üzerindeki dosyayı kabul etmeyebilir. '
                 'Görsellerin bir kısmını ayrı ek olarak (PDF) yüklemeyi düşünün')
    return {'paragraf': sum(1 for b in blocks if b['type'] == 'p'),
            'tablo': sum(1 for b in blocks if b['type'] == 'tbl'),
            'sayfa_sonu': sum(1 for b in blocks if b['type'] == 'pb'),
            'karakter': u16(cdata), 'ustbilgi': bool(hf.get('header')), 'altbilgi': bool(hf.get('footer')),
            'uyarilar': sorted(warn)}


def main():
    ap = argparse.ArgumentParser(description='Word (DOCX) belgesini UYAP UDF biçimine çevirir.')
    ap.add_argument('docx')
    ap.add_argument('-o', '--output', required=True)
    ap.add_argument('--sayfa-no', choices=('otomatik', 'var', 'yok'), default='otomatik',
                    help="altbilgi sayfa numarası: otomatik (DOCX'e sadık), var, yok")
    ap.add_argument('--hucre-dolgusu', choices=('bant', 'yazi', 'yok'), default='bant',
                    help="tablo hücresi dolgu rengi taklidi: bant (yazı + satır boyu zemin bandı), "
                         "yazi (yalnız koyu zeminde açık yazının arkası), yok")
    ap.add_argument('--yan-gorsel', action='store_true',
                    help='DENEYSEL: metin kaydırmalı (yanında metin akan) görseli kenarlıksız iki sütunlu '
                         'tabloyla taklit et; varsayılan kapalı (görsel satır içine alınır)')
    a = ap.parse_args()
    try:
        o = cevir(a.docx, a.output, a.sayfa_no, a.hucre_dolgusu, a.yan_gorsel)
    except DonusumHatasi as ex:
        sys.exit(f'HATA: {ex}')
    extra = ', '.join(x for x in (f'{o["sayfa_sonu"]} sayfa sonu' if o['sayfa_sonu'] else '',
                                  'üstbilgi' if o['ustbilgi'] else '',
                                  'altbilgi' if o['altbilgi'] else '') if x)
    print(f'Yazıldı: {a.output} ({o["paragraf"]} paragraf, {o["tablo"]} tablo, {o["karakter"]} karakter'
          + (f', {extra}' if extra else '') + ')')
    for m in o['uyarilar']:
        print('UYARI:', m)


if __name__ == '__main__':
    main()
