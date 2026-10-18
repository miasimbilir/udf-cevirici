#!/usr/bin/env python3
"""udf_docx.py — UYAP UDF belgesini Word (DOCX) biçimine çevirir (docx_udf.py'nin tersi).

Kullanım:
    python3 araclar/udf_docx.py girdi.udf -o cikti.docx

UDF'nin biçim modeli Word'ünkinin alt kümesidir; taşınanlar:
  - paragraf: hizalama, girintiler, üst/alt boşluk, satır aralığı, sekme durakları
  - yazı: kalın, italik, altı/üstü çizili, üst/alt simge, punto, yazı tipi, renk, zemin rengi
  - yerel listeler (NumberType / BulletType / SecListTypeLevelN / NumberSetted) -> numbering.xml
  - tablolar: sütun oranları, satır düzeyinde yatay birleştirme (gridSpan), iç içe tablo,
    kenarlık türü, başlık satırı; docx_udf.py'nin zemin bandı taklidi gerçek hücre dolgusuna döner
  - görseller: UDF'deki dosya yeniden kodlanmadan, bayt bayt gömülür (kalite kaybı eklenmez)
  - üstbilgi / altbilgi (yalnız ilk sayfa, ikinci sayfadan itibaren), sayfa numarası alanı
  - antet (bgImage): üstbilgiye bağlı, metnin arkasında tam sayfa görsel
  - sayfa sonu, kâğıt boyutu ve yönü, kenar boşlukları

Editör'ün kendi yazdığı, bu aracın üretmediği öğeler (space, tab, field, barcode, bilinmeyenler)
metni korunarak okunur; karşılığı olmayan her şey UYARI olarak bildirilir. E-imzalı UDF'den
çıkan Word belgesi imza taşımaz; çalışma kopyasıdır.

Editör davranışları 21.09.2026'da kurulu UYAP Editör motorunda deneyle doğrulandı:
alt düzey liste görünümü NumberType'a değil SecListTypeLevelN'e (yoksa a. / i. / (1) / (a) / (i)
varsayılanına) bağlıdır; NumberSetted=n listeyi n+1'den başlatır; pageNumber-spec bir bit
kümesidir (32 ortala, 64 sağa, 2048 "sayfa/toplam"); bgImage kenar payları düşülerek sayfaya gerilir.
"""
import argparse
import base64
import hashlib
import re
import struct
import sys
import zipfile
from xml.etree import ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NSDECL = (f'xmlns:w="{W}" xmlns:r="{R}" '
          'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
          'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
          'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')
XMLDECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.'

IMG_CHAR = '¸'                        # Editör'ün görsel yer tutucusu
NBSP = '\u00a0'
DEFAULT_FONT = 'Times New Roman'
DEFAULT_SIZE = 12.0
# Editör kâğıt kodu -> (ad, genişlik, yükseklik) twip, dikey
MEDIA = {1: ('A4', 11906, 16838), 2: ('A5', 8391, 11906), 3: ('Letter', 12240, 15840),
         4: ('B5', 10318, 14570), 5: ('Legal', 12240, 20160), 6: ('Executive', 10440, 15120)}
ALIGN = {'0': 'left', '1': 'center', '2': 'right', '3': 'both'}
TAB_ALIGN = {0: 'left', 1: 'right', 2: 'center', 4: 'decimal', 5: 'bar'}
TAB_LEADER = {1: 'dot', 2: 'hyphen', 3: 'underscore', 4: 'heavy'}
CHAR_KEYS = ('bold', 'italic', 'underline', 'strikethrough', 'superscript', 'subscript',
             'size', 'family', 'foreground', 'background')
LEAF_TEXT = ('content', 'tab', 'space', 'field')       # metni aynen taşınan yaprak öğeler
NUM_INDENT = 6.0                      # Editör numarayı metnin ~6 punto soluna, sağa yaslı çizer
BULLET_INDENT = 18.0
CELL_PAD_MAX = 14.2                   # bundan büyük ortak girinti hücre payı sayılmaz
CELL_VPAD_MAX = 6.0                   # hücrelerin ortak üst/alt boşluğu: en çok bu kadarı hücre payı

# Editör'ün alt düzey varsayılanları (SecListTypeLevelN yoksa; motorla çizdirilerek görüldü)
DEF_NUM = {2: 'NUMBER_TYPE_CHAR_SMALL_DOT', 3: 'NUMBER_TYPE_ROMAN_SMALL_DOT',
           4: 'NUMBER_TYPE_NUMBER_D_PARANTHESE', 5: 'NUMBER_TYPE_CHAR_SMALL_D_PARANTHESE',
           6: 'NUMBER_TYPE_ROMAN_SMALL_D_PARANTHESE'}
DEF_BUL = {2: 'BULLET_TYPE_RECTANGLE', 3: 'BULLET_TYPE_RECTANGLE_D'}
NUM_KIND = (('NUMBER_TYPE_NUMBER', 'decimal'), ('NUMBER_TYPE_CHAR_BIG', 'upperLetter'),
            ('NUMBER_TYPE_CHAR_SMALL', 'lowerLetter'), ('NUMBER_TYPE_ROMAN_BIG', 'upperRoman'),
            ('NUMBER_TYPE_ROMAN_SMALL', 'lowerRoman'))
# işaret -> (karakter, yazı tipi); docx_udf.py'nin geri tanıdığı karakterler seçildi
BULLETS = {'ELLIPSE': ('\uf0b7', 'Symbol'), 'ARROW': ('\uf0d8', 'Wingdings'),
           'DIAMOND': ('\uf075', 'Wingdings'), 'DIAMOND_2': ('\uf075', 'Wingdings'),
           'RECTANGLE': ('\uf0a7', 'Wingdings'), 'RECTANGLE_D': ('\uf06f', 'Wingdings'),
           'TRIANGLE': ('\u25b6', None)}

# Editör harfli listeyi 29 harfli Türk alfabesiyle sayar, sonra aa, bb... (21.09.2026'da motorla görüldü)
TR_LETTERS = 'abcçdefgğhıijklmnoöprsştuüvyz'
TR_UPPER = 'ABCÇDEFGĞHIİJKLMNOÖPRSŞTUÜVYZ'
LIT_BULLET = {'ELLIPSE': '\u2022', 'ARROW': '\u27a2', 'DIAMOND': '\u25c6', 'DIAMOND_2': '\u25c6',
              'RECTANGLE': '\u25aa', 'RECTANGLE_D': '\u25ab', 'TRIANGLE': '\u25b6'}
LITERAL_HANG = 22.0                   # sabit metin etiketli listede asılı girinti (punto)

_BAD = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]')


class DonusumHatasi(Exception):
    """Kullanıcıya gösterilecek, anlaşılır dönüştürme hatası."""


def esc(s):
    return (_BAD.sub('', s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            .replace('"', '&quot;'))


def _f(v, default=0.0):
    try:
        return float(str(v).replace(',', '.'))
    except (TypeError, ValueError):
        return default


def _true(a, key):
    return str(a.get(key, '')).lower() == 'true'


def _hex(java_int):
    """Java işaretli renk tam sayısı -> 'RRGGBB' (çözülemezse None)."""
    try:
        return '%06X' % (int(java_int) & 0xFFFFFF)
    except (TypeError, ValueError):
        return None


def _roman(n):
    out = ''
    for v, r in ((1000, 'm'), (900, 'cm'), (500, 'd'), (400, 'cd'), (100, 'c'), (90, 'xc'), (50, 'l'),
                 (40, 'xl'), (10, 'x'), (9, 'ix'), (5, 'v'), (4, 'iv'), (1, 'i')):
        while n >= v:
            out += r
            n -= v
    return out


def label(typ, n):
    """Editör'ün bir liste kalemi için çizdiği etiket (tür + sıra numarası)."""
    if typ.startswith('BULLET'):
        return LIT_BULLET.get(typ.replace('BULLET_TYPE_', ''), '\u2022')
    n = max(1, n)
    if typ.startswith('NUMBER_TYPE_CHAR'):
        abc = TR_UPPER if 'CHAR_BIG' in typ else TR_LETTERS
        val = abc[(n - 1) % len(abc)] * ((n - 1) // len(abc) + 1)
    elif typ.startswith('NUMBER_TYPE_ROMAN'):
        val = _roman(n).upper() if 'ROMAN_BIG' in typ else _roman(n)
    else:
        val = str(n)
    tail = typ.rsplit('_TYPE_', 1)[-1]
    if tail.endswith('D_PARANTHESE'):
        return f'({val})'
    if 'PARANTHESE' in tail:
        return f'{val})'
    if tail.endswith('TRE'):
        return f'{val}-'
    return f'{val}.'


def tw(pt):
    return int(round(pt * 20))


def img_kind(data):
    """(uzantı, içerik türü, (px_w, px_h) ya da None)"""
    if data[:8] == b'\x89PNG\r\n\x1a\n':
        return 'png', 'image/png', (struct.unpack('>II', data[16:24]) if len(data) >= 24 else None)
    if data[:2] == b'\xff\xd8':
        return 'jpeg', 'image/jpeg', None
    if data[:6] in (b'GIF87a', b'GIF89a'):
        return 'gif', 'image/gif', (struct.unpack('<HH', data[6:10]) if len(data) >= 10 else None)
    if data[:2] == b'BM':
        return 'bmp', 'image/bmp', None
    return None, None, None


class Part:
    """Bir DOCX parçası (document / headerN / footerN): kendi ilişki listesi vardır."""

    def __init__(self, pkg, name):
        self.pkg, self.name, self.rels = pkg, name, []

    def rel(self, typ, target):
        rid = f'rId{len(self.rels) + 1}'
        self.rels.append((rid, typ, target))
        return rid

    def image(self, data):
        return self.rel('image', self.pkg.media(data))

    def rels_xml(self):
        return (XMLDECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                + ''.join(f'<Relationship Id="{i}" Type="{REL}{t}" Target="{g}"/>' for i, t, g in self.rels)
                + '</Relationships>')


class Package:
    def __init__(self):
        self.files, self.exts, self._media, self.overrides = {}, {}, {}, {}

    def media(self, data):
        key = hashlib.sha1(data).hexdigest()
        if key not in self._media:
            ext, ctype, _ = img_kind(data)
            name = f'media/image{len(self._media) + 1}.{ext}'
            self._media[key] = name
            self.files['word/' + name] = data
            self.exts[ext] = ctype
        return self._media[key]

    def add(self, name, xml, ctype=None):
        self.files[name] = xml.encode('utf-8') if isinstance(xml, str) else xml
        if ctype:
            self.overrides['/' + name] = ctype

    def write(self, path):
        types = (XMLDECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                 '<Default Extension="xml" ContentType="application/xml"/>'
                 + ''.join(f'<Default Extension="{e}" ContentType="{c}"/>' for e, c in sorted(self.exts.items()))
                 + ''.join(f'<Override PartName="{n}" ContentType="{c}"/>' for n, c in sorted(self.overrides.items()))
                 + '</Types>')
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', types.encode('utf-8'))
            for name, data in self.files.items():
                z.writestr(name, data)


class Cevirici:
    def __init__(self, path):
        self.warn = set()
        try:
            z = zipfile.ZipFile(path)
        except FileNotFoundError:
            raise DonusumHatasi(f'dosya bulunamadı: {path}')
        except zipfile.BadZipFile:
            raise DonusumHatasi('dosya bir UDF (zip) arşivi değil ya da bozuk.')
        names = z.namelist()
        if 'content.xml' not in names:
            raise DonusumHatasi('arşivde content.xml yok; bu bir UYAP UDF belgesi değil.')
        raw = z.read('content.xml')
        try:
            try:
                self.root = ET.fromstring(raw)
            except ET.ParseError:
                txt = _BAD.sub(' ', raw.decode('utf-8', 'replace')).lstrip('\ufeff \n\r\t')
                if txt.startswith('<?xml'):
                    txt = txt.split('?>', 1)[1]
                self.root = ET.fromstring(txt.strip())
        except ET.ParseError as ex:
            raise DonusumHatasi(f'content.xml okunamadı (bozuk XML): {ex}')
        self.extra = [n for n in names if n != 'content.xml']
        self.signed = any(re.search(r'sign|imza|\.sgn$|\.p7s$', n, re.I) for n in self.extra)
        if self.signed:
            self.warn.add('bu UDF e-imzalı; üretilen Word belgesi imza taşımaz (çalışma kopyasıdır)')
        c = self.root.find('content')
        self.cdata = (c.text if c is not None else '') or ''
        self.cd16 = self.cdata.encode('utf-16-le')
        self.els = self.root.find('elements')
        if self.els is None:
            raise DonusumHatasi('UDF içinde <elements> bölümü yok; belge boş ya da tanınmayan sürüm.')
        self.styles = {s.get('name'): dict(s.attrib) for s in self.root.iter('style') if s.get('name')}
        self._chain = {}
        base = self.chain(self.els.get('resolver') or 'hvl-default')
        self.base_font = base.get('family') or DEFAULT_FONT
        self.base_size = _f(base.get('size'), DEFAULT_SIZE) or DEFAULT_SIZE
        if self.base_font == 'Dialog':                 # Editör'ün arayüz varsayılanı; belgede TNR çizilir
            self.base_font = DEFAULT_FONT
        self.pkg = Package()
        self.doc = Part(self.pkg, 'document')
        self.pic_id = 0
        self.stats = {'paragraf': 0, 'tablo': 0, 'gorsel': 0, 'sayfa_sonu': 0}
        self.absorbed = 0                    # gerçek dikey birleştirmeye açılan iç tablo sayısı
        self._page()
        self._lists_scan()

    # ---------- okuma yardımcıları ----------
    def chain(self, name):
        """Stil zinciri (resolver -> üst stil) birleştirilmiş öznitelikler; çocuk üsttekini ezer."""
        key = name
        if key in self._chain:
            return self._chain[key]
        order, seen = [], set()
        while name and name in self.styles and name not in seen:
            seen.add(name)
            order.append(self.styles[name])
            name = self.styles[name].get('resolver')
        out = {}
        for st in reversed(order):
            out.update({k: v for k, v in st.items() if k not in ('name', 'resolver', 'description')})
        self._chain[key] = out
        return out

    def sl(self, el):
        try:
            so, ln = int(el.get('startOffset')), int(el.get('length'))
        except (TypeError, ValueError):
            return ''
        if ln <= 0 or so < 0:
            return ''
        return self.cd16[2 * so:2 * (so + ln)].decode('utf-16-le', 'replace')

    def para_attrs(self, p):
        a = dict(self.chain(p.get('resolver') or self.els.get('resolver') or 'hvl-default'))
        a.update(p.attrib)
        return a

    def char_attrs(self, leaf, pa):
        """Swing çözümlemesi: yaprağın kendi resolver'ı varsa o stil zinciri, yoksa paragraf."""
        a = dict(self.chain(leaf.get('resolver'))) if leaf.get('resolver') in self.styles else \
            {k: v for k, v in pa.items() if k in CHAR_KEYS}
        a.update({k: v for k, v in leaf.attrib.items() if k in CHAR_KEYS})
        return a

    @staticmethod
    def leaves(p):
        """Metne işaret eden gerçek yapraklar (sarmalayıcı öğenin içindekiler dâhil), belge sırasıyla."""
        out = []
        for c in p.iter():
            if c is p or c.get('startOffset') is None:
                continue
            if any(d is not c and d.get('startOffset') is not None for d in c.iter()):
                continue
            out.append(c)
        return out

    # ---------- sayfa ----------
    def _page(self):
        pf = self.root.find('properties/pageFormat')
        a = pf.attrib if pf is not None else {}
        try:
            code = int(a.get('mediaSizeName') or 1)
        except ValueError:
            code = 1
        if code not in MEDIA:
            self.warn.add(f'tanınmayan kâğıt kodu ({code}); A4 kullanıldı')
            code = 1
        _, pw, ph = MEDIA[code]
        self.landscape = str(a.get('paperOrientation', '1')) == '0'
        if self.landscape:
            pw, ph = ph, pw
        self.page_w, self.page_h = pw, ph
        self.margins = tuple(_f(a.get(k), d) for k, d in (('leftMargin', 70.85), ('rightMargin', 70.85),
                                                          ('topMargin', 70.85), ('bottomMargin', 70.85)))
        self.hf_off = (_f(a.get('headerFOffset'), 20.0), _f(a.get('footerFOffset'), 35.4))
        self.text_w = max(72.0, pw / 20.0 - self.margins[0] - self.margins[1])

    # ---------- listeler ----------
    def _lists_scan(self):
        """Belgedeki her ListId için düzey -> Editör türü (Editör'ün çizdiği kurala göre)."""
        sec, own = {}, {}
        for p in self.root.iter('paragraph'):
            lid = p.get('ListId')
            if lid is None or not (_true(p.attrib, 'Numbered') or _true(p.attrib, 'Bulleted')):
                continue
            lvl = self._level(p)
            v = p.get(f'SecListTypeLevel{lvl}')           # Editör yalnız paragrafın KENDİ düzeyininkine bakar
            if v:
                sec.setdefault(lid, {}).setdefault(lvl, v)
            typ = p.get('NumberType') if _true(p.attrib, 'Numbered') else p.get('BulletType') or 'BULLET_TYPE_ELLIPSE'
            own.setdefault(lid, {}).setdefault(lvl, typ or 'NUMBER_TYPE_NUMBER_DOT')
        # Editör'ün sayacı: aynı ListId sürer, sığ bir kalem gelince derin düzeyler yeniden başlar
        self.ordinal, counters, top = {}, {}, {}
        for p in self.root.iter('paragraph'):
            lid = p.get('ListId')
            if lid is None or not (_true(p.attrib, 'Numbered') or _true(p.attrib, 'Bulleted')):
                continue
            lvl = self._level(p)
            c = counters.setdefault(lid, {})
            for k in [k for k in c if k > lvl]:
                del c[k]
            if p.get('NumberSetted') not in (None, ''):
                c[lvl] = int(_f(p.get('NumberSetted')))
            c[lvl] = c.get(lvl, 0) + 1
            self.ordinal[id(p)] = c[lvl]
            top[(lid, lvl)] = max(top.get((lid, lvl), 0), c[lvl])
        self.lists = {}
        for lid, lv in own.items():
            levels = {}
            bulleted = lv.get(1, next(iter(lv.values()))).startswith('BULLET')
            for n in range(1, 10):
                if n in sec.get(lid, {}):
                    levels[n] = sec[lid][n]
                elif n == 1:
                    levels[n] = lv.get(1) or ('BULLET_TYPE_ELLIPSE' if bulleted else 'NUMBER_TYPE_NUMBER_DOT')
                else:
                    kind_b = lv[n].startswith('BULLET') if n in lv else bulleted
                    levels[n] = (DEF_BUL.get(n, 'BULLET_TYPE_RECTANGLE') if kind_b
                                 else DEF_NUM.get(n, 'NUMBER_TYPE_NUMBER_DOT'))
            # Word harfleri İngiliz alfabesiyle sayar (…c, d…), Editör Türk alfabesiyle (…c, ç, d…):
            # 4. kaleme ulaşan harfli düzey varsa etiketler kaymasın diye bütün liste sabit metin olur
            literal = any(levels[n].startswith('NUMBER_TYPE_CHAR') and top.get((lid, n), 0) >= 4 for n in levels)
            self.lists[lid] = {'levels': levels, 'abs': len(self.lists) + 1, 'literal': literal}
        self.nums = []                       # (numId, absId, {ilvl: start})
        self.cur_num = {}

    @staticmethod
    def _level(p):
        try:
            return min(9, max(1, int(p.get('ListLevel') or 1)))
        except ValueError:
            return 1

    def num_id(self, p):
        lid = p.get('ListId')
        spec = self.lists.get(lid)
        if not spec:
            return None
        lvl = self._level(p)
        setted = p.get('NumberSetted')
        start = None
        if setted not in (None, ''):
            try:
                start = int(float(setted)) + 1           # Editör: NumberSetted=n -> n+1'den başlar
            except ValueError:
                start = None
        if lid not in self.cur_num or start is not None:
            nid = len(self.nums) + 1
            self.nums.append((nid, spec['abs'], {lvl - 1: start} if start is not None else {}))
            self.cur_num[lid] = nid
        return self.cur_num[lid]

    def numbering_xml(self):
        out = [XMLDECL, f'<w:numbering {NSDECL}>']
        for lid, spec in self.lists.items():
            if spec['literal']:
                continue
            out.append(f'<w:abstractNum w:abstractNumId="{spec["abs"]}"><w:multiLevelType w:val="hybridMultilevel"/>')
            for n in range(1, 10):
                typ = spec['levels'][n]
                left = tw(36.0 * n)
                if typ.startswith('BULLET'):
                    ch, font = BULLETS.get(typ.replace('BULLET_TYPE_', ''), BULLETS['ELLIPSE'])
                    rpr = (f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}" w:hint="default"/></w:rPr>'
                           if font else '')
                    out.append(f'<w:lvl w:ilvl="{n - 1}"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
                               f'<w:lvlText w:val="{ch}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="{left}" '
                               f'w:hanging="{tw(BULLET_INDENT)}"/></w:pPr>{rpr}</w:lvl>')
                    continue
                fmt = next((f for k, f in NUM_KIND if typ.startswith(k)), 'decimal')
                tail = typ.rsplit('_TYPE_', 1)[-1]
                if tail.endswith('D_PARANTHESE'):
                    text = f'(%{n})'
                elif 'PARANTHESE' in tail:
                    text = f'%{n})'
                elif tail.endswith('TRE'):
                    text = f'%{n}-'
                else:
                    text = f'%{n}.'
                out.append(f'<w:lvl w:ilvl="{n - 1}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
                           f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="right"/><w:pPr><w:ind w:left="{left}" '
                           f'w:hanging="{tw(NUM_INDENT)}"/></w:pPr></w:lvl>')
            out.append('</w:abstractNum>')
        for nid, aid, over in self.nums:
            ov = ''.join(f'<w:lvlOverride w:ilvl="{i}"><w:startOverride w:val="{s}"/></w:lvlOverride>'
                         for i, s in over.items())
            out.append(f'<w:num w:numId="{nid}"><w:abstractNumId w:val="{aid}"/>{ov}</w:num>')
        out.append('</w:numbering>')
        return ''.join(out)

    # ---------- yazı ----------
    def rpr(self, a, skip_bg=None):
        x = ''
        fam = a.get('family')
        if fam == 'Dialog':
            fam = None
        if fam and fam != self.base_font:
            f = esc(fam)
            x += f'<w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{f}"/>'
        if _true(a, 'bold'):
            x += '<w:b/><w:bCs/>'
        if _true(a, 'italic'):
            x += '<w:i/><w:iCs/>'
        if _true(a, 'strikethrough'):
            x += '<w:strike/>'
        col = _hex(a.get('foreground')) if a.get('foreground') not in (None, '') else None
        if col and col != '000000':
            x += f'<w:color w:val="{col}"/>'
        size = _f(a.get('size'), self.base_size) or self.base_size
        if abs(size - self.base_size) > 0.01:
            hp = max(2, int(round(size * 2)))
            x += f'<w:sz w:val="{hp}"/><w:szCs w:val="{hp}"/>'
        if _true(a, 'underline'):
            x += '<w:u w:val="single"/>'
        bg = _hex(a.get('background')) if a.get('background') not in (None, '') else None
        if bg and bg != skip_bg and bg != 'FFFFFF':
            x += f'<w:shd w:val="clear" w:color="auto" w:fill="{bg}"/>'
        if _true(a, 'superscript'):
            x += '<w:vertAlign w:val="superscript"/>'
        elif _true(a, 'subscript'):
            x += '<w:vertAlign w:val="subscript"/>'
        return f'<w:rPr>{x}</w:rPr>' if x else ''

    @staticmethod
    def run(text, rpr):
        inner = ''
        for part in re.split('(\t|\n)', text):
            if part == '\t':
                inner += '<w:tab/>'
            elif part == '\n':
                inner += '<w:br/>'
            elif part:
                inner += f'<w:t xml:space="preserve">{esc(part)}</w:t>'
        return f'<w:r>{rpr}{inner}</w:r>' if inner else ''

    def drawing(self, part, data, w_pt, h_pt, anchor=None):
        self.pic_id += 1
        n = self.pic_id
        rid = part.image(data)
        cx, cy = int(w_pt * 12700), int(h_pt * 12700)
        pic = ('<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
               f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{n}" name="Gorsel {n}"/><pic:cNvPicPr/></pic:nvPicPr>'
               f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
               f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
               '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>')
        if anchor is None:
            return ('<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
                    f'<wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{n}" name="Gorsel {n}"/>'
                    '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
                    f'{pic}</wp:inline></w:drawing></w:r>')
        x, y = int(anchor[0] * 12700), int(anchor[1] * 12700)
        return ('<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
                'relativeHeight="0" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">'
                '<wp:simplePos x="0" y="0"/>'
                f'<wp:positionH relativeFrom="page"><wp:posOffset>{x}</wp:posOffset></wp:positionH>'
                f'<wp:positionV relativeFrom="page"><wp:posOffset>{y}</wp:posOffset></wp:positionV>'
                f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
                f'<wp:docPr id="{n}" name="Antet {n}"/><wp:cNvGraphicFramePr/>{pic}</wp:anchor></w:drawing></w:r>')

    # ---------- paragraf ----------
    def tabs_xml(self, tabset, shift, own):
        """UDF TabSet -> w:tabs. Editör konumu paragrafın LeftIndent'inden, Word sayfa kenarından
        ölçer: konuma LeftIndent eklenir. Sondaki, Word'ün varsayılan 36'lık ızgarasına düşen sola
        hizalı duraklar yazılmaz (Word onları kendisi koyar)."""
        if not tabset or not own and shift < 0.05:
            return ''
        stops = []
        for item in tabset.split(','):
            bits = item.strip().split(':')
            if not bits[0]:
                continue
            pos = _f(bits[0], None)
            if pos is None:
                continue
            al = int(_f(bits[1], 0)) if len(bits) > 1 else 0
            ld = int(_f(bits[2], 0)) if len(bits) > 2 else 0
            stops.append((pos + shift, al, ld))
        stops.sort()
        while stops:
            pos, al, ld = stops[-1]
            if al == 0 and ld == 0 and abs(pos / 36.0 - round(pos / 36.0)) < 0.005:
                stops.pop()
            else:
                break
        if not stops:
            return ''
        x = ''
        for pos, al, ld in stops:
            lead = f' w:leader="{TAB_LEADER[ld]}"' if ld in TAB_LEADER else ''
            x += f'<w:tab w:val="{TAB_ALIGN.get(al, "left")}"{lead} w:pos="{tw(pos)}"/>'
        return f'<w:tabs>{x}</w:tabs>'

    def para(self, p, ctx):
        """UDF <paragraph> -> <w:p>. ctx: part, pad (hücre payı), fill (hücre dolgusu), avail."""
        part = ctx['part']
        pa = self.para_attrs(p)
        pad_l, pad_r = ctx.get('pad', (0.0, 0.0))
        fill = ctx.get('fill')
        items = []                                   # ('t', metin, öznitelikler) | ('img', öğe)
        for leaf in self.leaves(p):
            if leaf.tag == 'image':
                items.append(['img', leaf, None])
                continue
            text = self.sl(leaf)
            if leaf.tag == 'barcode':
                self.warn.add('barkod taşınmadı')
                text = text.replace(IMG_CHAR, '')
            elif leaf.tag == 'field':
                self.warn.add('form alanları düz metne çevrildi')
            elif leaf.tag not in LEAF_TEXT:
                self.warn.add(f'tanınmayan öğe <{leaf.tag}>: metni korundu, biçimi taşınmadı')
            items.append(['t', text, self.char_attrs(leaf, pa)])
        mark = items[-1][2] if items and items[-1][0] == 't' else \
            next((i[2] for i in reversed(items) if i[0] == 't'), {k: v for k, v in pa.items() if k in CHAR_KEYS})
        for it in reversed(items):                   # paragraf sonu '\n' metne girmez
            if it[0] == 't' and it[1].endswith('\n'):
                it[1] = it[1][:-1]
                break
        banded = False
        if fill:                                     # docx_udf'nin zemin bandı: baştaki/sondaki NBSP dilimleri
            texts = [i for i in items if i[0] == 't' and i[1]]
            for seq in (texts, texts[::-1]):
                for it in seq:
                    if it[1].strip(NBSP) == '' and _hex(it[2].get('background')) == fill:
                        it[1] = ''
                        banded = True
                    else:
                        break
        runs = ''
        for kind, val, a in items:
            if kind == 't':
                runs += self.run(val, self.rpr(a, fill))
                continue
            try:
                data = base64.b64decode(re.sub(r'\s+', '', val.get('imageData') or ''))
            except (ValueError, TypeError):
                data = b''
            ext, _, px = img_kind(data)
            if not ext:
                self.warn.add('okunamayan bir görsel atlandı')
                continue
            w_pt, h_pt = _f(val.get('width')), _f(val.get('height'))
            if w_pt <= 0 or h_pt <= 0:
                w_pt, h_pt = ((px[0] * 0.75, px[1] * 0.75) if px else (120.0, 90.0))
            avail = ctx.get('avail') or self.text_w
            if w_pt > avail:
                w_pt, h_pt = avail, h_pt * avail / w_pt
            runs += self.drawing(part, data, w_pt, h_pt)
            self.stats['gorsel'] += 1

        li = _f(pa.get('LeftIndent'))
        ri = _f(pa.get('RightIndent'))
        if banded:
            li, ri = 0.0, 0.0
        else:
            li, ri = max(0.0, li - pad_l), max(0.0, ri - pad_r)
        hang, first = _f(pa.get('Hanging')), _f(pa.get('FirstLineIndent'))
        ppr = ''
        if ctx.pop('page_break', False):
            ppr += '<w:pageBreakBefore/>'
        is_list = (_true(p.attrib, 'Numbered') or _true(p.attrib, 'Bulleted')) and p.get('ListId') in self.lists
        literal = is_list and self.lists[p.get('ListId')]['literal']
        if literal:
            typ = self.lists[p.get('ListId')]['levels'][self._level(p)]
            la = dict(next((a for k, _v, a in items if k == 't' and a), mark))
            for k in ('underline', 'strikethrough', 'superscript', 'subscript', 'background'):
                la.pop(k, None)
            runs = self.run(label(typ, self.ordinal.get(id(p), 1)) + '\t', self.rpr(la)) + runs
            self.warn.add('harf sıralı liste Editör\'deki etiketleriyle (…c, ç, d…) sabit metin olarak yazıldı; '
                          'Word\'ün otomatik harf sırası (…c, d, e…) farklı olduğu için otomatik numaralandırma kullanılmadı')
        elif is_list:
            ppr += (f'<w:numPr><w:ilvl w:val="{self._level(p) - 1}"/>'
                    f'<w:numId w:val="{self.num_id(p)}"/></w:numPr>')
        if any(k == 't' and '\t' in v for k, v, _ in items):
            ppr += self.tabs_xml(pa.get('TabSet'), li, 'TabSet' in p.attrib)
        before, after = _f(pa.get('SpaceAbove')), _f(pa.get('SpaceBelow'))
        vt, vb = ctx.get('vpad', (0.0, 0.0))
        before, after = max(0.0, before - vt), max(0.0, after - vb)
        line = int(round(240 * (1 + _f(pa.get('LineSpacing')))))
        ppr += (f'<w:spacing w:before="{tw(before)}" w:after="{tw(after)}" '
                f'w:line="{max(line, 120)}" w:lineRule="auto"/>')
        if literal:
            ind = f' w:left="{tw(li)}" w:hanging="{tw(LITERAL_HANG)}"'
        elif is_list:
            # Editör numarayı LeftIndent'in soluna çizer, metin LeftIndent'te başlar
            bul = self.lists[p.get('ListId')]['levels'][self._level(p)].startswith('BULLET')
            ind = f' w:left="{tw(li)}" w:hanging="{tw(BULLET_INDENT if bul else NUM_INDENT)}"'
        else:
            # UDF: ilk satır = LeftIndent + FirstLineIndent, devam = LeftIndent + Hanging
            ind = f' w:left="{tw(li + hang)}"'
            delta = first - hang
            if delta > 0.05:
                ind += f' w:firstLine="{tw(delta)}"'
            elif delta < -0.05:
                ind += f' w:hanging="{tw(-delta)}"'
        if ri > 0.05:
            ind += f' w:right="{tw(ri)}"'
        if ind.strip() != 'w:left="0"':
            ppr += f'<w:ind{ind}/>'
        jc = ALIGN.get(str(pa.get('Alignment', '0')).split('.')[0], 'left')
        if jc != 'left':
            ppr += f'<w:jc w:val="{jc}"/>'
        ppr += self.rpr(mark, fill)
        self.stats['paragraf'] += 1
        return f'<w:p><w:pPr>{ppr}</w:pPr>{runs}</w:p>'

    # ---------- tablo ----------
    def cell_fill(self, cell):
        """Hücredeki bütün yazı aynı zemin rengini taşıyorsa o renk (gerçek hücre dolgusuna çevrilir).
        Hücrede başka renkle vurgulanmış bir kelime olsa da docx_udf'nin zemin bandı (paragrafın iki
        ucunda aynı renkli NBSP dilimleri) hücre rengini gösterir; yoksa bant NBSP'leri Word'e sızıyordu."""
        colors, band, uniform = set(), set(), True
        for p in cell.findall('paragraph'):
            pa = self.para_attrs(p)
            lv = [l for l in self.leaves(p) if l.tag != 'image' and self.sl(l).strip('\n')]
            if lv and all(self.sl(l).strip(NBSP + '\n') == '' for l in (lv[0], lv[-1])):
                ca, cb = (_hex(self.char_attrs(l, pa).get('background') or '') for l in (lv[0], lv[-1]))
                if ca and ca == cb and ca != 'FFFFFF':
                    band.add(ca)
            for leaf in self.leaves(p):
                if leaf.tag == 'image' or not self.sl(leaf).strip(' \t\n'):
                    continue
                bg = self.char_attrs(leaf, pa).get('background')
                col = _hex(bg) if bg not in (None, '') else None
                if not col or col == 'FFFFFF':
                    uniform = False
                else:
                    colors.add(col)
        if uniform and len(colors) == 1:
            return colors.pop()
        return band.pop() if len(band) == 1 else None

    @staticmethod
    def _fracs(row, cells, tspans):
        sp = [_f(x, 1.0) or 1.0 for x in (row.get('columnSpans') or '').split(',') if x.strip()]
        if len(sp) != len(cells):
            sp = tspans if len(tspans) == len(cells) else [1.0] * max(1, len(cells))
        tot = sum(sp) or 1.0
        return [x / tot for x in sp]

    @staticmethod
    def _only_table(cell):
        kids = [k for k in cell if k.tag in ('paragraph', 'table')]
        return kids[0] if len(kids) == 1 and kids[0].tag == 'table' else None

    def _logical_rows(self, t):
        """UDF satırları -> Word satırları: [(başlık_mı, [(oran, hücre, vmerge)])].

        docx_udf.py dikey birleştirmeyi iç içe tabloyla taklit eder: dış satırda birleşik sütun düz
        hücre, aradaki sütun blokları yalnız bir iç tablo taşıyan hücredir. Bu kalıp Word'ün gerçek
        dikey birleştirmesine (w:vMerge) geri açılır; kalıba uymayan iç tablo Word'de iç içe kalır."""
        tspans = [_f(x, 1.0) or 1.0 for x in (t.get('columnSpans') or '').split(',') if x.strip()]
        out = []
        for r in t.findall('row'):
            cells = r.findall('cell')
            if not cells:
                continue
            fr = self._fracs(r, cells, tspans)
            head = r.get('rowType') == 'headerRow'
            inner = [self._only_table(c) for c in cells]
            irows = [[x for x in n.findall('row') if x.findall('cell')] if n is not None else None for n in inner]
            ks = {len(x) for x in irows if x is not None}
            if any(n is not None for n in inner) and any(n is None for n in inner) and len(ks) == 1 \
                    and min(ks) >= 2:
                for j in range(min(ks)):
                    row = []
                    for c, f, n, rs in zip(cells, fr, inner, irows):
                        if n is None:
                            row.append((f, c, 'restart' if j == 0 else 'continue'))
                            continue
                        ics = rs[j].findall('cell')
                        ispans = [_f(x, 1.0) or 1.0 for x in (n.get('columnSpans') or '').split(',') if x.strip()]
                        for ic, iff in zip(ics, self._fracs(rs[j], ics, ispans)):
                            row.append((f * iff, ic, None))
                    out.append((head, row))
                self.absorbed += sum(1 for n in inner if n is not None)
            else:
                out.append((head, [(f, c, None) for c, f in zip(cells, fr)]))
        return out

    def table(self, t, ctx):
        part, avail = ctx['part'], ctx.get('avail') or self.text_w
        rows = self._logical_rows(t)
        if not rows:
            return ''
        bounds = [0.0, 1.0]
        for _, row in rows:
            acc = 0.0
            for f, _c, _v in row[:-1]:
                acc += f
                if all(abs(acc - b) > 0.004 for b in bounds):
                    bounds.append(acc)
        bounds.sort()
        near = lambda v: min(range(len(bounds)), key=lambda i: abs(bounds[i] - v))

        real = [c for _, row in rows for _f2, c, v in row if v != 'continue']
        fills = {id(c): self.cell_fill(c) for c in real}
        lefts, rights, tops, bottoms = [], [], [], []
        for c in real:
            ps = c.findall('paragraph')
            if ps and self._only_table(c) is None and not c.findall('table'):
                tops.append(_f(self.para_attrs(ps[0]).get('SpaceAbove')))
                bottoms.append(_f(self.para_attrs(ps[-1]).get('SpaceBelow')))
            for p in ps:
                if _true(p.attrib, 'Numbered') or _true(p.attrib, 'Bulleted'):
                    continue
                pa = self.para_attrs(p)
                if fills[id(c)] and self._is_banded(p, pa, fills[id(c)]):
                    continue
                lefts.append(_f(pa.get('LeftIndent')))
                rights.append(_f(pa.get('RightIndent')))
        pad_l = min(min(lefts), CELL_PAD_MAX) if lefts else 0.0
        pad_r = min(min(rights), CELL_PAD_MAX) if rights else 0.0
        pad_t = min(min(tops), CELL_VPAD_MAX) if tops else 0.0
        pad_b = min(min(bottoms), CELL_VPAD_MAX) if bottoms else 0.0

        border = t.get('border') or 'borderCell'
        b1 = ' w:val="single" w:sz="4" w:space="0" w:color="000000"'
        b0 = ' w:val="nil"'
        outer = b0 if border == 'borderNone' else b1
        in_h = b1 if border in ('borderCell', 'borderRow') else b0
        in_v = b1 if border == 'borderCell' else b0
        if border not in ('borderCell', 'borderTable', 'borderRow', 'borderNone'):
            self.warn.add(f'tanınmayan tablo kenarlık türü ({border}); tam kenarlık kullanıldı')
            outer = in_h = in_v = b1
        x = (f'<w:tbl><w:tblPr><w:tblW w:w="{tw(avail)}" w:type="dxa"/><w:tblBorders><w:top{outer}/>'
             f'<w:left{outer}/><w:bottom{outer}/><w:right{outer}/><w:insideH{in_h}/><w:insideV{in_v}/>'
             '</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar>'
             f'<w:top w:w="{tw(pad_t)}" w:type="dxa"/><w:left w:w="{tw(pad_l)}" w:type="dxa"/>'
             f'<w:bottom w:w="{tw(pad_b)}" w:type="dxa"/><w:right w:w="{tw(pad_r)}" w:type="dxa"/></w:tblCellMar>'
             '</w:tblPr><w:tblGrid>'
             + ''.join(f'<w:gridCol w:w="{tw((bounds[i + 1] - bounds[i]) * avail)}"/>'
                       for i in range(len(bounds) - 1)) + '</w:tblGrid>')
        merged_fill = {}
        for head, row in rows:
            x += '<w:tr>'
            if head:
                x += '<w:trPr><w:tblHeader/></w:trPr>'
            acc = 0.0
            for f, c, vm in row:
                i0, i1 = near(acc), near(acc + f)
                acc += f
                i0 = min(i0, len(bounds) - 2)
                i1 = min(max(i1, i0 + 1), len(bounds) - 1)
                cw = (bounds[i1] - bounds[i0]) * avail
                tcpr = f'<w:tcW w:w="{tw(cw)}" w:type="dxa"/>'
                if i1 - i0 > 1:
                    tcpr += f'<w:gridSpan w:val="{i1 - i0}"/>'
                if vm:
                    tcpr += '<w:vMerge w:val="restart"/>' if vm == 'restart' else '<w:vMerge/>'
                fill = fills.get(id(c)) if vm != 'continue' else merged_fill.get(i0)
                if vm == 'restart':
                    merged_fill[i0] = fill
                if fill:
                    tcpr += f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'
                if vm == 'continue':
                    x += f'<w:tc><w:tcPr>{tcpr}</w:tcPr><w:p/></w:tc>'
                    continue
                sub = {'part': part, 'pad': (pad_l, pad_r), 'fill': fill,
                       'avail': max(20.0, cw - pad_l - pad_r)}
                kids = [k for k in c if k.tag in ('paragraph', 'table')]
                plain = not c.findall('table')
                body, last = '', None
                for k in kids:
                    if k.tag == 'paragraph':
                        sub['vpad'] = ((pad_t if plain and k is kids[0] else 0.0),
                                       (pad_b if plain and k is kids[-1] else 0.0))
                        body += self.para(k, sub)
                    else:
                        if last == 'table':
                            body += self.spacer()
                        body += self.table(k, {'part': part, 'avail': sub['avail']})
                    last = k.tag
                if last != 'paragraph':              # Word: hücre paragrafla bitmek zorunda
                    body += self.spacer() if last == 'table' else '<w:p/>'
                x += f'<w:tc><w:tcPr>{tcpr}</w:tcPr>{body}</w:tc>'
            x += '</w:tr>'
        self.stats['tablo'] += 1
        return x + '</w:tbl>'

    def _is_banded(self, p, pa, fill):
        lv = [l for l in self.leaves(p) if l.tag != 'image' and self.sl(l).strip('\n')]
        if not lv:
            return False
        first = lv[0]
        return (self.sl(first).strip(NBSP + '\n') == ''
                and _hex(self.char_attrs(first, pa).get('background')) == fill) or \
               (self.sl(lv[-1]).strip(NBSP + '\n') == ''
                and _hex(self.char_attrs(lv[-1], pa).get('background')) == fill)

    @staticmethod
    def break_para():
        return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

    @staticmethod
    def spacer():
        """Word art arda gelen iki tabloyu birleştirir; araya görünmeyecek kadar küçük paragraf."""
        return ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
                '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>')

    # ---------- gövde ----------
    def blocks(self, container, ctx):
        out, last = '', None
        for el in container:
            if el.tag == 'paragraph':
                out += self.para(el, ctx)
            elif el.tag == 'table':
                if last == 'table':
                    out += self.spacer()
                if ctx.pop('page_break', False):
                    out += self.break_para()
                out += self.table(el, ctx)
            elif el.tag == 'page-break':
                inner = [p for p in el.iter('paragraph') if self._ptext(p).strip()]
                if ctx.get('page_break'):                # art arda iki sayfa sonu: ilki ayrı paragraf
                    out += self.break_para()
                ctx['page_break'] = True
                self.stats['sayfa_sonu'] += 1
                for p in inner:
                    out += self.para(p, ctx)
                last = 'page-break'
                continue
            elif el.tag in ('header', 'footer'):
                continue
            elif el.get('startOffset') is not None:          # paragrafsız yaprak: kendi paragrafı
                wrap = ET.Element('paragraph')
                wrap.append(el)
                out += self.para(wrap, ctx)
            else:
                self.warn.add(f'tanınmayan öğe <{el.tag}>: içindeki paragraflar sırayla aktarıldı')
                out += self.blocks(el, ctx)
            last = el.tag
        if ctx.get('top') and ctx.pop('page_break', False):
            out += self.break_para()
        return out

    def _ptext(self, p):
        return ''.join(self.sl(l) for l in self.leaves(p) if l.tag != 'image').replace('\n', '')

    # ---------- üstbilgi / altbilgi ----------
    def page_number_para(self, f):
        try:
            bits = int((f.get('pageNumber-spec') or 'BSP32_0').replace('BSP32_', '') or 0)
        except ValueError:
            bits = 0
        if not bits:
            return ''
        jc = 'center' if bits & 32 else 'right' if bits & 64 else 'left'
        a = {'family': f.get('pageNumber-fontFace'), 'size': f.get('pageNumber-fontSize'),
             'bold': f.get('pageNumber-fontBold', ''), 'italic': f.get('pageNumber-fontItalic', ''),
             'underline': f.get('pageNumber-fontUnderline', ''), 'foreground': f.get('pageNumber-color')}
        rpr = self.rpr({k: v for k, v in a.items() if v not in (None, '')})

        def field(name):
            return (f'<w:r>{rpr}<w:fldChar w:fldCharType="begin"/></w:r>'
                    f'<w:r>{rpr}<w:instrText xml:space="preserve"> {name} </w:instrText></w:r>'
                    f'<w:r>{rpr}<w:fldChar w:fldCharType="separate"/></w:r>'
                    f'<w:r>{rpr}<w:t>1</w:t></w:r><w:r>{rpr}<w:fldChar w:fldCharType="end"/></w:r>')
        x = self.run(f.get('pageNumber-foreStr') or '', rpr) + field('PAGE')
        if bits & 2048:
            x += self.run(f.get('pageNumber-seperator') or '/', rpr) + field('NUMPAGES')
        x += self.run(f.get('pageNumber-afterStr') or '', rpr)
        return (f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
                f'<w:jc w:val="{jc}"/>{rpr}</w:pPr>{x}</w:p>')

    def header_footer(self):
        """(sectPr için başvurular, titlePg gerekli mi). Editör tek üstbilgi/altbilgi tutar (sonuncusu)."""
        bg = self.root.find('properties/bgImage')
        bg_data = b''
        if bg is not None and (bg.get('bgImageData') or '').strip():
            try:
                bg_data = base64.b64decode(re.sub(r'\s+', '', bg.get('bgImageData')))
            except (ValueError, TypeError):
                bg_data = b''
            if not img_kind(bg_data)[0]:
                bg_data = b''
                self.warn.add('antet (arka plan görseli) okunamadı, taşınmadı')
        spec = {}
        for kind in ('header', 'footer'):
            found = self.els.findall(kind)
            if len(found) > 1:
                self.warn.add(f'birden çok {"üstbilgi" if kind == "header" else "altbilgi"} öğesi var; '
                              'Editör gibi yalnız sonuncusu alındı')
            el = found[-1] if found else None
            where = 'all'
            if el is not None:
                start, stop = el.get('startPage'), el.get('stopPage')
                if stop == '1' and not start:
                    where = 'first'
                elif start == '2' and not stop:
                    where = 'rest'
                elif start or stop:
                    self.warn.add(f'{"üstbilgi" if kind == "header" else "altbilgi"} sayfa aralığı '
                                  f'({start or "1"}–{stop or "son"}) Word\'de karşılanamadı; bütün sayfalara konuldu')
            spec[kind] = (el, where)
        title = any(w != 'all' for _, w in spec.values())
        refs, n = '', 0
        for kind in ('header', 'footer'):
            el, where = spec[kind]
            tag = 'hdr' if kind == 'header' else 'ftr'
            for typ in (('default', 'first') if title else ('default',)):
                show = el is not None and (where == 'all' or (where == 'first') == (typ == 'first'))
                want_bg = kind == 'header' and bg_data
                if not show and not want_bg:
                    continue
                n += 1
                part = Part(self.pkg, f'{kind}{n}')
                body = ''
                if show:
                    ctx = {'part': part, 'avail': self.text_w}
                    paras = list(el)
                    pn = self.page_number_para(el) if kind == 'footer' else ''
                    if pn:                           # Editör'ün numara için ayırdığı boş yer paragrafı
                        while paras and paras[-1].tag == 'paragraph' and not self._ptext(paras[-1]).strip() \
                                and not paras[-1].findall('image'):
                            paras.pop()
                    holder = ET.Element('x')
                    holder.extend(paras)
                    body = self.blocks(holder, ctx) + pn
                if want_bg:
                    m = [_f(bg.get(k)) for k in ('bgImageLeftMargin', 'bgImageUpMargin',
                                                 'bgImageRigtMargin', 'bgImageBottomMargin')]
                    w_pt = max(10.0, self.page_w / 20.0 - m[0] - m[2])
                    h_pt = max(10.0, self.page_h / 20.0 - m[1] - m[3])
                    anchor = self.drawing(part, bg_data, w_pt, h_pt, anchor=(m[0], m[1]))
                    if body.startswith('<w:p><w:pPr>'):
                        i = body.index('</w:pPr>') + len('</w:pPr>')
                        body = body[:i] + anchor + body[i:]
                    elif body.startswith('<w:p>'):
                        body = '<w:p>' + anchor + body[len('<w:p>'):]
                    else:                            # tabloyla başlıyor: antet için 1 puntoluk taşıyıcı paragraf
                        body = ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
                                f'</w:pPr>{anchor}</w:p>' + body)
                body = body or '<w:p/>'
                self.pkg.add(f'word/{kind}{n}.xml', f'{XMLDECL}<w:{tag} {NSDECL}>{body}</w:{tag}>',
                             CT + kind + '+xml')
                if part.rels:
                    self.pkg.add(f'word/_rels/{kind}{n}.xml.rels', part.rels_xml())
                rid = self.doc.rel(kind, f'{kind}{n}.xml')
                refs += f'<w:{kind}Reference w:type="{typ}" r:id="{rid}"/>'
        if spec['footer'][0] is None:
            # UDF'de altbilgi yok = sayfa numarası yok. Word'e boş altbilgi yazılır; yoksa belge geri
            # çevrilince "altbilgi hiç yok" sayılıp sayfa numarası eklenir (02.10.2026 gidiş dönüş testi).
            n += 1
            self.pkg.add(f'word/footer{n}.xml', f'{XMLDECL}<w:ftr {NSDECL}><w:p/></w:ftr>', CT + 'footer+xml')
            rid = self.doc.rel('footer', f'footer{n}.xml')
            refs += f'<w:footerReference w:type="default" r:id="{rid}"/>'
        if bg_data:
            self.warn.add('antet (arka plan görseli) Word\'de üstbilgiye bağlı, metnin arkasında görsel olarak yerleştirildi')
        start = ''
        f_el = spec['footer'][0]
        if f_el is not None and (f_el.get('pageNumber-pageStartNumStr') or '').strip().isdigit():
            start = f'<w:pgNumType w:start="{int(f_el.get("pageNumber-pageStartNumStr"))}"/>'
        return refs, title, start, spec

    # ---------- paket ----------
    def build(self, out_path):
        keep = dict(self.stats)
        refs, title, pgnum, spec = self.header_footer()
        self.stats = keep                            # sayım yalnız gövdeyi gösterir
        body = self.blocks(self.els, {'part': self.doc, 'avail': self.text_w, 'top': True})
        if not body.endswith('</w:p>'):              # Word: gövde tabloyla bitemez
            body += self.spacer()
        l, r, t, b = self.margins
        orient = ' w:orient="landscape"' if self.landscape else ''
        sect = (f'<w:sectPr>{refs}<w:pgSz w:w="{self.page_w}" w:h="{self.page_h}"{orient}/>'
                f'<w:pgMar w:top="{tw(t)}" w:right="{tw(r)}" w:bottom="{tw(b)}" w:left="{tw(l)}" '
                f'w:header="{tw(self.hf_off[0])}" w:footer="{tw(self.hf_off[1])}" w:gutter="0"/>'
                f'{pgnum}{"<w:titlePg/>" if title else ""}</w:sectPr>')
        self.pkg.add('word/document.xml',
                     f'{XMLDECL}<w:document {NSDECL}><w:body>{body}{sect}</w:body></w:document>',
                     CT + 'document.main+xml')
        f = esc(self.base_font)
        hp = int(round(self.base_size * 2))
        self.pkg.add('word/styles.xml',
                     f'{XMLDECL}<w:styles {NSDECL}><w:docDefaults><w:rPrDefault><w:rPr>'
                     f'<w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{f}"/>'
                     f'<w:sz w:val="{hp}"/><w:szCs w:val="{hp}"/><w:lang w:val="tr-TR"/></w:rPr></w:rPrDefault>'
                     '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
                     '</w:pPrDefault></w:docDefaults>'
                     '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
                     '<w:qFormat/></w:style>'
                     '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>'
                     '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" '
                     'w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
                     '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>'
                     '</w:style></w:styles>', CT + 'styles+xml')
        self.doc.rel('styles', 'styles.xml')
        self.pkg.add('word/settings.xml',
                     f'{XMLDECL}<w:settings {NSDECL}><w:defaultTabStop w:val="720"/><w:compat>'
                     '<w:compatSetting w:name="compatibilityMode" '
                     'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>',
                     CT + 'settings+xml')
        self.doc.rel('settings', 'settings.xml')
        if any(not v['literal'] for v in self.lists.values()):
            self.pkg.add('word/numbering.xml', self.numbering_xml(), CT + 'numbering+xml')
            self.doc.rel('numbering', 'numbering.xml')
        self.pkg.add('word/_rels/document.xml.rels', self.doc.rels_xml())
        self.pkg.add('_rels/.rels',
                     XMLDECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                     f'<Relationship Id="rId1" Type="{REL}officeDocument" Target="word/document.xml"/></Relationships>')
        self.pkg.write(out_path)
        return spec


def cevir(udf, cikti):
    """UDF'yi DOCX'e çevirir; özet sözlüğü döndürür, sorun olursa DonusumHatasi atar."""
    c = Cevirici(udf)
    try:
        spec = c.build(cikti)
    except DonusumHatasi:
        raise
    except Exception as ex:                          # arayüze traceback sızmasın
        raise DonusumHatasi(f'belge çevrilemedi: {type(ex).__name__}: {ex}')
    return dict(c.stats, dikey_birlesik=c.absorbed, karakter=len(c.cd16) // 2, ustbilgi=spec['header'][0] is not None,
                altbilgi=spec['footer'][0] is not None, imzali=c.signed, uyarilar=sorted(c.warn))


def main():
    ap = argparse.ArgumentParser(description='UYAP UDF belgesini Word (DOCX) biçimine çevirir.')
    ap.add_argument('udf')
    ap.add_argument('-o', '--output', required=True)
    a = ap.parse_args()
    try:
        o = cevir(a.udf, a.output)
    except DonusumHatasi as ex:
        sys.exit(f'HATA: {ex}')
    extra = ', '.join(x for x in (f'{o["gorsel"]} görsel' if o['gorsel'] else '',
                                  f'{o["sayfa_sonu"]} sayfa sonu' if o['sayfa_sonu'] else '',
                                  'üstbilgi' if o['ustbilgi'] else '',
                                  'altbilgi' if o['altbilgi'] else '') if x)
    print(f'Yazıldı: {a.output} ({o["paragraf"]} paragraf, {o["tablo"]} tablo'
          + (f', {extra}' if extra else '') + ')')
    for m in o['uyarilar']:
        print('UYARI:', m)


if __name__ == '__main__':
    main()
