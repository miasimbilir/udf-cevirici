#!/usr/bin/env python3
"""udf_docx_testi.py — UDF -> DOCX çeviricisinin (udf_docx.py) senaryo testleri.

Üç grup:
  1. GİDİŞ DÖNÜŞ: udf_senaryo_testi.py'nin her DOCX senaryosu UDF'ye çevrilir, o UDF Word'e geri
     çevrilir, çıkan Word yeniden UDF yapılır. İki UDF karşılaştırılır: metin, biçim (karakter
     karakter), paragraf ölçüleri, liste türleri, tablo/görsel/sayfa sonu sayıları. Böylece Word'e
     yazılan her bilginin gerçekten okunabilir olduğu, kendi okuyucumuzla kanıtlanır.
  2. EDİTÖR ÖRNEKLERİ: Editör'ün yazdığı ama docx_udf.py'nin üretmediği yapılar elle kurulur
     (SecListTypeLevelN, NumberSetted, space/tab/field yaprakları, stil zinciri, bgImage, sayfa
     numarası bitleri, startPage, borderTable/borderRow, headerRow, imzalı arşiv...).
  3. BOZUK GİRDİ: temiz hata, traceback yok.
Her çıktı udf_dogrula.dogrula_docx() denetiminden de geçer.

Kullanım:
    python3 araclar/udf_docx_testi.py [-k süzgeç] [--tut KLASOR]
"""
import argparse
import base64
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from xml.etree import ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import docx_udf                                   # noqa: E402
import udf_docx                                   # noqa: E402
import udf_dogrula                                # noqa: E402
import udf_senaryo_testi as ileri                 # noqa: E402

W = udf_docx.W
NBSP = '\u00a0'
TOL = 0.6                                         # punto; twip yuvarlaması için pay


def wq(t):
    return f'{{{W}}}{t}'


# ---------------------------------------------------------------- UDF modeli (karşılaştırma için)
PARA_NUM = ('LeftIndent', 'RightIndent', 'Hanging', 'FirstLineIndent', 'SpaceAbove', 'SpaceBelow', 'LineSpacing')
CHAR = ('bold', 'italic', 'underline', 'strikethrough', 'superscript', 'subscript', 'size', 'family',
        'foreground', 'background')


def model(path):
    z = zipfile.ZipFile(path)
    root = ET.fromstring(z.read('content.xml'))
    cd16 = (root.find('content').text or '').encode('utf-16-le')
    els = root.find('elements')
    paras = []
    for p in els.iter('paragraph'):
        text, sig = '', []
        for c in p:
            if c.get('startOffset') is None:
                continue
            so, ln = int(c.get('startOffset')), int(c.get('length'))
            t = cd16[2 * so:2 * (so + ln)].decode('utf-16-le', 'replace').replace('\n', '')
            if c.tag == 'image':
                t = '\x01'
            a = tuple((k, c.get(k)) for k in CHAR if c.get(k) not in (None, 'false'))
            text += t
            sig += [a] * len(t)
        # zemin bandı (NBSP) karşılaştırma dışı: genişlik tahminine bağlı bir taklittir
        keep = [i for i, ch in enumerate(text) if ch != NBSP]
        text = ''.join(text[i] for i in keep)
        sig = [sig[i] for i in keep]
        paras.append({'text': text, 'sig': sig, 'a': dict(p.attrib)})
    return {'paras': paras, 'tables': sum(1 for _ in els.iter('table')),
            'images': sum(1 for _ in els.iter('image')), 'pb': len(els.findall('page-break')),
            'header': els.findall('header'), 'footer': els.findall('footer'),
            'page': dict(root.find('properties/pageFormat').attrib), 'els': els,
            'bg': _bg(root.find('properties/bgImage'))}


def _bg(b):
    """(arka plan var mı, sol, üst, sağ, alt pay)"""
    if b is None or not (b.get('bgImageData') or '').strip():
        return None
    return tuple(float(b.get(k) or 0) for k in ('bgImageLeftMargin', 'bgImageUpMargin', 'bgImageRigtMargin',
                                                 'bgImageBottomMargin'))


def compare(m1, m2, loose=()):
    """İki UDF modeli arasındaki farklar (boş liste = aynı). loose: bilinen, belgelenmiş farklar."""
    errs = []
    p1 = [p for p in m1['paras'] if p['text'].strip()]
    p2 = [p for p in m2['paras'] if p['text'].strip()]
    if 'sabit_etiket' in loose and len(p1) == len(p2):
        # harfli liste Word'e bilerek sabit metin etiketle yazılır (Türk alfabesi); ikinci turda liste değildir
        for a, b in zip(p1, p2):
            m = re.fullmatch(r'(\S{1,8}\t)' + re.escape(a['text']), b['text'])
            if a['a'].get('ListId') and m:
                n = len(m.group(1))
                b['text'], b['sig'], b['lit'] = b['text'][n:], b['sig'][n:], True
    if [p['text'] for p in p1] != [p['text'] for p in p2]:
        a, b = [p['text'] for p in p1], [p['text'] for p in p2]
        i = next((k for k in range(min(len(a), len(b))) if a[k] != b[k]), min(len(a), len(b)))
        errs.append(f'metin farkı (paragraf {i + 1}/{len(a)}↔{len(b)}): '
                    f'{(a[i] if i < len(a) else "—")[:50]!r} ↔ {(b[i] if i < len(b) else "—")[:50]!r}')
        return errs
    for k in ('tables', 'images', 'pb'):
        if m1[k] != m2[k]:
            errs.append(f'{k} sayısı {m1[k]} ↔ {m2[k]}')
    for i, (a, b) in enumerate(zip(p1, p2)):
        tag = f'paragraf {i + 1} ({a["text"][:24]!r})'
        if b.get('lit'):
            continue
        if a['sig'] != b['sig'] and 'bicim' not in loose:
            j = next(k for k in range(len(a['sig'])) if a['sig'][k] != b['sig'][k])
            errs.append(f'{tag}: yazı biçimi farkı, {j}. karakter: {dict(a["sig"][j])} ↔ {dict(b["sig"][j])}')
        if (a['a'].get('Alignment') or '0') != (b['a'].get('Alignment') or '0'):
            errs.append(f'{tag}: hizalama {a["a"].get("Alignment")} ↔ {b["a"].get("Alignment")}')
        for k in PARA_NUM:
            if k in loose:
                continue
            x, y = float(a['a'].get(k) or 0), float(b['a'].get(k) or 0)
            if abs(x - y) > (0.006 if k == 'LineSpacing' else TOL):
                errs.append(f'{tag}: {k} {x:g} ↔ {y:g}')
        for k in ('Numbered', 'Bulleted', 'NumberType', 'BulletType', 'ListLevel'):
            if a['a'].get(k) != b['a'].get(k) and 'liste' not in loose:
                errs.append(f'{tag}: {k} {a["a"].get(k)} ↔ {b["a"].get(k)}')
        if 'TabSet' not in loose and _tabs(a['a'].get('TabSet')) != _tabs(b['a'].get('TabSet')):
            errs.append(f'{tag}: TabSet {a["a"].get("TabSet")} ↔ {b["a"].get("TabSet")}')
    for k in ('mediaSizeName', 'paperOrientation'):
        if m1['page'].get(k) != m2['page'].get(k):
            errs.append(f'sayfa {k}: {m1["page"].get(k)} ↔ {m2["page"].get(k)}')
    for k in ('leftMargin', 'rightMargin', 'topMargin', 'bottomMargin'):
        if abs(float(m1['page'].get(k) or 0) - float(m2['page'].get(k) or 0)) > TOL:
            errs.append(f'sayfa {k}: {m1["page"].get(k)} ↔ {m2["page"].get(k)}')
    b1, b2 = m1.get('bg'), m2.get('bg')
    if (b1 is None) != (b2 is None) or (b1 and any(abs(x - y) > 2.0 for x, y in zip(b1, b2))):
        errs.append(f'sayfa arka planı (antet) {b1} ↔ {b2}')
    for kind in ('header', 'footer'):
        r1 = [(e.get('startPage'), e.get('stopPage')) for e in m1[kind]]
        r2 = [(e.get('startPage'), e.get('stopPage')) for e in m2[kind]]
        if r1 != r2:
            errs.append(f'{kind} aralığı {r1} ↔ {r2}')
        s1 = [e.get('pageNumber-spec') for e in m1[kind]]
        s2 = [e.get('pageNumber-spec') for e in m2[kind]]
        if s1 != s2:
            errs.append(f'{kind} sayfa numarası {s1} ↔ {s2}')
    t1 = [(t.get('columnCount'), t.get('border'), len(t.findall('row'))) for t in m1['els'].iter('table')]
    t2 = [(t.get('columnCount'), t.get('border'), len(t.findall('row'))) for t in m2['els'].iter('table')]
    if t1 != t2 and 'tablo' not in loose:
        errs.append(f'tablo yapısı {t1} ↔ {t2}')
    return errs[:6]


def _tabs(ts):
    """Karşılaştırma için: sondaki varsayılan (36'lık, sola hizalı) duraklar atılır."""
    if not ts:
        return []
    out = []
    for item in ts.split(','):
        pos, al, ld = (item.split(':') + ['0', '0'])[:3]
        out.append((round(float(pos)), al, ld))
    return [s for s in out if not (s[1] == '0' and s[2] == '0')]


# Gidiş dönüşte bilinen ve kabul edilen farklar (her biri gerekçesiyle)
LOOSE = {
    # dikey birleştirme taklidi iç içe tablo kurar; blok hizası tahmini boşluklarla sağlanır
    'birlesik_hucre': ('SpaceAbove', 'SpaceBelow'),
    # dikey hizalama taklidi (tahmini üst boşluk) ikinci turda yeniden hesaplanır
    'hucre_dolgu_hizalama': ('SpaceAbove', 'SpaceBelow'),
}
LOOSE.update({
    # 4+ kalemli harfli liste: Editör etiketleri (…c, ç, d…) Word'de sabit metin; liste niteliği bilerek bırakılır
    'liste_harf_alfabe_uyarisi': ('sabit_etiket',),
    'liste_bicim_ve_baslangic': ('sabit_etiket',),      # C) ile başlayan liste NumberSetted ile 4. harfe ulaşır
})
# Gidiş dönüşe girmeyenler: yan görsel taklidi beklemede (varsayılan kapalı, ayrı bayrak ister)
SKIP = {'gorsel_anchor', 'gorsel_yan_metin'}


# ---------------------------------------------------------------- elle UDF kurucu (Editör örnekleri)
class B:
    def __init__(self):
        self.text, self.off = [], 0

    def leaf(self, t, attrs='', tag='content'):
        n = len(t.encode('utf-16-le')) // 2
        x = f'<{tag}{attrs} startOffset="{self.off}" length="{n}" />'
        self.text.append(t)
        self.off += n
        return x

    def para(self, runs, pattrs='', eol_tag='content'):
        xs = ''.join(self.leaf(*r) if isinstance(r, tuple) else r for r in runs) + self.leaf('\n', '', eol_tag)
        return f'<paragraph{pattrs}>{xs}</paragraph>'

    def image(self, data, w_pt, h_pt):
        n = self.off
        self.text.append('¸')
        self.off += 1
        return (f'<image imageData="{base64.b64encode(data).decode()}" width="{w_pt}" height="{h_pt}" '
                f'startOffset="{n}" length="1" />')

    def table(self, rows, spans, border='borderCell'):
        out = (f'<table tableName="Sabit" columnCount="{len(spans)}" '
               f'columnSpans="{",".join(map(str, spans))}" border="{border}">')
        for ra, cells in rows:
            out += f'<row rowName="r"{ra}>' + ''.join(f'<cell>{c}</cell>' for c in cells) + '</row>'
        return out + '</table>'

    def save(self, path, elements, page='mediaSizeName="1" paperOrientation="1"', bg='', styles=None, extra=None):
        cd = ''.join(self.text)
        styles = styles or ('<style name="default" description="Geçerli" family="Dialog" size="12" bold="false" '
                            'italic="false" foreground="-13421773" /><style name="hvl-default" '
                            'family="Times New Roman" size="12" description="Gövde" />'
                            '<style name="defaultdefault" resolver="hvl-default" description="Gövde" />')
        xml = ('<?xml version="1.0" encoding="UTF-8" ?> \n\n<template format_id="1.8" >\n'
               f'<content><![CDATA[{cd}]]></content>\n<properties><pageFormat {page} leftMargin="70.85" '
               'rightMargin="56.7" topMargin="85.0" bottomMargin="60.0" headerFOffset="20.0" '
               f'footerFOffset="30.0" />{bg}</properties>\n'
               f'<elements resolver="hvl-default" >\n{elements}\n</elements>\n<styles>{styles}</styles>\n</template>\n')
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('content.xml', xml.encode('utf-8'))
            for name, data in (extra or {}).items():
                z.writestr(name, data)


class D:
    """Üretilen DOCX üzerinde beklenti yazmayı kolaylaştıran okuyucu."""

    def __init__(self, path, ozet):
        self.z = zipfile.ZipFile(path)
        self.names = set(self.z.namelist())
        self.ozet = ozet
        self.warn = ' | '.join(ozet['uyarilar'])
        self.doc = ET.fromstring(self.z.read('word/document.xml'))
        self.body = self.doc.find(wq('body'))

    def xml(self, part):
        return ET.fromstring(self.z.read(part)) if part in self.names else None

    def paras(self, root=None):
        return list((root if root is not None else self.body).iter(wq('p')))

    @staticmethod
    def text(p):
        return ''.join((t.text or '') if t.tag == wq('t') else '\t' for t in p.iter() if t.tag in (wq('t'), wq('tab')))

    def find(self, needle, root=None):
        return next((p for p in self.paras(root) if needle in self.text(p)), None)

    @staticmethod
    def ppr(p, tag, attr=None):
        el = p.find(f'{wq("pPr")}/{wq(tag)}')
        if el is None:
            return None
        return el.get(wq(attr)) if attr else el

    def levels(self, numid):
        num = self.xml('word/numbering.xml')
        n = next(x for x in num.findall(wq('num')) if x.get(wq('numId')) == numid)
        aid = n.find(wq('abstractNumId')).get(wq('val'))
        a = next(x for x in num.findall(wq('abstractNum')) if x.get(wq('abstractNumId')) == aid)
        lv = {l.get(wq('ilvl')): (l.find(wq('numFmt')).get(wq('val')), l.find(wq('lvlText')).get(wq('val')))
              for l in a.findall(wq('lvl'))}
        ov = {o.get(wq('ilvl')): o.find(wq('startOverride')).get(wq('val')) for o in n.findall(wq('lvlOverride'))}
        return lv, ov


ORNEK = []


def ornek(name, note):
    def deco(fn):
        ORNEK.append((name, note, fn))
        return fn
    return deco


def LIST(lid, lvl, typ, extra='', li=None):
    kind = 'Bulleted="true" BulletType' if typ.startswith('BULLET') else 'Numbered="true" NumberType'
    return (f' Alignment="0" LeftIndent="{li if li is not None else 30.0 * lvl}" {kind}="{typ}" '
            f'ListId="{lid}" ListLevel="{lvl}"{extra}')


@ornek('liste_alt_duzey_varsayilan', 'SecList yok: alt düzeyler Editör varsayılanı (a. / i. / (1))')
def _(path):
    b = B()
    N = 'NUMBER_TYPE_NUMBER_DOT'
    els = [b.para([(f'kalem düzey {l}', '')], LIST(7, l, N)) for l in (1, 2, 3, 4, 1)]
    b.save(path, ''.join(els))

    def exp(d):
        p = d.find('kalem düzey 1')
        lv, _ = d.levels(p.find(f'{wq("pPr")}/{wq("numPr")}/{wq("numId")}').get(wq('val')))
        want = {'0': ('decimal', '%1.'), '1': ('lowerLetter', '%2.'), '2': ('lowerRoman', '%3.'), '3': ('decimal', '(%4)')}
        return [f'düzey {k}: {lv.get(k)} beklenen {v}' for k, v in want.items() if lv.get(k) != v]
    return exp


@ornek('liste_seclist', 'SecListTypeLevelN düzey görünümünü belirler; NumberType ezilir')
def _(path):
    b = B()
    sec = ' SecListTypeLevel1="NUMBER_TYPE_NUMBER_PARANTHESE" SecListTypeLevel2="NUMBER_TYPE_ROMAN_BIG_TRE"'
    els = [b.para([(f'sec {l}', '')], LIST(9, l, 'NUMBER_TYPE_NUMBER_DOT', sec)) for l in (1, 2)]
    b.save(path, ''.join(els))

    def exp(d):
        lv, _ = d.levels(d.find('sec 1').find(f'{wq("pPr")}/{wq("numPr")}/{wq("numId")}').get(wq('val')))
        want = {'0': ('decimal', '%1)'), '1': ('upperRoman', '%2-')}
        return [f'düzey {k}: {lv.get(k)} beklenen {v}' for k, v in want.items() if lv.get(k) != v]
    return exp


@ornek('liste_numbersetted', 'NumberSetted=n listeyi n+1\'den başlatır; ayrı ListId ayrı sayaç')
def _(path):
    b = B()
    N = 'NUMBER_TYPE_NUMBER_DOT'
    els = [b.para([('altıncı', '')], LIST(3, 1, N, ' NumberSetted="5"')), b.para([('yedinci', '')], LIST(3, 1, N)),
           b.para([('ara metin', '')], ' Alignment="0"'), b.para([('sekizinci', '')], LIST(3, 1, N)),
           b.para([('yeni liste', '')], LIST(4, 1, N))]
    b.save(path, ''.join(els))

    def exp(d):
        nid = lambda t: d.find(t).find(f'{wq("pPr")}/{wq("numPr")}/{wq("numId")}').get(wq('val'))
        e = []
        if len({nid('altıncı'), nid('yedinci'), nid('sekizinci')}) != 1:
            e.append('aynı ListId farklı numId aldı (numara sürmez)')
        if nid('yeni liste') == nid('altıncı'):
            e.append('farklı ListId aynı numId aldı')
        _, ov = d.levels(nid('altıncı'))
        if ov.get('0') != '6':
            e.append(f'başlangıç {ov} (beklenen 6)')
        return e
    return exp


@ornek('liste_girinti', 'Liste: metin LeftIndent\'te, numara solunda sağa yaslı (asılı girinti)')
def _(path):
    b = B()
    b.save(path, b.para([('kalem', '')], LIST(5, 1, 'NUMBER_TYPE_CHAR_BIG_D_PARANTHESE', li=42.5))
           + b.para([('işaret', '')], LIST(6, 1, 'BULLET_TYPE_ARROW', li=42.5)))

    def exp(d):
        e = []
        ind = d.ppr(d.find('kalem'), 'ind')
        if ind.get(wq('left')) != '850' or not ind.get(wq('hanging')):
            e.append(f'liste girintisi {ind.attrib}')
        lv, _ = d.levels(d.find('kalem').find(f'{wq("pPr")}/{wq("numPr")}/{wq("numId")}').get(wq('val')))
        if lv['0'] != ('upperLetter', '(%1)'):
            e.append(f'(A) biçimi: {lv["0"]}')
        lv, _ = d.levels(d.find('işaret').find(f'{wq("pPr")}/{wq("numPr")}/{wq("numId")}').get(wq('val')))
        if lv['0'][0] != 'bullet':
            e.append(f'madde işareti: {lv["0"]}')
        return e
    return exp


@ornek('liste_turk_alfabesi', 'Harfli liste 4. kaleme ulaşırsa etiketler Editör\'deki gibi (…c, ç, d…) sabit metin olur')
def _(path):
    b = B()
    T = 'NUMBER_TYPE_CHAR_SMALL_PARANTHESE'
    els = []
    for i in range(5):
        els.append(b.para([(f'bent {i + 1}', '')], LIST(11, 1, T)))
        if i in (0, 3):
            for j in range(2):
                els.append(b.para([(f'alt {i + 1}.{j + 1}', '')],
                                  LIST(11, 2, 'NUMBER_TYPE_NUMBER_DOT', ' SecListTypeLevel2="NUMBER_TYPE_NUMBER_DOT"')))
    els += [b.para([(f'kısa {i + 1}', '')], LIST(12, 1, 'NUMBER_TYPE_CHAR_BIG_DOT')) for i in range(3)]
    b.save(path, ''.join(els))

    def exp(d):
        e = []
        want = {'bent 4': 'ç)\tbent 4', 'bent 5': 'd)\tbent 5', 'alt 4.1': '1.\talt 4.1', 'alt 4.2': '2.\talt 4.2'}
        for k, v in want.items():
            p = d.find(k)
            if p is None or d.text(p) != v:
                e.append(f'{k}: {d.text(p) if p is not None else None!r} (beklenen {v!r})')
            elif d.ppr(p, 'numPr') is not None:
                e.append(f'{k}: sabit metin etiketin yanında otomatik numara da var')
        if d.ppr(d.find('kısa 3'), 'numPr') is None:
            e.append('3 kalemlik harfli liste otomatik kalmalıydı (A, B, C iki alfabede aynı)')
        if 'Türk' not in d.warn and 'ç, d' not in d.warn:
            e.append('alfabe uyarısı yok')
        return e
    return exp


@ornek('yaprak_turleri', 'space / tab / field yaprakları: metin korunur, alan için uyarı')
def _(path):
    b = B()
    els = b.para([('Ad', ''), ('\t', '', 'tab'), (': ', ''), ('Ayşe YILMAZ', ' fieldName="ad" fieldType="1"', 'field')],
                 ' Alignment="0"')
    els += b.para([], ' Alignment="0"', eol_tag='space')
    els += b.para([('son satır', '')], ' Alignment="0"')
    b.save(path, els)

    def exp(d):
        e = []
        if d.find('Ayşe YILMAZ') is None or '\t' not in d.text(d.find('Ayşe YILMAZ')):
            e.append('alan metni ya da sekme kayıp')
        if len(d.paras()) < 3:
            e.append('space yaprağı boş paragraf vermedi')
        if 'form alanları' not in d.warn:
            e.append('alan uyarısı yok')
        return e
    return exp


@ornek('stil_zinciri', 'Biçim çözümlemesi: yaprak resolver\'ı > paragraf öznitelikleri > stil zinciri')
def _(path):
    b = B()
    styles = ('<style name="hvl-default" family="Times New Roman" size="12" />'
              '<style name="baslik" resolver="hvl-default" family="Arial" size="16" bold="true" />')
    els = b.para([('stilden', ' resolver="baslik"'), (' paragraftan', '')], ' Alignment="1" italic="true" size="9"')
    b.save(path, els, styles=styles)

    def exp(d):
        p = d.find('stilden')
        runs = p.findall(wq('r'))
        r1, r2 = runs[0].find(wq('rPr')), runs[1].find(wq('rPr'))
        e = []
        if r1 is None or r1.find(wq('b')) is None or r1.find(wq('sz')).get(wq('val')) != '32' \
                or r1.find(wq('rFonts')).get(wq('ascii')) != 'Arial' or r1.find(wq('i')) is not None:
            e.append('resolver stilinden gelen biçim yanlış')
        if r2 is None or r2.find(wq('i')) is None or r2.find(wq('sz')).get(wq('val')) != '18':
            e.append('paragraftan miras kalan biçim yanlış')
        return e
    return exp


@ornek('girinti_ve_bosluk', 'Hanging/FirstLineIndent/LineSpacing/SpaceAbove-Below -> w:ind, w:spacing')
def _(path):
    b = B()
    els = b.para([('asılı', '')], ' Alignment="3" LeftIndent="10.0" Hanging="30.0" RightIndent="5.0" '
                                 'SpaceAbove="6.0" SpaceBelow="12.0" LineSpacing="0.5"')
    els += b.para([('ilk satır', '')], ' Alignment="2" LeftIndent="20.0" FirstLineIndent="35.4"')
    b.save(path, els)

    def exp(d):
        e = []
        ind, sp = d.ppr(d.find('asılı'), 'ind'), d.ppr(d.find('asılı'), 'spacing')
        if (ind.get(wq('left')), ind.get(wq('hanging')), ind.get(wq('right'))) != ('800', '600', '100'):
            e.append(f'asılı girinti {ind.attrib}')
        if (sp.get(wq('before')), sp.get(wq('after')), sp.get(wq('line'))) != ('120', '240', '360'):
            e.append(f'boşluk {sp.attrib}')
        if d.ppr(d.find('asılı'), 'jc', 'val') != 'both':
            e.append('iki yana yaslama yok')
        ind = d.ppr(d.find('ilk satır'), 'ind')
        if (ind.get(wq('left')), ind.get(wq('firstLine'))) != ('400', '708'):
            e.append(f'ilk satır girintisi {ind.attrib}')
        return e
    return exp


@ornek('sekme_girintiden', 'TabSet konumları LeftIndent\'ten ölçülür: Word konumu = konum + girinti')
def _(path):
    b = B()
    els = b.para([('Ek-1\t4 sayfa', '')], ' Alignment="0" LeftIndent="50.0" TabSet="350.0:1:1,382.0:0:0,418.0:0:0"')
    b.save(path, els)

    def exp(d):
        tabs = d.ppr(d.find('Ek-1'), 'tabs')
        got = [(t.get(wq('val')), t.get(wq('pos')), t.get(wq('leader'))) for t in tabs] if tabs is not None else []
        return [] if got == [('right', '8000', 'dot')] else [f'sekme durakları {got}']
    return exp


@ornek('tablo_yatay_birlesik', 'Satır columnSpans -> gridSpan; ortak ızgara; başlık satırı; kenarlık türleri')
def _(path):
    b = B()
    c = lambda t: b.para([(t, '')], ' Alignment="0" LeftIndent="3.0" RightIndent="1.0"')
    t1 = b.table([(' rowType="headerRow" columnSpans="300"', [c('BAŞLIK')]),
                  (' rowType="dataRow"', [c('a'), c('b'), c('c')]),
                  (' rowType="dataRow" columnSpans="100,200"', [c('d'), c('ef')])], [100, 100, 100])
    t2 = b.table([(' rowType="dataRow"', [c('çerçeve'), c('yalnız dış')])], [150, 50], 'borderTable')
    t3 = b.table([(' rowType="dataRow"', [c('çizgisiz'), c('tablo')])], [100, 100], 'borderNone')
    b.save(path, t1 + t2 + t3)

    def exp(d):
        e = []
        tbls = d.body.findall(wq('tbl'))
        if len(tbls) != 3:
            return [f'{len(tbls)} tablo (art arda tablolar Word\'de birleşmemeli)']
        rows = tbls[0].findall(wq('tr'))
        span = lambda tc: (tc.find(f'{wq("tcPr")}/{wq("gridSpan")}').get(wq('val'))
                           if tc.find(f'{wq("tcPr")}/{wq("gridSpan")}') is not None else '1')
        got = [[span(tc) for tc in r.findall(wq('tc'))] for r in rows]
        if got != [['3'], ['1', '1', '1'], ['1', '2']]:
            e.append(f'gridSpan {got}')
        if rows[0].find(f'{wq("trPr")}/{wq("tblHeader")}') is None:
            e.append('başlık satırı işareti yok')
        mar = tbls[0].find(f'{wq("tblPr")}/{wq("tblCellMar")}/{wq("left")}').get(wq('w'))
        if mar != '60':
            e.append(f'hücre payı {mar} (Editör teamülü 3 punto -> 60)')
        ind = d.ppr(d.find('BAŞLIK'), 'ind')
        if ind is not None:
            e.append('hücre payı paragraftan düşülmedi')
        bt = tbls[1].find(f'{wq("tblPr")}/{wq("tblBorders")}')
        if bt.find(wq('insideV')).get(wq('val')) != 'nil' or bt.find(wq('top')).get(wq('val')) != 'single':
            e.append('borderTable: iç çizgi kapalı, dış açık olmalı')
        if tbls[2].find(f'{wq("tblPr")}/{wq("tblBorders")}/{wq("top")}').get(wq('val')) != 'nil':
            e.append('borderNone çizgili çıktı')
        return e
    return exp


@ornek('tablo_ic_ice', 'Hücre içinde tablo: Word iç içe tablo; hücre paragrafla biter')
def _(path):
    b = B()
    c = lambda t: b.para([(t, '')], ' Alignment="0"')
    inner = b.table([(' rowType="dataRow"', [c('iç 1'), c('iç 2')])], [100, 100])
    b.save(path, b.table([(' rowType="dataRow"', [c('dış'), inner])], [100, 100]))

    def exp(d):
        outer = d.body.find(wq('tbl'))
        tcs = outer.find(wq('tr')).findall(wq('tc'))
        e = []
        if tcs[1].find(wq('tbl')) is None:
            e.append('iç tablo yok')
        if tcs[1][-1].tag != wq('p'):
            e.append('hücre paragrafla bitmiyor')
        return e
    return exp


@ornek('gorsel_bayt_bayt', 'Görsel yeniden kodlanmaz: UDF\'deki baytlar DOCX\'te aynen durur')
def _(path):
    b = B()
    data = ileri.png(60, 30)
    b.save(path, b.para([b.image(data, 120.0, 60.0)], ' Alignment="1"') + b.para([('alt yazı', '')], ' Alignment="1"'))

    def exp(d):
        media = [n for n in d.names if n.startswith('word/media/')]
        e = []
        if len(media) != 1 or d.z.read(media[0]) != data:
            e.append('görsel baytları değişmiş ya da sayı yanlış')
        ext = d.body.find(f'.//{{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}}extent')
        if ext is None or ext.get('cx') != str(120 * 12700):
            e.append('görsel ölçüsü taşınmadı')
        return e
    return exp


@ornek('antet_bgimage', 'bgImage -> her sayfada üstbilgiye bağlı, metnin arkasında, kenar payları düşülmüş görsel')
def _(path):
    b = B()
    data = base64.b64encode(ileri.png(40, 56)).decode()
    bg = (f'<bgImage bgImageSource="" bgImageData="{data}" bgImageBottomMargin="100" bgImageUpMargin="50" '
          'bgImageRigtMargin="30" bgImageLeftMargin="10" />')
    hdr = '<header stopPage="1">' + b.para([('İLK SAYFA ÜSTBİLGİSİ', '')], ' Alignment="1"') + '</header>'
    b.save(path, hdr + b.para([('gövde', '')], ' Alignment="0"'), bg=bg)

    def exp(d):
        e = []
        heads = [d.xml(n) for n in sorted(d.names) if re.fullmatch(r'word/header\d+\.xml', n)]
        if len(heads) != 2:
            return [f'{len(heads)} üstbilgi parçası (ilk + diğer sayfalar bekleniyordu)']
        wp = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
        for h in heads:
            an = h.find(f'.//{wp}anchor')
            if an is None or an.get('behindDoc') != '1':
                e.append('antet görseli metnin arkasında değil ya da eksik')
                continue
            ex = an.find(f'{wp}extent')
            if (ex.get('cx'), ex.get('cy')) != (str(int((595.3 - 40) * 12700)), str(int((841.9 - 150) * 12700))):
                e.append(f'antet ölçüsü {ex.attrib}')
        if sum(1 for h in heads if d.find('İLK SAYFA', h) is not None) != 1:
            e.append('üstbilgi metni yalnız ilk sayfa parçasında olmalı')
        if d.doc.find(f'.//{wq("titlePg")}') is None:
            e.append('titlePg yok')
        return e
    return exp


@ornek('sayfa_numarasi_bitleri', 'pageNumber-spec bitleri: 32 ortala, 64 sağ, 2048 sayfa/toplam; ön/son ek')
def _(path):
    b = B()
    body = b.para([('gövde', '')], ' Alignment="0"')
    foot = ('<footer pageNumber-spec="BSP32_2088" pageNumber-seperator=" / " pageNumber-foreStr="Sayfa " '
            'pageNumber-afterStr="" pageNumber-fontFace="Arial" pageNumber-fontSize="9" pageNumber-fontBold="true" '
            'pageNumber-pageStartNumStr="3">' + b.para([('Büro altbilgisi', '')], ' Alignment="1"')
            + b.para([('  ', '')], ' Alignment="1"') + '</footer>')
    b.save(path, body + foot)

    def exp(d):
        f = d.xml('word/footer1.xml')
        e = []
        ps = d.paras(f)
        instr = [i.text.strip() for i in f.iter(wq('instrText'))]
        if instr != ['PAGE', 'NUMPAGES']:
            e.append(f'alanlar {instr}')
        if len(ps) != 2 or d.ppr(ps[-1], 'jc', 'val') != 'center':
            e.append(f'{len(ps)} altbilgi paragrafı ya da hiza yanlış (yer tutucu atılmalı, numara ortada)')
        if 'Sayfa ' not in d.text(ps[-1]) or ' / ' not in d.text(ps[-1]):
            e.append('ön ek / ayırıcı yok')
        pn = d.doc.find(f'.//{wq("pgNumType")}')
        if pn is None or pn.get(wq('start')) != '3':
            e.append('başlangıç sayfa numarası taşınmadı')
        return e
    return exp


@ornek('altbilgi_ikinci_sayfadan', 'startPage="2": ilk sayfa boş, diğer sayfalarda üstbilgi; altbilgi her sayfada')
def _(path):
    b = B()
    hdr = '<header startPage="2">' + b.para([('DEVAM ÜSTBİLGİSİ', '')], ' Alignment="2"') + '</header>'
    body = b.para([('gövde', '')], ' Alignment="0"')
    foot = '<footer pageNumber-spec="BSP32_0">' + b.para([('HER SAYFA ALTBİLGİ', '')], ' Alignment="1"') + '</footer>'
    b.save(path, hdr + body + foot)

    def exp(d):
        e = []
        sect = d.doc.find(f'.//{wq("sectPr")}')
        refs = {(r.tag.split('}')[1], r.get(wq('type'))) for r in sect if 'Reference' in r.tag}
        if refs != {('headerReference', 'default'), ('footerReference', 'default'), ('footerReference', 'first')}:
            e.append(f'başvurular {sorted(refs)}')
        if sect.find(wq('titlePg')) is None:
            e.append('titlePg yok')
        return e
    return exp


@ornek('sayfa_a5_yatay', 'mediaSizeName=2 + paperOrientation=0 -> A5 yatay; kenar boşlukları')
def _(path):
    b = B()
    b.save(path, b.para([('yatay', '')], ' Alignment="0"'), page='mediaSizeName="2" paperOrientation="0"')

    def exp(d):
        sz, mar = d.doc.find(f'.//{wq("pgSz")}'), d.doc.find(f'.//{wq("pgMar")}')
        e = []
        if (sz.get(wq('w')), sz.get(wq('h')), sz.get(wq('orient'))) != ('11906', '8391', 'landscape'):
            e.append(f'sayfa {sz.attrib}')
        if (mar.get(wq('left')), mar.get(wq('right')), mar.get(wq('top')), mar.get(wq('bottom'))) != \
                ('1417', '1134', '1700', '1200'):
            e.append(f'kenar boşlukları {mar.attrib}')
        return e
    return exp


@ornek('sayfa_sonu', 'page-break -> sonraki paragrafta pageBreakBefore (boş satır üretmez); tablodan önce ayrı')
def _(path):
    b = B()
    c = lambda t: b.para([(t, '')], ' Alignment="0"')
    els = c('bir') + '<page-break>' + b.para([], '') + '</page-break>' + c('iki')
    els += '<page-break>' + b.para([], '') + '</page-break>' + b.table([(' rowType="dataRow"', [c('hücre')])], [100])
    b.save(path, els)

    def exp(d):
        e = []
        if d.ppr(d.find('iki'), 'pageBreakBefore') is None:
            e.append('pageBreakBefore yok')
        if not any(br.get(wq('type')) == 'page' for br in d.body.iter(wq('br'))):
            e.append('tablodan önceki sayfa sonu kayıp')
        if d.ozet['sayfa_sonu'] != 2:
            e.append(f'sayfa sonu sayısı {d.ozet["sayfa_sonu"]}')
        return e
    return exp


@ornek('imzali_ve_bilinmeyen', 'İmza dosyası olan arşiv + tanınmayan öğe: metin korunur, ikisi de uyarılır')
def _(path):
    b = B()
    els = b.para([('imzalı metin', ''), ('[kod]', ' x="1"', 'yenitur')], ' Alignment="0"')
    els += '<kutu>' + b.para([('kutu içi', '')], ' Alignment="0"') + '</kutu>'
    b.save(path, els, extra={'sign.sgn': b'\x30\x82'})

    def exp(d):
        e = []
        if d.find('kutu içi') is None or d.find('[kod]') is None:
            e.append('tanınmayan öğelerin metni kayıp')
        if 'e-imzalı' not in d.warn or 'tanınmayan öğe' not in d.warn:
            e.append(f'uyarılar eksik: {d.warn}')
        if not d.ozet['imzali']:
            e.append('imzalı bayrağı yok')
        return e
    return exp


@ornek('ozel_karakterler', 'XML özel karakterleri, UTF-16 çiftleri (emoji) ve kontrol karakteri')
def _(path):
    b = B()
    els = b.para([('A & B <c> "d" 😀 son', ''), ('\x0b', '')], ' Alignment="0"') + b.para([('ikinci', ' bold="true"')], '')
    b.save(path, els)

    def exp(d):
        e = []
        if d.find('A & B <c> "d" 😀 son') is None:
            e.append('özel karakterli metin bozuldu')
        p = d.find('ikinci')
        if p is None or p.find(f'{wq("r")}/{wq("rPr")}/{wq("b")}') is None:
            e.append('emoji sonrası ofset kaydı (UTF-16)')
        return e
    return exp


@ornek('hucre_dolgusu_geri', 'docx_udf zemin bandı -> gerçek hücre dolgusu; NBSP bandı metne girmez')
def _(path):
    b = B()
    bg = ' background="-2500135"'                       # D9D9D9
    band = NBSP * 8
    cell1 = b.para([(band, bg), ('Başlık', bg + ' bold="true"'), (band, bg)], ' Alignment="1" LeftIndent="0.5" RightIndent="0.5"')
    cell2 = b.para([('dolgusuz', '')], ' Alignment="0" LeftIndent="5.4" RightIndent="5.4"')
    b.save(path, b.table([(' rowType="dataRow"', [cell1, cell2])], [100, 100]))

    def exp(d):
        e = []
        tcs = d.body.find(wq('tbl')).find(wq('tr')).findall(wq('tc'))
        shd = tcs[0].find(f'{wq("tcPr")}/{wq("shd")}')
        if shd is None or shd.get(wq('fill')) != 'D9D9D9':
            e.append('hücre dolgusu kurulmadı')
        if NBSP in d.text(d.find('Başlık')) or d.find('Başlık').find(f'.//{wq("rPr")}/{wq("shd")}') is not None:
            e.append('zemin bandı metinde kaldı')
        if tcs[1].find(f'{wq("tcPr")}/{wq("shd")}') is not None:
            e.append('dolgusuz hücre boyandı')
        if d.body.find(f'.//{wq("tblCellMar")}/{wq("left")}').get(wq('w')) != '108':
            e.append('hücre payı bantlı paragraftan etkilendi')
        return e
    return exp


@ornek('hucre_dolgusu_vurgulu', 'Zemin bandlı hücrede başka renkle vurgulu kelime: NBSP bandı Word\'e sızmaz (18.10.2026)')
def _(path):
    b = B()
    bg, sari = ' background="-2500135"', ' background="-256"'     # D9D9D9, FFFF00
    band = NBSP * 8
    cell1 = b.para([(band, bg), ('Kısa ', bg), ('vurgulu', sari), (band, bg)],
                   ' Alignment="1" LeftIndent="0.5" RightIndent="0.5"')
    cell2 = b.para([('dolgusuz', '')], ' Alignment="0" LeftIndent="5.4" RightIndent="5.4"')
    b.save(path, b.table([(' rowType="dataRow"', [cell1, cell2])], [100, 100]))

    def exp(d):
        e = []
        tcs = d.body.find(wq('tbl')).find(wq('tr')).findall(wq('tc'))
        shd = tcs[0].find(f'{wq("tcPr")}/{wq("shd")}')
        if shd is None or shd.get(wq('fill')) != 'D9D9D9':
            e.append('hücre dolgusu kurulmadı')
        if NBSP in d.text(d.find('vurgulu')):
            e.append('zemin bandı (NBSP) metne sızdı')
        return e
    return exp


@ornek('gorsel_okunamayan', 'Okunamayan görsel: uyarıyla atlanır, belge yine çevrilir (18.10.2026)')
def _(path):
    b = B()
    b.save(path, b.para([('önce', '')], ' Alignment="0"') + b.para([b.image(b'bozuk veri', 40.0, 40.0)], ' Alignment="0"')
           + b.para([b.image(ileri.png(20, 10), 20.0, 10.0)], ' Alignment="0"') + b.para([('sonra', '')], ' Alignment="0"'))

    def exp(d):
        e = []
        if 'okunamayan bir görsel atlandı' not in d.warn:
            e.append(f'uyarı yok: {d.warn}')
        if len([n for n in d.names if n.startswith('word/media/')]) != 1:
            e.append('sağlam görsel taşınmadı')
        return e
    return exp


def bad_inputs(work):
    res = []
    cases = {'bozuk.udf': b'bu bir zip degil', 'olmayan.udf': None}
    p = os.path.join(work, 'icerik_yok.udf')
    with zipfile.ZipFile(p, 'w') as z:
        z.writestr('baska.xml', '<a/>')
    p2 = os.path.join(work, 'bozuk_xml.udf')
    with zipfile.ZipFile(p2, 'w') as z:
        z.writestr('content.xml', '<template><content>yarım')
    for name, data in cases.items():
        if data is not None:
            open(os.path.join(work, name), 'wb').write(data)
    for name in list(cases) + ['icerik_yok.udf', 'bozuk_xml.udf']:
        src = os.path.join(work, name)
        r = subprocess.run([sys.executable, os.path.join(HERE, 'udf_docx.py'), src, '-o', src + '.docx'],
                           capture_output=True, text=True, encoding='utf-8', errors='replace',
                           env=dict(os.environ, PYTHONIOENCODING='utf-8'))
        errs = []
        if r.returncode == 0:
            errs.append('çıkış kodu 0')
        if 'Traceback' in r.stderr:
            errs.append('Python traceback sızdı')
        if os.path.exists(src + '.docx'):
            errs.append('hata hâlinde çıktı dosyası bırakıldı')
        res.append((f'gecersiz_girdi:{name}', 'Geçersiz/bozuk girdi → temiz hata', errs))
    return res


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('-k', default='', help='senaryo adına göre süz')
    ap.add_argument('--tut', help='üretilen dosyaları bu klasörde sakla')
    a = ap.parse_args()
    work = a.tut or tempfile.mkdtemp(prefix='udf_docx_test_')
    os.makedirs(work, exist_ok=True)
    rows = []

    for name, note, fn in ileri.SCEN:                       # 1) gidiş dönüş
        if (a.k and a.k not in 'gd:' + name) or name in SKIP:
            continue
        d1, u1 = os.path.join(work, name + '.docx'), os.path.join(work, name + '.udf')
        d2, u2 = os.path.join(work, name + '.ters.docx'), os.path.join(work, name + '.ters.udf')
        try:
            if fn(d1) is None:
                rows.append(('gd:' + name, note, 'ATLANDI', ['gerekli araç yok']))
                continue
            docx_udf.cevir(d1, u1)
            udf_docx.cevir(u1, d2)
            ok, lines = udf_dogrula.dogrula_docx(u1, d2)
            errs = [l for l in lines if l.startswith('HATA')]
            if not errs:
                docx_udf.cevir(d2, u2)
                errs = compare(model(u1), model(u2), LOOSE.get(name, ()))
        except Exception as ex:
            errs = [f'{type(ex).__name__}: {ex}'[:200]]
        rows.append(('gd:' + name, note, 'GEÇTİ' if not errs else 'KALDI', errs))

    for name, note, fn in ORNEK:                            # 2) Editör örnekleri
        if a.k and a.k not in name:
            continue
        u, d = os.path.join(work, 'ed_' + name + '.udf'), os.path.join(work, 'ed_' + name + '.docx')
        try:
            exp = fn(u)
            ozet = udf_docx.cevir(u, d)
            ok, lines = udf_dogrula.dogrula_docx(u, d)
            errs = [l for l in lines if l.startswith('HATA')] + exp(D(d, ozet))
        except Exception as ex:
            errs = [f'{type(ex).__name__}: {ex}'[:200]]
        rows.append((name, note, 'GEÇTİ' if not errs else 'KALDI', errs))

    if not a.k or 'gecersiz' in a.k:                        # 3) bozuk girdi
        for name, note, errs in bad_inputs(work):
            rows.append((name, note, 'GEÇTİ' if not errs else 'KALDI', errs))

    if not rows:
        sys.exit('süzgece uyan senaryo yok')
    wn = max(len(r[0]) for r in rows)
    for name, note, st, errs in rows:
        print(f'{st:8} {name:{wn}}  {note}')
        for e in errs:
            print(f'{"":8} {"":{wn}}    ↳ {e}')
    tot = {s: sum(1 for r in rows if r[2] == s) for s in ('GEÇTİ', 'KALDI', 'ATLANDI')}
    print('\nÖZET:', ', '.join(f'{k} {v}' for k, v in tot.items() if v), f'| dosyalar: {work}')
    if not a.tut:
        shutil.rmtree(work, ignore_errors=True)
    sys.exit(0 if tot['KALDI'] == 0 else 1)


if __name__ == '__main__':
    main()
