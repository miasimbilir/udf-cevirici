#!/usr/bin/env python3
"""udf_onizle.py — UDF'yi UYAP Doküman Editörü'nün kendi çizim motoruyla PNG'ye basar.

Pencere açılmaz; bilgisayarda kurulu Editör'ün jar'ları ve kendi Java'sı kullanılır.
Böylece bir UDF'nin Editör'de NASIL GÖRÜNECEĞİ, Editör açılmadan görülebilir.

Kullanım:
    python3 araclar/udf_onizle.py belge.udf                 # belge-sayfa-1.png, -2.png ...
    python3 araclar/udf_onizle.py belge.udf -o KLASOR --olcek 1.5
Editör standart yerde değilse:  UYAP_EDITOR_HOME=/yol/Contents  (Java/ ve PlugIns/ içeren klasör)
Dış kütüphane gerekmez; sayfalar Java tarafında ayrılır.
"""
import argparse
import glob
import os
import re
import shutil
import subprocess
import sys

# PyInstaller paketinde veri dosyaları sys._MEIPASS altına açılır
HERE = getattr(sys, '_MEIPASS', None) or os.path.dirname(os.path.abspath(__file__))
CLS = os.path.join(HERE, 'udf_onizle')


def editor_home():
    cands = [os.environ.get('UYAP_EDITOR_HOME'),
             '/Applications/Uyap Doküman Editörü.app/Contents',
             os.path.expanduser('~/Applications/Uyap Doküman Editörü.app/Contents')]
    cands += glob.glob(r'C:\Program Files*\UYAP*\*Edit*') + glob.glob(r'C:\UYAP*\*Edit*')
    # tek klasörlü kurulum (C:\Program Files\UYAP Doküman Editörü gibi) — Windows'ta sınanmadı
    cands += glob.glob(r'C:\Program Files*\*UYAP*') + glob.glob(r'C:\*UYAP*') + \
        glob.glob(os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Programs', '*UYAP*'))
    for c in cands:
        if c and glob.glob(os.path.join(c, '**', 'editor_lib.jar'), recursive=True):
            return c
    return None


class OnizlemeHatasi(Exception):
    pass


def onizle(udf, out_dir=None, olcek=1.0, tek=False):
    """UDF'yi Editör motoruyla çizer; (bilgi_satiri, [png_yollari]) döndürür."""
    home = editor_home()
    if not home:
        raise OnizlemeHatasi('UYAP Doküman Editörü bulunamadı (UYAP_EDITOR_HOME ile yol verin).')
    jars = sorted(glob.glob(os.path.join(home, '**', '*.jar'), recursive=True))
    jars = [j for j in jars if os.path.basename(j) != 'updater.jar' and 'plugins' not in j.lower()
            and not any(k == 'jre' or k.endswith('.jre') for k in re.split(r'[\\/]', j.lower()))]
    java = next(iter(glob.glob(os.path.join(home, '**', 'bin', 'java'), recursive=True)
                     + glob.glob(os.path.join(home, '**', 'bin', 'java.exe'), recursive=True)), None) \
        or shutil.which('java')
    if not java:
        raise OnizlemeHatasi('Java bulunamadı.')
    if not os.path.exists(os.path.join(CLS, 'UdfRender.class')):
        if not shutil.which('javac'):
            raise OnizlemeHatasi('UdfRender.class yok ve derlemek için javac bulunamadı.')
        subprocess.run(['javac', '--release', '8', '-nowarn', '-cp', os.pathsep.join(jars),
                        '-d', CLS, os.path.join(CLS, 'UdfRender.java')], check=True,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    out_dir = out_dir or os.path.dirname(os.path.abspath(udf))
    os.makedirs(out_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(udf))[0]
    prefix = os.path.join(out_dir, stem)
    for old in glob.glob(glob.escape(prefix) + '-sayfa-*.png') + glob.glob(glob.escape(prefix) + '-tum.png'):
        os.remove(old)                                   # önceki çizimin artıkları karışmasın
    cmd = [java, '-Xmx2g', '-Dapple.awt.UIElement=true', '-cp', os.pathsep.join([CLS] + jars),
           'UdfRender', os.path.abspath(udf), prefix, str(olcek)] + (['tek'] if tek else [])
    kw = {}
    if os.name == 'nt':                                  # Windows'ta konsol penceresi açılmasın
        kw['creationflags'] = 0x08000000
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, errors='replace', timeout=300, **kw)
    except subprocess.TimeoutExpired:
        raise OnizlemeHatasi('Editör 300 sn içinde yanıt vermedi (bir iletişim kutusu beklemiş olabilir).')
    lines = r.stderr.splitlines()
    info = [l for l in lines if l.startswith('UDFRENDER-OK')]
    pages = [l.split(' ', 1)[1] for l in lines if l.startswith('UDFRENDER-SAYFA ')]
    if r.returncode != 0 or not pages:
        tail = ' | '.join((r.stderr or r.stdout).splitlines()[-4:])
        raise OnizlemeHatasi(f'çizim başarısız: {tail}')
    return ' '.join(info), pages


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('udf')
    ap.add_argument('-o', '--out', help='çıktı klasörü (varsayılan: UDF ile aynı yer)')
    ap.add_argument('--olcek', type=float, default=1.0)
    ap.add_argument('--tek', action='store_true', help='sayfalara bölme, tek PNG bırak')
    a = ap.parse_args()
    try:
        info, pages = onizle(a.udf, a.out, a.olcek, a.tek)
    except OnizlemeHatasi as ex:
        sys.exit(f'HATA: {ex}')
    print(info)
    for p in pages:
        print(p)


if __name__ == '__main__':
    main()
