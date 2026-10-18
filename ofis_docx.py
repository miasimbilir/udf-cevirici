#!/usr/bin/env python3
"""ofis_docx.py — .doc / .rtf / .odt / .pages belgesini, bilgisayardaki ofis programına .docx yaptırır.

Bu biçimleri kendimiz düzgün okuyamayız (macOS'un yerleşik dönüştürücüsü tabloları, üst/altbilgiyi ve
sayfa boyutunu kaybediyor; 03.10.2026'da denendi). Bu yüzden belge, kullanıcının bilgisayarında kurulu
programa açtırılıp .docx olarak kaydettirilir; ardından docx_udf.py ile UDF'ye çevrilir.

Sıra:
  Windows : Microsoft Word (COM, görünmez) -> LibreOffice
  macOS   : .pages için Pages; diğerleri için Microsoft Word (yoksa Pages, o da yoksa LibreOffice)
Kullanıcının dosyasına dokunulmaz: kopyası geçici bir klasörde açılır, .docx de orada üretilir.

Kullanım:
    python3 araclar/ofis_docx.py belge.doc -o belge.docx
"""
import argparse
import os
import shutil
import subprocess
import sys
import tempfile
import uuid

UZANTILAR = ('.doc', '.dot', '.docm', '.dotx', '.dotm', '.rtf', '.odt', '.pages')
ZAMAN = 180                                       # saniye; büyük belgede Word yavaş açılabilir
_NO_WINDOW = 0x08000000 if os.name == 'nt' else 0

WORD_MAC = 'com.microsoft.Word'
PAGES_MAC = 'com.apple.Pages'


class OfisHatasi(Exception):
    """Kullanıcıya gösterilecek, anlaşılır hata."""


def destekli(yol):
    return os.path.splitext(yol)[1].lower() in UZANTILAR


def _mac_uygulama_var(bundle):
    r = subprocess.run(['mdfind', f'kMDItemCFBundleIdentifier == "{bundle}"'],
                       capture_output=True, text=True, timeout=20)
    return any(p.endswith('.app') for p in r.stdout.splitlines())


def _libreoffice():
    for p in (shutil.which('soffice'), shutil.which('libreoffice'),
              '/Applications/LibreOffice.app/Contents/MacOS/soffice',
              r'C:\Program Files\LibreOffice\program\soffice.exe',
              r'C:\Program Files (x86)\LibreOffice\program\soffice.exe'):
        if p and os.path.exists(p):
            return p
    return None


def _calistir(cmd, env=None):
    r = subprocess.run(cmd, capture_output=True, text=True, errors='replace', timeout=ZAMAN, env=env,
                       creationflags=_NO_WINDOW)
    return r.returncode, (r.stderr or r.stdout or '').strip()


# ---------------------------------------------------------------- macOS
# Word ve Pages korumalı alanda (sandbox) çalışır. AppleScript'in "open" komutuyla verilen dosya Word'de
# "Dosyaya Erişim İzni Ver" penceresi açıyor; Finder'daki çift tıklamanın yaptığı gibi `open -b` ile
# açılan dosya ise izin sormadan açılıyor (03.10.2026'da sınandı). Word sonucu yalnız kendi klasörüne
# (~/Library/Containers/com.microsoft.Word/Data/tmp) izinsiz yazabiliyor; Pages ise /tmp'ye yazabiliyor
# (Pages'in kendi klasörüne macOS bize erişim vermiyor).
_AS_BEKLE = '''
        repeat 240 times
            if (exists document ad) then exit repeat
            delay 0.5
        end repeat
        set d to document ad'''

_AS_WORD = '''on run argv
    set ad to item 1 of argv
    set hedef to item 2 of argv
    set kapat to item 3 of argv
    set yeni to item 4 of argv
    tell application id "com.microsoft.Word"''' + _AS_BEKLE + '''
        save as d file name hedef file format format document
        -- Yalnız bizim (benzersiz adlı) belgemiz kapatılır; "active document" kullanıcının belgesi olabilir.
        -- Farklı kaydedince belgenin adı değişir: önce yeni adla, olmazsa eski başvuruyla denenir.
        -- "count of documents" Word'de hata verir (-1708); "count documents" çalışır (18.10.2026 sınandı).
        try
            close document yeni saving no
        on error
            try
                close d saving no
            end try
        end try
        if kapat is "1" and (count documents) is 0 then quit saving no
    end tell
end run'''

_AS_PAGES = '''on run argv
    set ad to item 1 of argv
    set hedef to item 2 of argv
    set kapat to item 3 of argv
    tell application id "com.apple.Pages"''' + _AS_BEKLE + '''
        export d to (POSIX file hedef) as Microsoft Word
        close d saving no
        if kapat is "1" and (count documents) is 0 then quit
    end tell
end run'''


def _mac_calisiyor(bundle):
    r = subprocess.run(['osascript', '-e', f'application id "{bundle}" is running'],
                       capture_output=True, text=True, timeout=20)
    return r.stdout.strip() == 'true'


def _mac(kaynak, bundle, betik):
    d = tempfile.mkdtemp(prefix='udf_ofis_', dir='/tmp' if os.path.isdir('/tmp') else None)
    ad = 'udf-cevirici-' + uuid.uuid4().hex[:8] + os.path.splitext(kaynak)[1].lower()   # kullanıcının açık belgeleriyle karışmasın
    kopya = os.path.join(d, ad)
    if bundle == WORD_MAC:
        kap = os.path.expanduser('~/Library/Containers/com.microsoft.Word/Data/tmp')
        hedef = os.path.join(kap, os.path.splitext(ad)[0] + '.docx')
    else:
        hedef = os.path.join(d, 'cikti.docx')
    try:
        (shutil.copytree if os.path.isdir(kaynak) else shutil.copy2)(kaynak, kopya)   # eski .pages klasör olabilir
        kapat = '0' if _mac_calisiyor(bundle) else '1'       # kullanıcı zaten açmışsa programı kapatma
        kod, msg = _calistir(['open', '-g', '-b', bundle, kopya])
        if kod != 0:
            return None, msg or 'program açılamadı'
        kod, msg = _calistir(['osascript', '-e', betik, ad, hedef, kapat, os.path.basename(hedef)])
        if os.path.exists(hedef) and os.path.getsize(hedef) > 0:
            with open(hedef, 'rb') as f:
                return f.read(), None
        if '-1743' in msg or 'not allowed' in msg.lower():
            return None, ('macOS izin vermedi. Sistem Ayarları → Gizlilik ve Güvenlik → Otomasyon bölümünde '
                          'UDF Çevirici için ' + ('Word' if bundle == WORD_MAC else 'Pages') + ' kutusunu açın')
        return None, msg or f'çıkış kodu {kod}'
    finally:
        shutil.rmtree(d, ignore_errors=True)
        if bundle == WORD_MAC and os.path.exists(hedef):
            os.remove(hedef)


# ---------------------------------------------------------------- Windows
# Quit([ref]0): PowerShell 5.1'de Quit(0) "Use [ref]" hatası veriyor, görünmez Word açık kalıyordu ve bu
# hata asıl hatayı ("Parola yanlış") örtüyordu (18.10.2026 Windows denemesi). Kapatma hatası artık yutulur.
# Betiğin içine çift tırnak konmaz: -Command argümanı olarak geçerken kaçış sorunu çıkarabilir.
_PS_WORD = r"""
$ErrorActionPreference = 'Stop'
$w = $null
try {
    $w = New-Object -ComObject Word.Application
    $w.Visible = $false
    $w.DisplayAlerts = 0
    # Parolalı belgede Word görünmez bir parola penceresinde asılı kalmasın: yanlış parola verilir, hemen hata döner
    $d = $w.Documents.Open($env:UDF_KAYNAK, $false, $true, $false, '__udf_parola_yok__')
    $d.SaveAs2($env:UDF_HEDEF, 16)
    $d.Close(0)
} finally {
    if ($w) {
        try { $w.Quit([ref]0) } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) } catch { }
    }
}
"""


def _windows_word(kaynak, d):
    hedef = os.path.join(d, 'belge.docx')
    env = dict(os.environ, UDF_KAYNAK=kaynak, UDF_HEDEF=hedef)    # Türkçe adlar tırnak/kodlama sorunu çıkarmasın
    kod, msg = _calistir(['powershell', '-NoProfile', '-NonInteractive', '-Command', _PS_WORD], env)
    if os.path.exists(hedef) and os.path.getsize(hedef) > 0:
        with open(hedef, 'rb') as f:
            return f.read(), None
    if 'ComObject' in msg or '80040154' in msg or 'Word.Application' in msg:
        return None, 'Microsoft Word kurulu değil'
    if 'parola' in msg.lower() or 'password' in msg.lower():
        return None, 'belge parolayla korunuyor; Word\'de açıp parolasını kaldırın'
    return None, (msg.splitlines() or [''])[0].strip() or f'çıkış kodu {kod}'      # PowerShell'in ilk satırı yeter


def _libre(kaynak, d):
    lo = _libreoffice()
    if not lo:
        return None, 'LibreOffice kurulu değil'
    kod, msg = _calistir([lo, '--headless', '--convert-to', 'docx', '--outdir', d, kaynak])
    hedef = os.path.join(d, os.path.splitext(os.path.basename(kaynak))[0] + '.docx')
    if os.path.exists(hedef):
        with open(hedef, 'rb') as f:
            return f.read(), None
    return None, msg or f'çıkış kodu {kod}'


def docx_yap(kaynak, hedef):
    """kaynak (.doc/.rtf/.odt/.pages) -> hedef .docx. Döner: kullanılan programın adı.
    Başarısızsa OfisHatasi (kullanıcıya gösterilecek açıklamayla)."""
    if not os.path.exists(kaynak):
        raise OfisHatasi(f'dosya bulunamadı: {kaynak}')
    uz = os.path.splitext(kaynak)[1].lower()
    if uz not in UZANTILAR:
        raise OfisHatasi(f'{uz} biçimi desteklenmiyor')
    denenen = []
    adaylar = []
    if sys.platform == 'darwin':
        if uz == '.pages':
            if not _mac_uygulama_var(PAGES_MAC):
                raise OfisHatasi('Pages belgesini çevirmek için bu Mac\'te Pages kurulu olmalı.')
            adaylar = [('Pages', lambda: _mac(kaynak, PAGES_MAC, _AS_PAGES))]
        elif _mac_uygulama_var(WORD_MAC):            # Word varsa yalnız Word: sonuç onun dönüşümüdür
            adaylar = [('Microsoft Word', lambda: _mac(kaynak, WORD_MAC, _AS_WORD))]
        elif _mac_uygulama_var(PAGES_MAC):
            adaylar = [('Pages', lambda: _mac(kaynak, PAGES_MAC, _AS_PAGES))]
    elif uz == '.pages':
        raise OfisHatasi('Pages belgesi yalnız Mac\'te çevrilebilir. Pages\'te Dosya → Dışa Aktar → Word '
                         'ile .docx olarak kaydedip onu çevirin.')
    d = tempfile.mkdtemp(prefix='udf_ofis_')
    try:
        if os.name == 'nt':
            adaylar = [('Microsoft Word', lambda: _windows_word(os.path.abspath(kaynak), d))]
            if _libreoffice():                        # Windows'ta Word'ün kurulu olup olmadığı ancak denenince anlaşılır
                adaylar.append(('LibreOffice', lambda: _libre(os.path.abspath(kaynak), d)))
        if not adaylar and uz != '.pages':             # Word/Pages yoksa son çare
            adaylar.append(('LibreOffice', lambda: _libre(os.path.abspath(kaynak), d)))
        for ad, fn in adaylar:
            try:
                veri, neden = fn()
            except subprocess.TimeoutExpired:
                veri, neden = None, 'yanıt vermedi (zaman aşımı)'
            except Exception as ex:                   # beklenmeyen: sıradakine geç
                veri, neden = None, f'{type(ex).__name__}: {ex}'
            if veri and veri[:2] == b'PK':
                with open(hedef, 'wb') as f:
                    f.write(veri)
                return ad
            denenen.append(f'{ad}: {neden}')
    finally:
        shutil.rmtree(d, ignore_errors=True)
    ipucu = ('Pages\'te Dosya → Dışa Aktar → Word' if uz == '.pages'
             else 'Word\'de "Farklı Kaydet → Word Belgesi (.docx)"')
    raise OfisHatasi(f'{uz} belgesi Word biçimine çevrilemedi ({"; ".join(denenen) or "uygun program yok"}). '
                     f'{ipucu} ile kaydedip .docx dosyasını çevirin.')


def main():
    ap = argparse.ArgumentParser(description='.doc/.rtf/.odt/.pages belgesini kurulu ofis programıyla .docx yapar.')
    ap.add_argument('kaynak')
    ap.add_argument('-o', '--output', required=True)
    a = ap.parse_args()
    try:
        ad = docx_yap(a.kaynak, a.output)
    except OfisHatasi as ex:
        sys.exit(f'HATA: {ex}')
    print(f'Yazıldı: {a.output} ({ad} ile)')


if __name__ == '__main__':
    main()
