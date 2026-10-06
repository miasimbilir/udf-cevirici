#!/usr/bin/env python3
"""
UDF Çevirici — masaüstü uygulaması.

Word (.docx) belgesini, biçimini koruyarak UYAP Doküman Editörü'nün UDF biçimine;
UDF belgesini de Word'e çevirir. Yönü dosyanın uzantısı belirler. .doc, .rtf, .odt ve .pages
belgeleri önce bilgisayardaki Word'e (Pages belgesi için Pages'e) Word biçimine çevirtilir. Hiçbir dış programa
bağlı değildir; belgeler bilgisayardan çıkmaz. İnternet yalnızca kullanıcı güncelleme
denetimi düğmesine bastığında kullanılır.

Akış:  belgeleri ekle → çevir → (her çıktı kaynağıyla karşılaştırılıp doğrulanır) → rapor / önizleme
"""
import json
import os
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
import urllib.error
import urllib.request
import webbrowser
import zipfile
from tkinter import filedialog, messagebox, ttk

if getattr(sys, "frozen", False):                   # PyInstaller paketi
    sys.path.insert(0, os.path.dirname(os.path.abspath(sys.executable)))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import ofis_docx
import udf_docx
import udf_onizle
from docx_udf import DonusumHatasi, cevir
from udf_dogrula import dogrula, dogrula_docx

# Sürükle-bırak Tkinter'da yerleşik değil; paket yoksa düğmeyle devam edilir.
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    TEMEL_PENCERE, SURUKLENEBILIR = TkinterDnD.Tk, True
except Exception:
    TEMEL_PENCERE, SURUKLENEBILIR = tk.Tk, False

UYGULAMA_ADI = "UDF Çevirici"
SURUM = "1.2.1"
YAZAR = "Av. Arb. Mevlana İbrahim Asım Bilir"
YAZAR_EK = "av.ibrahimbilir@gmail.com"
TELIF = "© 2026 Av. Arb. Mevlana İbrahim Asım Bilir"
LISANS = "Ücretsiz kullanılabilir ve dağıtılabilir; satılamaz."

# Güncelleme denetimi YALNIZCA kullanıcı düğmeye bastığında yapılır; program
# kendiliğinden internete çıkmaz ve hiçbir belge gönderilmez. Depo herkese
# açıldığında adı buraya yazılır; boşken düğme hiç gösterilmez.
DEPO = "miasimbilir/udf-cevirici"
SURUM_API = f"https://api.github.com/repos/{DEPO}/releases/latest"
SURUM_SAYFA = f"https://github.com/{DEPO}/releases/latest"

SAYFA_NO = {"Word'deki gibi": "otomatik", "Ekle": "var", "Ekleme": "yok"}
KABUL = (".docx", ".udf") + ofis_docx.UZANTILAR      # bırakılabilen belgeler
DOLGU = {"Renk bandı": "bant", "Yalnız yazı": "yazi", "Yok": "yok"}


def surum_sayilari(etiket):
    """'v1.2' → (1, 2). Karşılaştırma için sayıya çevirir."""
    parcalar = []
    for p in str(etiket).lstrip("vVsS").split("."):
        rakam = "".join(c for c in p if c.isdigit())
        parcalar.append(int(rakam) if rakam else 0)
    return tuple(parcalar) or (0,)


def son_surumu_sor():
    """Depodaki en son sürüm etiketini döndürür. Ağ hatasında istisna atar."""
    istek = urllib.request.Request(SURUM_API, headers={
        "User-Agent": f"UDF-Cevirici/{SURUM}", "Accept": "application/vnd.github+json"})
    with urllib.request.urlopen(istek, timeout=8) as yanit:
        return json.load(yanit).get("tag_name") or ""


ACIK_TEMA = {"yesil": "#127a3d", "kirmizi": "#b4232a", "soluk": "#6b7280", "sari": "#8a6100",
             "cizgi": "#d7d9dd", "kagit": "#ffffff", "yazi": "#1c1f24", "secim": "#d3e3f5"}
KOYU_TEMA = {"yesil": "#5fd08a", "kirmizi": "#f08d92", "soluk": "#9aa3ae", "sari": "#e3b341",
             "cizgi": "#3a4048", "kagit": "#242a31", "yazi": "#e8eaed", "secim": "#33415a"}


def tema_sec(pencere):
    """Arka planın parlaklığına göre okunur renk kümesi."""
    try:
        r, g, b = pencere.winfo_rgb(ttk.Style().lookup("TFrame", "background")
                                    or pencere.cget("background"))
        return KOYU_TEMA if (r + g + b) / 3 < 32768 else ACIK_TEMA
    except Exception:
        return ACIK_TEMA


# --------------------------------------------------------------------------
# Çevirme işi (arayüzden bağımsız; sınama da bunu kullanır)
# --------------------------------------------------------------------------
def bos_ad(klasor, ad, uzanti=".docx"):
    """Var olan bir dosyanın üzerine yazmayacak ad: 'A.docx' doluysa "A (UDF'den).docx", o da doluysa 2, 3…"""
    aday = os.path.join(klasor, ad + uzanti)
    n = 1
    while os.path.exists(aday):
        aday = os.path.join(klasor, f"{ad} (UDF'den{'' if n == 1 else ' ' + str(n)}){uzanti}")
        n += 1
    return aday


def belge_cevir(kaynak, hedef_klasor=None, sayfa_no="otomatik", dolgu="bant"):
    """Bir belgeyi çevirir ve DOĞRULAR. Doğrulama geçmezse çıktı bırakılmaz.

    Yönü uzantı belirler: .docx → .udf, .udf → .docx. Word'e çevirirken var olan bir Word
    belgesinin üzerine ASLA yazılmaz (asıl belge çoğu zaman UDF'nin yanında durur).
    Döner: {'kaynak', 'cikti', 'yon', 'tamam', 'ozet', 'uyarilar', 'dogrulama', 'hata'}
    """
    ad, uzanti = os.path.splitext(os.path.basename(kaynak.rstrip("/\\")))
    klasor = hedef_klasor or os.path.dirname(os.path.abspath(kaynak.rstrip("/\\")))
    ters = uzanti.lower() == ".udf"
    ofis = ofis_docx.destekli(kaynak)
    s = {"kaynak": kaynak, "cikti": "", "yon": "udf>word" if ters else "word>udf", "tamam": False,
         "ozet": None, "uyarilar": [], "dogrulama": [], "hata": ""}
    if uzanti.lower() not in (".docx", ".udf") and not ofis:
        s["hata"] = "yalnız Word (.docx, .doc, .rtf, .odt), Pages ve UDF belgeleri çevrilir."
        return s
    cikti = bos_ad(klasor, ad) if ters else os.path.join(klasor, ad + ".udf")
    s["cikti"] = cikti
    gecici = cikti + ".gecici"
    ara = None                                              # .doc/.rtf/.odt/.pages -> geçici .docx
    try:
        if ters:
            s["ozet"] = udf_docx.cevir(kaynak, gecici)
            tamam, satirlar = dogrula_docx(kaynak, gecici)
        else:
            docx, on_uyari = kaynak, []
            if ofis:
                ara = os.path.join(tempfile.mkdtemp(prefix="udf_cevirici_"), "belge.docx")
                program = ofis_docx.docx_yap(kaynak, ara)
                docx = ara
                on_uyari = [f"belge önce {program} ile Word biçimine çevrildi; UDF ve doğrulama bu hâline göre "
                            "yapıldı"]
            s["ozet"] = cevir(docx, gecici, sayfa_no, dolgu)
            s["ozet"]["uyarilar"] = on_uyari + s["ozet"]["uyarilar"]
            tamam, satirlar = dogrula(docx, gecici)
        s["uyarilar"] = s["ozet"]["uyarilar"]
        s["dogrulama"] = satirlar
        if not tamam:
            s["hata"] = "çıktı doğrulamadan geçmedi; güvenli olmadığı için dosya yazılmadı."
            return s
        if ters and os.path.exists(cikti):                  # araya başka bir dosya girdiyse yine ezme
            cikti = s["cikti"] = bos_ad(klasor, ad)
        os.replace(gecici, cikti)
        s["tamam"] = True
    except (DonusumHatasi, udf_docx.DonusumHatasi, ofis_docx.OfisHatasi) as e:
        s["hata"] = str(e)
    except OSError as e:
        s["hata"] = f"dosya yazılamadı: {e}"
    except Exception as e:                                  # beklenmeyen: programı düşürme
        s["hata"] = f"beklenmeyen hata ({type(e).__name__}): {e}"
    finally:
        if os.path.exists(gecici):
            try:
                os.remove(gecici)
            except OSError:
                pass
        if ara:
            import shutil
            shutil.rmtree(os.path.dirname(ara), ignore_errors=True)
    return s


# Çekirdeğin teknik uyarıları → kullanıcının anlayacağı kısa cümleler (sonuç kutusunda gösterilir;
# teknik metin Rapor'da kalır). Eşleşme alt dizgiyle yapılır; eşleşmeyen uyarı olduğu gibi gösterilir.
SADE_UYARI = (
    ("dipnotlar belge sonuna", "Dipnotlar sayfa altında değil, belgenin sonunda yer alıyor (UDF'de sayfa altı dipnot yok)."),
    ("içindekiler tablosundaki", "İçindekiler tablosundaki sayfa numaraları Word'deki hâliyle yazıldı; Editör'de sayfalar kayarsa numaralar tutmayabilir."),
    ("hücre dolgusu UDF", "Tablo hücrelerinin dolgu rengi UDF'de yok; renk, yazının arkasına bant olarak verildi."),
    ("hücre gölgelendirmesi", "Tablo hücrelerinin dolgu rengi taşınmadı."),
    ("dikey birleştir", "Dikey birleştirilmiş bazı tablo hücreleri taşınamadı; alttaki hücre boş kaldı."),
    ("korundu; devam sayfalarının", "İlk sayfanın üstbilgisi/altbilgisi korundu; devam sayfalarındaki farklı olan taşınmadı (Editör tek üstbilgi tutuyor)."),
    ("metin kaydırmalı görsel", "Yanından metin akan görsel iki sütunlu tabloyla taklit edildi; yeri Word'dekinden biraz farklı olabilir."),
    ("serbest konumlu görsel", "Serbest konumlu ya da yanından metin akan görsel satır içine alındı; yeri Word'dekinden farklı olabilir."),
    ("metin kutusu", "Metin kutusu düz paragrafa çevrildi; kutunun yeri ve çerçevesi taşınmadı."),
    ("şekil/grafik", "Şekil, grafik ya da SmartArt taşınmadı."),
    ("desteklenmeyen görsel", "Bir görsel (EMF/WMF/PDF gibi) çizilemediği için taşınamadı."),
    ("çevrilemedi, atlandı", "Bir görsel PNG'ye çevrilemediği için taşınamadı."),
    ("VML görsel", "Eski tip bir görsel taşınamadı."),
    ("emoji", "Editör'ün açamadığı karakterler (emoji gibi) □ ile değiştirildi."),
    ("simgenin Unicode", "Bazı özel simgeler □ olarak yazıldı."),
    ("sayfa boyutu Letter", "Sayfa boyutu Letter (Amerikan); A4 istiyorsanız Word'de sayfa boyutunu A4 yapıp yeniden çevirin."),
    ("tanıdığı boyutlardan değil", "Bu sayfa boyutu Editör'de yok; belge A4 yazıldı, satır sonları Word'dekinden farklı olabilir."),
    ("hücre içindeki tablo", "Tablo içindeki tablo düz paragraflara çevrildi."),
    ("hücresindeki sayfa sonu", "Tablo içindeki sayfa sonu atlandı."),
    ("sütun sayısından fazla hücre", "Bir tablo satırında fazla hücre vardı; fazlası atıldı."),
    ("harfli listede Editör Türk", "Harfli listede Editör Türk alfabesiyle sayar: Word'deki d, e… bentleri Editör'de "
                                   "ç, d… görünür. Metinde bent atıfı varsa kontrol edin."),
    ("çok basamaklı numaralandırma", "\"1.1.\" gibi çok basamaklı numaralar Editör'de yok; yalnız son basamak numaralandı."),
    ("ilk sayfanınki kullanıldı, devam sayfalarındaki", "Antet sayfa zemini olarak yazıldı. İlk sayfanın anteti kullanıldı; "
                                                        "devam sayfalarındaki farklı antet taşınmadı."),
    ("Editör'de her sayfada görünür", "Antet sayfa zemini olarak yazıldı. Word'de yalnız bazı sayfalardaydı; "
                                      "Editör'de her sayfada görünür."),
    ("UDF'de sayfa arka planı olarak yazıldı", "Antet, UDF'de sayfa zemini olarak yazıldı (Editör'de her sayfada görünür)."),
    ("ile Word biçimine çevrildi", None),                   # aşağıda programın adıyla doldurulur
    ("sayfadaki boyutuna göre küçültüldü", "Fotoğraflar sayfadaki boyutlarına göre küçültüldü; baskıda fark görülmez, dosya küçüldü."),
    ("MB üzerindeki dosyayı kabul etmeyebilir", "UDF dosyası 10 MB'ı aşıyor; UYAP kabul etmeyebilir. Fotoğrafların bir kısmını ayrı ek olarak yüklemeyi düşünün."),
    # --- UDF → Word ---
    ("e-imzalı", "Bu UDF e-imzalıydı; Word belgesi imza taşımaz, yalnızca çalışma kopyasıdır."),
    ("harf sıralı liste", "Harfli liste etiketleri (a, b, c, ç…) Editör'deki hâliyle düz metin yazıldı; "
                          "Word'de otomatik numara değildir."),
    ("antet (arka plan görseli) Word", "Antet, Word'de üstbilgiye bağlı arka plan görseli olarak yerleştirildi."),
    ("antet (arka plan görseli) okunamadı", "Antet görseli okunamadığı için taşınamadı."),
    ("form alanları", "Form alanları düz metne çevrildi."),
    ("barkod", "Barkod taşınmadı."),
    ("tanınmayan öğe", "Editör'e özgü bazı öğelerin metni korundu, biçimi taşınmadı."),
    ("sayfa aralığı", "Üstbilgi/altbilgi yalnız belirli sayfalar için ayarlanmıştı; Word'de bütün sayfalara konuldu."),
    ("birden çok", "Belgede birden çok üstbilgi/altbilgi vardı; Editör'ün yaptığı gibi sonuncusu alındı."),
    ("okunamayan bir görsel", "Bir görsel okunamadığı için atlandı."),
    ("tanınmayan kâğıt kodu", "Kâğıt boyutu tanınmadı; A4 kullanıldı."),
    ("tanınmayan tablo kenarlık", "Bir tablonun kenarlık türü tanınmadı; tam kenarlık kullanıldı."),
)


def sade_uyarilar(sonuclar, en_cok=4):
    """Bütün belgelerin uyarılarını sadeleştirir, yineleneni atar. Döner: (gösterilecekler, kalan sayısı)."""
    gorulen, liste = set(), []
    for r in sonuclar:
        for u in r["uyarilar"]:
            sade = next((m for a, m in SADE_UYARI if a in u), u[:1].upper() + u[1:])
            if sade is None:                                # "belge önce Microsoft Word ile Word biçimine…"
                program = u.split("belge önce ", 1)[-1].split(" ile ", 1)[0]
                sade = f"Belge önce {program} ile Word biçimine çevrildi, sonra UDF yapıldı."
            if sade not in gorulen:
                gorulen.add(sade)
                liste.append(sade)
    return liste[:en_cok], max(0, len(liste) - en_cok)


def rapor_metni(sonuclar):
    s = ["UDF ÇEVİRİCİ — ÇEVİRİ RAPORU", ""]
    for i, r in enumerate(sonuclar, 1):
        s.append(f"{i}. {os.path.basename(r['kaynak'])}   ({'UDF → Word' if r['yon'] == 'udf>word' else 'Word → UDF'})")
        if r["tamam"]:
            o = r["ozet"]
            s.append(f"   ✓ Çevrildi → {r['cikti']}")
            ek = [f"{o['paragraf']} paragraf", f"{o['tablo']} tablo"]
            if o.get("gorsel"):
                ek.append(f"{o['gorsel']} görsel")
            if o.get("dikey_birlesik"):
                ek.append("dikey birleşik hücreler Word'ün kendi birleştirmesine çevrildi")
            if o["sayfa_sonu"]:
                ek.append(f"{o['sayfa_sonu']} sayfa sonu")
            if o["ustbilgi"]:
                ek.append("üstbilgi")
            if o["altbilgi"]:
                ek.append("altbilgi")
            s.append("   " + ", ".join(ek))
        else:
            s.append(f"   ✕ Çevrilemedi: {r['hata']}")
        for u in r["uyarilar"]:
            s.append(f"   ! {u}")
        if r["dogrulama"]:
            s.append("   Doğrulama:")
            s += [f"     {d}" for d in r["dogrulama"]]
        s.append("")
    if any(r["yon"] == "word>udf" for r in sonuclar):
        s += ["UDF'yi UYAP'a yüklemeden önce UYAP Doküman Editörü'nde açıp bir kez gözle",
              "kontrol etmeniz önerilir."]
    if any(r["yon"] == "udf>word" for r in sonuclar):
        s += ["Word'e çevrilen belgede satır ve sayfa sonları Editör'dekinden farklı düşebilir;",
              "görseller UDF'deki kalitesiyle, yeniden sıkıştırılmadan aktarılır."]
    return "\n".join(s)


def klasorde_goster(yol):
    try:
        if sys.platform == "darwin":
            subprocess.run(["open", "-R", yol])
        elif os.name == "nt":
            subprocess.run(["explorer", "/select,", os.path.normpath(yol)])
        else:
            subprocess.run(["xdg-open", os.path.dirname(yol)])
    except Exception:
        pass


def dosya_ac(yol):
    try:
        if sys.platform == "darwin":
            subprocess.run(["open", yol])
        elif os.name == "nt":
            os.startfile(yol)                               # noqa: yalnız Windows
        else:
            subprocess.run(["xdg-open", yol])
    except Exception:
        pass


# --------------------------------------------------------------------------
# Pencereler
# --------------------------------------------------------------------------
class RaporPenceresi(tk.Toplevel):
    def __init__(self, ana, baslik, metin, renk, genislik=86, yukseklik=26):
        super().__init__(ana)
        self.title(baslik)
        self.transient(ana)
        cerceve = ttk.Frame(self, padding=12)
        cerceve.pack(fill="both", expand=True)
        kutu = tk.Text(cerceve, width=genislik, height=yukseklik, wrap="word",
                       font=("Menlo" if sys.platform == "darwin" else "Consolas", 11),
                       background=renk["kagit"], foreground=renk["yazi"],
                       relief="solid", borderwidth=1, padx=8, pady=8)
        kaydir = ttk.Scrollbar(cerceve, command=kutu.yview)
        kutu.configure(yscrollcommand=kaydir.set)
        kaydir.pack(side="right", fill="y")
        kutu.pack(side="left", fill="both", expand=True)
        kutu.insert("1.0", metin)
        kutu.configure(state="disabled")
        alt = ttk.Frame(self, padding=(12, 0, 12, 12))
        alt.pack(fill="x")
        self.btn = ttk.Button(alt, text="Panoya kopyala", command=lambda: self.kopyala(metin))
        self.btn.pack(side="left")
        ttk.Button(alt, text="Kapat", command=self.destroy).pack(side="right")

    def kopyala(self, metin):
        self.clipboard_clear()
        self.clipboard_append(metin)
        self.btn.configure(text="Kopyalandı ✓")
        self.after(1500, lambda: self.btn.configure(text="Panoya kopyala"))


class Uygulama(TEMEL_PENCERE):
    def __init__(self):
        super().__init__()
        self.title(UYGULAMA_ADI)
        self.geometry("620x720")
        self.minsize(600, 620)
        self.yollar = []
        self.sonuclar = []
        self.calisiyor = False
        self.hedef_klasor = None                    # None = Word dosyasının yanına
        self.renk = tema_sec(self)
        self.editor_var = bool(udf_onizle.editor_home())
        self._kur()
        try:                                        # Finder'dan uygulamaya sürükleme
            self.createcommand("::tk::mac::OpenDocument", lambda *y: self.ekle(list(y)))
        except tk.TclError:
            pass

    # ---------------------------------------------------------------- yerleşim
    def _kur(self):
        dis = ttk.Frame(self, padding=14)
        dis.pack(fill="both", expand=True)

        ust = ttk.Frame(dis)
        ust.pack(fill="x")
        ttk.Label(ust, text=UYGULAMA_ADI, font=("Helvetica", 15, "bold")).pack(side="left")
        ttk.Button(ust, text="Hakkında", width=9, command=self.hakkinda_ac).pack(side="right")
        ttk.Label(dis, text="Word belgesini UDF'ye, UDF belgesini Word'e biçimini bozmadan çevirir.",
                  foreground=self.renk["soluk"], font=("Helvetica", 11)).pack(anchor="w", pady=(2, 0))

        # --- bırakma alanı ---
        self.birak = tk.Frame(dis, highlightthickness=2, highlightbackground=self.renk["cizgi"],
                              highlightcolor=self.renk["cizgi"], bd=0)
        self.birak.pack(fill="x", pady=(12, 10))
        ic = ttk.Frame(self.birak, padding=18)
        ic.pack(fill="both", expand=True)

        # Sürükle-bırak yazısı ancak GERÇEKTEN kurulabildiyse yazılır: paket
        # içinde tkdnd kütüphanesi eksikse import başarılı olur ama kayıt çöker.
        surukleme = False
        if SURUKLENEBILIR:
            try:
                self.drop_target_register(DND_FILES)
                self.dnd_bind("<<Drop>>", self._birakildi)
                surukleme = True
            except Exception:
                surukleme = False
        self.birak_yazi = ttk.Label(ic, font=("Helvetica", 13),
                                    text="Word ya da UDF belgelerini buraya bırakın" if surukleme
                                    else "Çevrilecek Word ya da UDF belgelerini seçin")
        self.birak_yazi.pack()
        alt_yazi = ttk.Label(ic, foreground=self.renk["soluk"], font=("Helvetica", 11),
                             text="veya tıklayıp seçin · Word ve Pages → UDF · .udf → Word" if surukleme
                             else "tıklayıp seçin · Word ve Pages → UDF · .udf → Word")
        alt_yazi.pack(pady=(3, 0))
        for w in (self.birak, ic, self.birak_yazi, alt_yazi):
            w.bind("<Button-1>", lambda e: self.dosya_ekle())

        # --- dosya listesi ---
        self.liste_cerceve = ttk.Frame(dis)
        self.liste = tk.Listbox(self.liste_cerceve, height=5, activestyle="none",
                                highlightthickness=0, borderwidth=1, relief="solid",
                                selectmode="extended", font=("Helvetica", 12),
                                background=self.renk["kagit"], foreground=self.renk["yazi"],
                                selectbackground=self.renk["secim"],
                                selectforeground=self.renk["yazi"])
        self.liste.pack(fill="x")
        alt_liste = ttk.Frame(self.liste_cerceve)
        alt_liste.pack(fill="x", pady=(5, 0))
        self.sayi_yazi = ttk.Label(alt_liste, text="", foreground=self.renk["soluk"],
                                   font=("Helvetica", 11))
        self.sayi_yazi.pack(side="left")
        ttk.Button(alt_liste, text="Temizle", width=8, command=self.temizle).pack(side="right")
        ttk.Button(alt_liste, text="Çıkar", width=7, command=self.sil).pack(side="right", padx=5)

        # --- ayarlar ---
        self.ayar = ttk.LabelFrame(dis, text="Ayarlar", padding=(10, 6))
        s1 = self.ayar_word = ttk.Frame(self.ayar)        # yalnız Word → UDF için; listede .docx yoksa gizlenir
        s1.pack(fill="x")
        ttk.Label(s1, text="Sayfa numarası:", width=16).pack(side="left")
        self.sayfa_no = tk.StringVar(value="Word'deki gibi")
        ttk.Combobox(s1, textvariable=self.sayfa_no, values=list(SAYFA_NO), state="readonly",
                     width=16).pack(side="left")
        ttk.Label(s1, text="  Tablo dolgusu:").pack(side="left")
        self.dolgu = tk.StringVar(value="Renk bandı")
        ttk.Combobox(s1, textvariable=self.dolgu, values=list(DOLGU), state="readonly",
                     width=11).pack(side="left", padx=(4, 0))
        s2 = self.ayar_kayit = ttk.Frame(self.ayar)
        s2.pack(fill="x", pady=(6, 0))
        ttk.Label(s2, text="Kayıt yeri:", width=16).pack(side="left")
        self.hedef_yazi = ttk.Label(s2, text="Belgenin yanına", foreground=self.renk["soluk"])
        self.hedef_yazi.pack(side="left")
        ttk.Button(s2, text="Değiştir…", width=10, command=self.hedef_sec).pack(side="right")

        # --- eylem ve durum ---
        self.btn_cevir = ttk.Button(dis, text="Çevir", command=self.cevir_baslat)
        self.ilerleme = ttk.Progressbar(dis, mode="determinate")
        self.durum_kutu = tk.Frame(dis, highlightthickness=1, bd=0)
        durum_ic = ttk.Frame(self.durum_kutu, padding=11)
        durum_ic.pack(fill="both", expand=True)
        self.durum_baslik = ttk.Label(durum_ic, text="", font=("Helvetica", 13, "bold"))
        self.durum_baslik.pack(anchor="w")
        self.durum_detay = ttk.Label(durum_ic, text="", foreground=self.renk["soluk"],
                                     font=("Helvetica", 11), wraplength=548, justify="left")
        self.durum_detay.pack(anchor="w", pady=(2, 0))

        self.sonuc_cerceve = ttk.Frame(dis)
        ttk.Button(self.sonuc_cerceve, text="Rapor", command=self.rapor_ac).pack(side="left")
        ttk.Button(self.sonuc_cerceve, text="Klasörde göster",
                   command=self.klasor_ac).pack(side="left", padx=8)
        ttk.Button(self.sonuc_cerceve, text="Belgeyi aç", command=self.belge_ac).pack(side="left", padx=(0, 8))
        self.btn_onizle = ttk.Button(self.sonuc_cerceve, text="Editör önizlemesi",
                                     command=self.onizle_baslat)

        ttk.Label(dis, foreground=self.renk["soluk"], font=("Helvetica", 10),
                  text=f"{YAZAR} · s{SURUM}").pack(side="bottom", anchor="w")
        self.listeyi_ciz()

    # ------------------------------------------------------------ dosya işleri
    def _birakildi(self, olay):
        self.ekle(self.tk.splitlist(olay.data))

    def dosya_ekle(self):
        if self.calisiyor:
            return
        secilen = filedialog.askopenfilenames(
            title="Word, Pages ya da UDF belgelerini seçin",
            filetypes=[("Word, Pages ve UDF belgeleri", "*.docx *.doc *.rtf *.odt *.pages *.udf"),
                       ("Word belgesi", "*.docx *.doc *.rtf *.odt"), ("Pages belgesi", "*.pages"),
                       ("UYAP belgesi", "*.udf"), ("Tüm dosyalar", "*.*")])
        self.ekle(list(secilen))

    def ekle(self, yeni):
        if self.calisiyor:
            return
        atlanan = []
        for y in yeni:
            y = os.path.abspath(str(y))
            y = y.rstrip("/\\")
            if os.path.isdir(y) and not y.lower().endswith(".pages"):   # eski .pages belgesi bir klasördür
                self.ekle([os.path.join(y, f) for f in sorted(os.listdir(y))
                           if f.lower().endswith(KABUL) and not f.startswith("~$")])
                continue
            if not y.lower().endswith(KABUL) or os.path.basename(y).startswith("~$"):
                atlanan.append(os.path.basename(y))
                continue
            if y not in self.yollar and os.path.exists(y):
                self.yollar.append(y)
        self.sifirla()
        self.listeyi_ciz()
        if atlanan:
            self.durum_goster("sari", "Bazı dosyalar eklenmedi",
                              "Yalnız Word (.docx, .doc, .rtf, .odt), Pages ve UDF belgeleri çevrilir: "
                              + ", ".join(atlanan[:4]) + (" …" if len(atlanan) > 4 else ""))

    def sil(self):
        if self.calisiyor:
            return
        for i in reversed(self.liste.curselection()):
            del self.yollar[i]
        self.sifirla()
        self.listeyi_ciz()

    def temizle(self):
        if self.calisiyor:
            return
        self.yollar = []
        self.sifirla()
        self.listeyi_ciz()

    def listeyi_ciz(self):
        self.liste.delete(0, "end")
        durum = {r["kaynak"]: r for r in self.sonuclar}
        for y in self.yollar:
            r = durum.get(y)
            im = "   " if r is None else ("✓ " if r["tamam"] else "✕ ")
            yon = "→ Word" if y.lower().endswith(".udf") else "→ UDF"
            self.liste.insert("end", f" {im}{os.path.basename(y)}   {yon}")
            if r is not None:
                self.liste.itemconfigure("end", foreground=self.renk["yesil" if r["tamam"] else "kirmizi"])
        if self.yollar:
            self.liste_cerceve.pack(fill="x", after=self.birak)
            self.ayar.pack(fill="x", pady=(10, 0), after=self.liste_cerceve)
            self.btn_cevir.pack(pady=(12, 0), after=self.ayar)
            self.sayi_yazi.configure(text=f"{len(self.yollar)} belge")
            if any(not y.lower().endswith(".udf") for y in self.yollar):
                self.ayar_word.pack(fill="x", before=self.ayar_kayit)
            else:
                self.ayar_word.pack_forget()
        else:
            for w in (self.liste_cerceve, self.ayar, self.btn_cevir):
                w.pack_forget()

    def sifirla(self):
        self.sonuclar = []
        for w in (self.durum_kutu, self.sonuc_cerceve, self.ilerleme):
            w.pack_forget()

    def hedef_sec(self):
        k = filedialog.askdirectory(title="Çevrilen belgelerin kaydedileceği klasör")
        if k:
            self.hedef_klasor = k
            self.hedef_yazi.configure(text=k if len(k) < 44 else "…" + k[-42:])
        else:
            self.hedef_klasor = None
            self.hedef_yazi.configure(text="Belgenin yanına")

    def durum_goster(self, tur, baslik, detay):
        renk = self.renk[tur]
        self.durum_kutu.configure(highlightbackground=renk, highlightcolor=renk)
        self.durum_baslik.configure(text=baslik, foreground=renk)
        self.durum_detay.configure(text=detay)
        self.durum_kutu.pack(fill="x", pady=(12, 0))

    # ---------------------------------------------------------------- çevirme
    def cevir_baslat(self):
        if self.calisiyor or not self.yollar:
            return
        var_olan = []
        for y in self.yollar:
            if y.lower().endswith(".udf"):                  # Word çıktısı hiçbir zaman var olanın üzerine yazmaz
                continue
            hedef = os.path.join(self.hedef_klasor or os.path.dirname(y),
                                 os.path.splitext(os.path.basename(y))[0] + ".udf")
            if os.path.exists(hedef):
                var_olan.append(os.path.basename(hedef))
        if var_olan and not messagebox.askyesno(
                UYGULAMA_ADI, "Şu UDF dosyaları zaten var, üzerine yazılsın mı?\n\n"
                + "\n".join(var_olan[:6]) + ("\n…" if len(var_olan) > 6 else "")):
            return
        self.sifirla()
        self.calisiyor = True
        self.btn_cevir.configure(state="disabled", text="Çevriliyor…")
        self.ilerleme.configure(maximum=len(self.yollar), value=0)
        self.ilerleme.pack(fill="x", pady=(10, 0), after=self.btn_cevir)
        ayar = (self.hedef_klasor, SAYFA_NO[self.sayfa_no.get()], DOLGU[self.dolgu.get()])
        threading.Thread(target=self._cevir_isi, args=(list(self.yollar),) + ayar, daemon=True).start()

    def _cevir_isi(self, yollar, klasor, sayfa_no, dolgu):
        sonuclar = []
        for i, y in enumerate(yollar, 1):
            sonuclar.append(belge_cevir(y, klasor, sayfa_no, dolgu))
            self.after(0, lambda n=i: self.ilerleme.configure(value=n))
        self.after(0, lambda: self._cevir_bitti(sonuclar))

    def _cevir_bitti(self, sonuclar):
        self.calisiyor = False
        self.sonuclar = sonuclar
        self.btn_cevir.configure(state="normal", text="Çevir")
        self.ilerleme.pack_forget()
        self.listeyi_ciz()
        iyi = [r for r in sonuclar if r["tamam"]]
        kotu = [r for r in sonuclar if not r["tamam"]]
        self._sonuc_yaz(iyi, kotu, 4)
        self.sonuc_cerceve.pack(pady=(10, 0))
        if self.editor_var and any(r["yon"] == "word>udf" for r in iyi):
            self.btn_onizle.pack(side="left")
            self.btn_onizle.configure(state="normal")
        else:
            self.btn_onizle.pack_forget()
        self._sigdir()
        if self.winfo_reqheight() > self.winfo_screenheight() - 90:     # küçük ekran: daha az madde göster
            self._sonuc_yaz(iyi, kotu, 2)
            self._sigdir()

    def _sonuc_yaz(self, iyi, kotu, en_cok):
        """Sonuç kutusunun metni; uyarılar sade dille, en çok en_cok madde."""
        gosterilecek, kalan = sade_uyarilar(iyi, en_cok)
        yonler = {r["yon"] for r in iyi}
        farklar = ""
        if gosterilecek:                                    # uyarılar kutuda, sade dille; teknik metin Rapor'da
            bas = ("Word'den farklı olan yerler:" if yonler == {"word>udf"} else
                   "Editör'dekinden farklı olan yerler:" if yonler == {"udf>word"} else "Dikkat edilecek yerler:")
            farklar = ("\n\n" + bas + "\n" + "\n".join("•  " + u for u in gosterilecek)
                       + (f"\n•  … ve {kalan} husus daha (Rapor'da)" if kalan else ""))
        son = []
        if "word>udf" in yonler:
            son.append("UYAP'a yüklemeden önce UDF'yi Editör'de açıp bir kez kontrol edin.")
        if "udf>word" in yonler:
            yeni_ad = [os.path.basename(r["cikti"]) for r in iyi if r["yon"] == "udf>word" and "(UDF'den" in r["cikti"]]
            if yeni_ad:
                son.append("Aynı adda bir Word belgesi vardı; üzerine yazılmadı, yeni belge \""
                           + yeni_ad[0] + "\" adıyla kaydedildi.")
            son.append("Word'de satır ve sayfa sonları Editör'dekinden biraz farklı düşebilir.")
        if len(son) > 2:                                    # karışık çeviride kutu uzamasın
            son = son[:2]
        if iyi and not kotu:
            baslik = "Belge çevrildi" if len(iyi) == 1 else f"{len(iyi)} belge çevrildi"
            detay = "Doğrulama geçti; metin eksiksiz." + farklar + "\n\n" + "\n".join(son)
            self.durum_goster("yesil" if not gosterilecek else "sari", baslik, detay)
        elif iyi:
            self.durum_goster("sari", f"{len(iyi)} belge çevrildi, {len(kotu)} belge çevrilemedi",
                              "İlk sorun: " + kotu[0]["hata"] + farklar + "\n\nAyrıntı Rapor'da.")
        else:
            self.durum_goster("kirmizi", "Çevrilemedi",
                              kotu[0]["hata"] if kotu else "Bilinmeyen bir sorun oluştu.")

    def _sigdir(self):
        """Sonuç kutusu uzunsa pencereyi içeriğe göre uzatır (düğmeler alta taşıp kaybolmasın)."""
        self.update_idletasks()
        en = max(self.winfo_width(), self.winfo_reqwidth())
        boy = min(max(self.winfo_height(), self.winfo_reqheight()), self.winfo_screenheight() - 90)
        if (en, boy) != (self.winfo_width(), self.winfo_height()):
            self.geometry(f"{en}x{boy}")

    # ------------------------------------------------------------------ sonuç
    def rapor_ac(self):
        if self.sonuclar:
            RaporPenceresi(self, "Çeviri raporu", rapor_metni(self.sonuclar), self.renk)

    def _secili_sonuc(self):
        iyi = [r for r in self.sonuclar if r["tamam"]]
        sec = self.liste.curselection()
        if sec:
            y = self.yollar[sec[0]]
            for r in iyi:
                if r["kaynak"] == y:
                    return r
        return iyi[0] if iyi else None

    def klasor_ac(self):
        r = self._secili_sonuc()
        if r:
            klasorde_goster(r["cikti"])

    def belge_ac(self):
        r = self._secili_sonuc()
        if r:
            dosya_ac(r["cikti"])

    def onizle_baslat(self):
        r = self._secili_sonuc()
        if r and r["yon"] != "word>udf":                    # önizleme UDF çıktısı içindir
            r = next((x for x in self.sonuclar if x["tamam"] and x["yon"] == "word>udf"), None)
        if not r or self.calisiyor:
            return
        self.calisiyor = True
        self.btn_onizle.configure(state="disabled", text="Çiziliyor…")
        threading.Thread(target=self._onizle_isi, args=(r["cikti"],), daemon=True).start()

    def _onizle_isi(self, udf):
        try:
            klasor = tempfile.mkdtemp(prefix="udf_onizleme_")
            _, sayfalar = udf_onizle.onizle(udf, klasor, olcek=1.4)
            hata = None
        except Exception as e:
            sayfalar, hata = [], str(e)
        self.after(0, lambda: self._onizle_bitti(sayfalar, hata))

    def _onizle_bitti(self, sayfalar, hata):
        self.calisiyor = False
        self.btn_onizle.configure(state="normal", text="Editör önizlemesi")
        if hata or not sayfalar:
            messagebox.showwarning(UYGULAMA_ADI, "Önizleme çizilemedi.\n\n" + (hata or ""))
            return
        if sys.platform == "darwin":                        # Önizleme'de tek pencerede, sayfa sayfa
            subprocess.run(["open"] + sayfalar)
        else:
            dosya_ac(sayfalar[0])
            if len(sayfalar) > 1:
                klasorde_goster(sayfalar[0])

    # --------------------------------------------------------------- hakkında
    def hakkinda_ac(self):
        p = tk.Toplevel(self)
        p.title("Hakkında")
        p.transient(self)
        p.resizable(False, False)
        c = ttk.Frame(p, padding=20)
        c.pack(fill="both", expand=True)
        ttk.Label(c, text=UYGULAMA_ADI, font=("Helvetica", 16, "bold")).pack()
        ttk.Label(c, text=f"Sürüm {SURUM}", foreground=self.renk["soluk"]).pack(pady=(2, 10))
        ttk.Label(c, justify="center", wraplength=380, text=(
            "Word (.docx) belgesini; tabloları, görselleri, listeleri, üstbilgi ve altbilgisiyle "
            "UYAP Doküman Editörü'nün UDF biçimine, UDF belgesini de aynı şekilde Word'e çevirir. "
            ".doc, .rtf, .odt ve Pages belgeleri bilgisayardaki Word ya da Pages ile açılarak çevrilir. "
            "Her çıktı üretildikten sonra kaynağıyla karşılaştırılarak doğrulanır.\n\n"
            "Belgeler bilgisayarınızdan çıkmaz; "
            "program kendiliğinden internete bağlanmaz.")).pack()
        ttk.Separator(c).pack(fill="x", pady=12)
        ttk.Label(c, text=YAZAR, font=("Helvetica", 12, "bold")).pack()
        ttk.Label(c, text=YAZAR_EK, foreground=self.renk["soluk"]).pack()
        ttk.Label(c, text=LISANS, foreground=self.renk["soluk"]).pack(pady=(8, 0))
        ttk.Label(c, foreground=self.renk["soluk"], font=("Helvetica", 10), justify="center", text=(
            "Editör önizlemesi: " + ("kullanılabilir (UYAP Doküman Editörü bulundu)" if self.editor_var
                                     else "kapalı (UYAP Doküman Editörü bulunamadı)")
            + "\nBu program UYAP veya Adalet Bakanlığı ile bağlantılı değildir.")).pack(pady=(8, 0))
        if DEPO:
            yazi = ttk.Label(c, text="", foreground=self.renk["soluk"])
            btn_indir = ttk.Button(c, text="İndirme sayfasını aç",
                                   command=lambda: webbrowser.open(SURUM_SAYFA))
            dugme = ttk.Button(c, text="Güncellemeleri kontrol et")
            dugme.configure(command=lambda: self.guncelleme_kontrol(p, dugme, yazi, btn_indir))
            dugme.pack(pady=(12, 0))
            yazi.pack(pady=(6, 0))
        ttk.Button(c, text="Kapat", command=p.destroy).pack(pady=(14, 0))

    def guncelleme_kontrol(self, pencere, dugme, yazi, btn_indir):
        dugme.configure(state="disabled", text="Denetleniyor…")

        def bitti(etiket, hata):
            if not pencere.winfo_exists():
                return
            dugme.configure(state="normal", text="Güncellemeleri kontrol et")
            if hata:
                yazi.configure(text="Denetlenemedi (internet bağlantısı yok olabilir).")
            elif surum_sayilari(etiket) > surum_sayilari(SURUM):
                yazi.configure(text=f"Yeni sürüm var: {etiket}")
                btn_indir.pack(pady=(6, 0))
            else:
                yazi.configure(text="En güncel sürümü kullanıyorsunuz.")

        def is_parcacigi():
            try:
                e, h = son_surumu_sor(), None
            except (urllib.error.URLError, OSError, ValueError) as ex:
                e, h = "", ex
            self.after(0, lambda: bitti(e, h))
        threading.Thread(target=is_parcacigi, daemon=True).start()


# --------------------------------------------------------------------------
# Öz sınama
# --------------------------------------------------------------------------
def _ornek_docx(yol):
    """Sınama için küçük bir Word belgesi: biçimli paragraf + 2x2 tablo + PNG görsel."""
    import struct
    import zlib

    def parca(t, d):
        c = struct.pack(">I", len(d)) + t + d
        return c + struct.pack(">I", zlib.crc32(t + d) & 0xffffffff)
    ham = b"".join(b"\x00" + bytes((31, 58, 95)) * 24 for _ in range(12))
    png = (b"\x89PNG\r\n\x1a\n" + parca(b"IHDR", struct.pack(">IIBBBBB", 24, 12, 8, 2, 0, 0, 0))
           + parca(b"IDAT", zlib.compress(ham)) + parca(b"IEND", b""))
    ns = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
          'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
          'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
          'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')
    gorsel = ('<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="g"/>'
              '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
              '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name=""/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill>'
              '<a:blip r:embed="rId9"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic>'
              '</wp:inline></w:drawing></w:r>')
    hucre = lambda t: f'<w:tc><w:tcPr/><w:p><w:r><w:t>{t}</w:t></w:r></w:p></w:tc>'
    govde = ('<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr>'
             '<w:t>SINAMA BELGESİ</w:t></w:r></w:p>'
             '<w:p><w:r><w:t xml:space="preserve">Türkçe karakterler: ğüşıöç İĞÜŞÖÇ ve </w:t></w:r>'
             '<w:r><w:rPr><w:i/></w:rPr><w:t>italik</w:t></w:r><w:r><w:t xml:space="preserve">, CO</w:t></w:r>'
             '<w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>2</w:t></w:r></w:p>'
             '<w:tbl><w:tblPr><w:tblBorders><w:insideH w:val="single"/><w:insideV w:val="single"/>'
             '<w:top w:val="single"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/>'
             '<w:gridCol w:w="6000"/></w:tblGrid>'
             f'<w:tr>{hucre("Tarih")}{hucre("Olay")}</w:tr><w:tr>{hucre("01.01.2026")}{hucre("Tebliğ")}</w:tr>'
             f'</w:tbl><w:p>{gorsel}</w:p><w:p><w:r><w:t>Son paragraf.</w:t></w:r></w:p>')
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    with zipfile.ZipFile(yol, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/'
                   'package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.'
                   'openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType='
                   '"application/xml"/><Default Extension="png" ContentType="image/png"/><Override '
                   'PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.'
                   'wordprocessingml.document.main+xml"/></Types>')
        z.writestr("_rels/.rels", f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.'
                   f'openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{rel}/'
                   'officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/_rels/document.xml.rels", f'<?xml version="1.0" encoding="UTF-8"?><Relationships '
                   f'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship '
                   f'Id="rId9" Type="{rel}/image" Target="media/g.png"/></Relationships>')
        z.writestr("word/document.xml", f'<?xml version="1.0" encoding="UTF-8"?><w:document {ns}><w:body>'
                   f'{govde}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" '
                   'w:bottom="1417" w:left="1417"/></w:sectPr></w:body></w:document>')
        z.writestr("word/media/g.png", png)


def sinama(yollar=None, rapor_yolu=None):
    """Paketin bütünlüğünü doğrular (pencere göstermeden).

    Windows'ta --windowed derlenen exe'nin KONSOLU YOKTUR; print çıktısı hiçbir
    yere gitmez. Bu yüzden sonuç istenirse bir dosyaya da yazılır.
    """
    satirlar = []

    def yaz(*p):
        metin = " ".join(str(x) for x in p)
        satirlar.append(metin)
        print(metin)
    kod = 0
    try:
        assert tk.TkVersion >= 8.6, f"Tcl/Tk {tk.TkVersion} çok eski (8.6+ gerekir)"
        klasor = tempfile.mkdtemp(prefix="udf_cevirici_sinama_")
        ornek = os.path.join(klasor, "sınama belgesi.docx")
        _ornek_docx(ornek)
        for y in [ornek] + list(yollar or []):
            r = belge_cevir(y, klasor)
            yaz(f"{'✓' if r['tamam'] else '✕'} {os.path.basename(y)}"
                + ("" if r["tamam"] else f" — {r['hata']}"))
            for u in r["uyarilar"]:
                yaz(f"    ! {u}")
            for d in r["dogrulama"]:
                yaz(f"    {d}")
            assert r["tamam"] or y != ornek, "örnek belge çevrilemedi: " + r["hata"]
        o = belge_cevir(ornek, klasor)["ozet"]
        assert o["tablo"] == 1 and o["paragraf"] >= 4, f"örnek belgenin yapısı beklenenden farklı: {o}"
        with zipfile.ZipFile(os.path.join(klasor, "sınama belgesi.udf")) as z:
            xml = z.read("content.xml").decode("utf-8")
        assert "<image " in xml and "<table " in xml and "ğüşıöç İĞÜŞÖÇ" in xml, "UDF içeriği eksik"
        kotu = belge_cevir(os.path.join(klasor, "yok.docx"), klasor)
        assert not kotu["tamam"] and kotu["hata"], "olmayan dosya için hata üretilmedi"

        # ters yön: UDF → Word. Asıl Word belgesi aynı klasörde duruyor; üzerine YAZILMAMALI.
        with open(ornek, "rb") as f:
            asil = f.read()
        t = belge_cevir(os.path.join(klasor, "sınama belgesi.udf"), klasor)
        yaz(f"{'✓' if t['tamam'] else '✕'} sınama belgesi.udf → {os.path.basename(t['cikti'])}"
            + ("" if t["tamam"] else f" — {t['hata']}"))
        for d in t["dogrulama"]:
            yaz(f"    {d}")
        assert t["tamam"], "UDF → Word çevirisi başarısız: " + t["hata"]
        assert os.path.basename(t["cikti"]) == "sınama belgesi (UDF'den).docx", "Word çıktısının adı: " + t["cikti"]
        with open(ornek, "rb") as f:
            assert f.read() == asil, "asıl Word belgesinin üzerine yazıldı"
        with zipfile.ZipFile(t["cikti"]) as z:
            govde = z.read("word/document.xml").decode("utf-8")
            assert any(n.startswith("word/media/") for n in z.namelist()), "Word çıktısında görsel yok"
        assert "<w:tbl>" in govde and "ğüşıöç İĞÜŞÖÇ" in govde, "Word çıktısının içeriği eksik"
        kotu = belge_cevir(os.path.join(klasor, "yok.udf"), klasor)
        assert not kotu["tamam"] and kotu["hata"], "olmayan UDF için hata üretilmedi"

        # Arayüz kuruluyor ve bir belgeyi uçtan uca çevirebiliyor mu? (pencere gösterilmez)
        uyg = Uygulama()
        uyg.withdraw()
        uyg.ekle([ornek])
        assert uyg.yollar == [os.path.abspath(ornek)], "belge listeye eklenemedi"
        uyg._cevir_bitti([belge_cevir(ornek, klasor)])
        uyg.update()
        assert uyg.sonuclar and uyg.sonuclar[0]["tamam"], "arayüz üzerinden çeviri başarısız"
        uyg.ekle([os.path.join(klasor, "sınama belgesi.udf")])
        assert len(uyg.yollar) == 2, "UDF belgesi listeye eklenemedi"
        ornek_uyari = [{"uyarilar": ["dipnotlar belge sonuna numaralı paragraf olarak taşındı (…)",
                                     "bu UDF e-imzalı; üretilen Word belgesi imza taşımaz (çalışma kopyasıdır)",
                                     "bilinmeyen yeni bir uyarı"]}]
        sade, kalan = sade_uyarilar(ornek_uyari)
        assert sade[0].startswith("Dipnotlar sayfa altında") and sade[1].startswith("Bu UDF e-imzalıydı;") \
            and sade[2] == "Bilinmeyen yeni bir uyarı" and kalan == 0, f"uyarı sadeleştirme bozuk: {sade}"
        dnd = "yok"
        if SURUKLENEBILIR:
            try:
                uyg.drop_target_register(DND_FILES)
                dnd = "çalışıyor"
            except Exception as e:
                dnd = f"KURULAMADI ({type(e).__name__})"
        editor = "bulundu" if uyg.editor_var else "yok"
        uyg.destroy()
        if editor == "bulundu":                              # paketteki Java sınıfları yerinde mi?
            try:
                _, sayfalar = udf_onizle.onizle(os.path.join(klasor, "sınama belgesi.udf"),
                                                os.path.join(klasor, "onizleme"))
                editor += f", önizleme çalışıyor ({len(sayfalar)} sayfa)"
            except Exception as e:
                editor += f", önizleme ÇALIŞMADI ({e})"
        assert ofis_docx.destekli("x.doc") and ofis_docx.destekli("x.pages") and not ofis_docx.destekli("x.pdf"), \
            "ofis biçimleri tanınmıyor"
        if sys.platform == "darwin":
            ofis = ", ".join(f"{ad}: {'var' if ofis_docx._mac_uygulama_var(b) else 'yok'}"
                             for ad, b in (("Word", ofis_docx.WORD_MAC), ("Pages", ofis_docx.PAGES_MAC)))
        else:
            ofis = "Word: çeviride denenecek"
        try:
            import PIL                                       # noqa
            pil = "Pillow " + PIL.__version__
        except Exception:
            pil = "Pillow yok (JPEG için sistem aracı denenir)"
        yaz(f"SINAMA TAMAM — iki yön, çekirdek ve arayüz çalışıyor (Tcl/Tk {tk.TkVersion}, "
            f"sürükle-bırak: {dnd}, {pil}, {ofis}, UYAP Editör: {editor}).")
    except AssertionError as e:
        yaz(f"SINAMA BAŞARISIZ — {e}")
        kod = 1
    except Exception as e:
        yaz(f"SINAMA BAŞARISIZ — beklenmeyen hata: {type(e).__name__}: {e}")
        kod = 1
    if rapor_yolu:
        try:
            with open(rapor_yolu, "w", encoding="utf-8") as f:
                f.write("\n".join(satirlar) + "\n")
        except OSError as e:
            print("rapor yazılamadı:", e, file=sys.stderr)
    return kod


def main():
    if "--sinama" in sys.argv:
        rapor = None
        if "--rapor" in sys.argv:
            i = sys.argv.index("--rapor")
            rapor = sys.argv[i + 1] if i + 1 < len(sys.argv) else None
        return sinama([y for y in sys.argv[1:] if y.lower().rstrip("/").endswith(KABUL)], rapor)
    if tk.TkVersion < 8.6:
        print(f"UYARI: Tcl/Tk {tk.TkVersion} çok eski; pencere boş görünebilir.", file=sys.stderr)
    uyg = Uygulama()
    uyg.ekle([y for y in sys.argv[1:] if os.path.exists(y)])
    uyg.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
