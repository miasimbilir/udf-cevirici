#!/bin/bash
# UDF Çevirici — macOS için .app üretir (çalıştığı Mac'in mimarisi için).
# macOS'un sistem Python'u eski Tcl/Tk 8.5 ile gelir ve pencereyi BOŞ çizer; güncel Tk'li
# bir Python gerekir:   brew install python-tk@3.13
set -eo pipefail
cd "$(dirname "$0")"
PY="${PYTHON:-/opt/homebrew/bin/python3.13}"
[ -x "$PY" ] || PY="$(command -v python3.13 || command -v python3)"
echo "Python: $PY  ($("$PY" -c 'import tkinter,sys;print(sys.version.split()[0],"Tcl/Tk",tkinter.TkVersion)'))"
"$PY" -c "import tkinter,sys; sys.exit(0 if tkinter.TkVersion>=8.6 else 1)" || { echo "HATA: Tcl/Tk 8.6+ gerekir"; exit 1; }

[ -d .venv ] || "$PY" -m venv .venv
.venv/bin/pip install --quiet --upgrade pip pyinstaller tkinterdnd2 pillow

rm -rf build dist
.venv/bin/pyinstaller --noconfirm --clean --windowed \
  --name "UDF Cevirici" --icon ikon/uygulama.icns \
  --collect-all tkinterdnd2 \
  --add-data "udf_onizle:udf_onizle" \
  --osx-bundle-identifier "tr.arabuluculuk.udfcevirici" \
  uygulama.py >/dev/null

.venv/bin/python plist_yaz.py
# Info.plist değişti: paketin imzası yeniden (ad-hoc) atılmazsa Apple Silicon'da uygulama açılmaz
codesign --force --deep --sign - "dist/UDF Cevirici.app"
codesign --verify --deep "dist/UDF Cevirici.app" && echo "imza geçerli (ad-hoc)"

echo "--- sınama:"
rm -f sinama.txt                                   # önceki derlemenin raporu "geçti" sayılmasın
"dist/UDF Cevirici.app/Contents/MacOS/UDF Cevirici" --sinama --rapor sinama.txt | tail -3
grep -q "SINAMA TAMAM" sinama.txt || { echo "HATA: sınama başarısız; paketi dağıtmayın"; exit 1; }

MIMARI=$([ "$(uname -m)" = "arm64" ] && echo AppleSilicon || echo Intel)
SURUM=$(.venv/bin/python -c "from uygulama import SURUM; print(SURUM)")
PAKET="dist/UDF-Cevirici-$SURUM-macOS-$MIMARI.zip"
rm -rf dist/paket && mkdir dist/paket && cp -R "dist/UDF Cevirici.app" dist/paket/ && cp KULLANIM.txt LICENSE dist/paket/
(cd dist/paket && zip -r -y -q "../$(basename "$PAKET")" .)
rm -rf dist/paket
echo "TAMAM: $PAKET"
