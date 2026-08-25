@echo off
REM Bu dosyayi Windows'ta calistirin. Python 3 kurulu olmalidir (python.org).
REM Tek exe olusturur: dist\KoliEtiketi.exe
pip install --upgrade openpyxl reportlab pyinstaller
pyinstaller --onefile --windowed --name KoliEtiketi etiket_uygulamasi.py
echo.
echo Bitti! Exe dosyasi: dist\KoliEtiketi.exe
pause
