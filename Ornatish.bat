@echo off
setlocal
title Lotin-Kirill konvertor - o'rnatish
echo Lotin-Kirill konvertor (Excel uchun) / Uzbek Latin-Cyrillic converter for Excel
echo Muallif / Author: Begzodbek Akramjonov
echo.
set "SRC=%~dp0Konvertor.xlam"
set "DST=%APPDATA%\Microsoft\Excel\XLSTART"
if not exist "%SRC%" (
  echo Konvertor.xlam topilmadi. Arxivdagi hamma fayllarni bitta papkaga chiqarib, keyin qayta urinib ko'ring.
  pause
  exit /b 1
)
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>nul | find /I "EXCEL.EXE" >nul && (
  echo Excel ochiq turibdi. Uni yoping, keyin istalgan tugmani bosing...
  pause >nul
)
if not exist "%DST%" mkdir "%DST%"
copy /Y "%SRC%" "%DST%\Konvertor.xlam" >nul
if errorlevel 1 (
  echo Faylni ko'chirib bo'lmadi.
  pause
  exit /b 1
)
powershell -NoProfile -Command "Unblock-File -LiteralPath (Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART\Konvertor.xlam')" >nul 2>&1
echo Tayyor! Excel'ni oching - yuqorida "Lotin-Kirill" bo'limi chiqadi.
echo.
pause
