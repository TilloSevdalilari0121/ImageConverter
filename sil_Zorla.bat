@echo off
title Image Converter Kaldırma Aracı
color 0A

echo ===============================================
echo   IMAGE CONVERTER KALDIRMA ARACI
echo ===============================================
echo.

:: Yönetici kontrolü
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo HATA: Bu dosya yonetici olarak calistirilmalidir!
    pause
    exit
)

echo Yonetici haklari onaylandi...
echo.

:: Tüm registry temizleme işlemleri
echo [1/10] SystemFileAssociations temizleniyor...
reg delete "HKEY_CLASSES_ROOT\SystemFileAssociations\image\shell\ImageConverter" /f >nul 2>&1

echo [2/10] JPG uzantisi temizleniyor...
reg delete "HKEY_CLASSES_ROOT\.jpg\shell\ImageConverter" /f >nul 2>&1

echo [3/10] JPEG uzantisi temizleniyor...
reg delete "HKEY_CLASSES_ROOT\.jpeg\shell\ImageConverter" /f >nul 2>&1

echo [4/10] PNG uzantisi temizleniyor...
reg delete "HKEY_CLASSES_ROOT\.png\shell\ImageConverter" /f >nul 2>&1

echo [5/10] BMP uzantisi temizleniyor...
reg delete "HKEY_CLASSES_ROOT\.bmp\shell\ImageConverter" /f >nul 2>&1

echo [6/10] GIF uzantisi temizleniyor...
reg delete "HKEY_CLASSES_ROOT\.gif\shell\ImageConverter" /f >nul 2>&1

echo [7/10] TIF uzantisi temizleniyor...
reg delete "HKEY_CLASSES_ROOT\.tif\shell\ImageConverter" /f >nul 2>&1

echo [8/10] TIFF uzantisi temizleniyor...
reg delete "HKEY_CLASSES_ROOT\.tiff\shell\ImageConverter" /f >nul 2>&1

echo [9/10] Kullanici ayarlari temizleniyor...
reg delete "HKEY_CURRENT_USER\Software\ImageConverter" /f >nul 2>&1

echo [10/10] Tum dosya turleri temizleniyor...
reg delete "HKEY_CLASSES_ROOT\*\shell\ImageConverter" /f >nul 2>&1

echo.
echo Registry temizligi tamamlandi.
echo.

:: Shell değişikliklerini bildir
echo Shell degisiklikleri bildiriliyor...
reg add "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "EnableBalloonTips" /t REG_DWORD /d 0 /f >nul 2>&1

:: Explorer'ı yeniden başlat
echo Explorer yeniden baslatiliyor...
taskkill /f /im explorer.exe >nul 2>&1
timeout /t 3 /nobreak >nul
start explorer.exe

:: Shell cache temizleme
echo Shell cache temizleniyor...
ie4uinit.exe -ClearIconCache >nul 2>&1

echo.
echo ===============================================
echo   KALDIRMA ISLEMI TAMAMLANDI
echo ===============================================
echo.
echo Tum registry kayitlari temizlendi.
echo Explorer yeniden baslatildi.
echo Shell cache temizlendi.
echo.
echo Bu pencere 10 saniye sonra kapanacak...
timeout /t 10 /nobreak >nul