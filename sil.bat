@echo off
title Image Converter Kaldırma
color 0A

:: Sessiz mod - pencere gizli
if "%1" neq "admin" (
    powershell -Command "Start-Process '%0' -ArgumentList 'admin' -Verb RunAs -WindowStyle Hidden"
    exit /b
)

:: Registry temizleme (sessiz)
reg delete "HKEY_CLASSES_ROOT\SystemFileAssociations\image\shell\ImageConverter" /f >nul 2>&1
reg delete "HKEY_CLASSES_ROOT\.jpg\shell\ImageConverter" /f >nul 2>&1
reg delete "HKEY_CLASSES_ROOT\.jpeg\shell\ImageConverter" /f >nul 2>&1
reg delete "HKEY_CLASSES_ROOT\.png\shell\ImageConverter" /f >nul 2>&1
reg delete "HKEY_CLASSES_ROOT\.bmp\shell\ImageConverter" /f >nul 2>&1
reg delete "HKEY_CLASSES_ROOT\.gif\shell\ImageConverter" /f >nul 2>&1
reg delete "HKEY_CLASSES_ROOT\.tif\shell\ImageConverter" /f >nul 2>&1
reg delete "HKEY_CLASSES_ROOT\.tiff\shell\ImageConverter" /f >nul 2>&1
reg delete "HKEY_CURRENT_USER\Software\ImageConverter" /f >nul 2>&1
reg delete "HKEY_CLASSES_ROOT\*\shell\ImageConverter" /f >nul 2>&1

:: Explorer'ı GÜVENLE yenile (kapatmadan)
tasklist /fi "imagename eq explorer.exe" | find /i "explorer.exe" >nul
if %errorlevel% eql 0 (
    for /f "tokens=2 delims= " %%i in ('tasklist /fi "imagename eq explorer.exe" /fo table /nh') do (
        if not "%%i"=="INFO:" (
            taskkill /pid %%i /f >nul 2>&1
            timeout /t 1 /nobreak >nul
            start explorer.exe >nul 2>&1
            goto :done
        )
    )
)

:done
exit