@echo off

:: Sessiz yükleme - hiç pencere açılmasın
if "%1" neq "admin" (
    powershell -Command "Start-Process '%0' -ArgumentList 'admin' -Verb RunAs -WindowStyle Hidden"
    exit /b
)

:: Program dosyasının yolunu al
set "PROGRAM_PATH=%~dp0ImageConverter.exe"

:: Eğer program dosyası yoksa çık
if not exist "%PROGRAM_PATH%" exit /b

:: Registry kayıtlarını sessizce oluştur
reg add "HKEY_CLASSES_ROOT\SystemFileAssociations\image\shell\ImageConverter" /ve /d "Image Converter" /f >nul 2>&1
reg add "HKEY_CLASSES_ROOT\SystemFileAssociations\image\shell\ImageConverter" /v "Icon" /d "%PROGRAM_PATH%" /f >nul 2>&1
reg add "HKEY_CLASSES_ROOT\SystemFileAssociations\image\shell\ImageConverter\command" /ve /d "\"%PROGRAM_PATH%\" \"%%1\"" /f >nul 2>&1

:: Varsayılan ayarları oluştur
reg add "HKEY_CURRENT_USER\Software\ImageConverter" /v "NamingStyle" /t REG_DWORD /d 0 /f >nul 2>&1

exit