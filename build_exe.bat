@echo off
echo ========================================
echo Compilador de GLOSAS Automatizador
echo Hospital Universitario del Valle
echo ========================================
echo.

REM Verificar que existe credentials.json
if not exist "credentials.json" (
    echo ERROR: No se encontro credentials.json
    echo Por favor coloca el archivo credentials.json en esta carpeta
    pause
    exit /b 1
)

echo [1/3] Instalando dependencias...
pip install -r requirements.txt

echo.
echo [2/3] Limpiando compilaciones anteriores...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "glosas_file_downloader.spec" del /q "glosas_file_downloader.spec"

echo.
echo [3/3] Compilando ejecutable con PyInstaller...
pyinstaller --noconfirm ^
    --onefile ^
    --windowed ^
    --icon=NONE ^
    --name="GlosasAutomatizador" ^
    --add-data="credentials.json;." ^
    --hidden-import=pkg_resources.py2_warn ^
    --hidden-import=google.auth ^
    --hidden-import=google.auth.transport.requests ^
    --hidden-import=google.oauth2.credentials ^
    --hidden-import=google_auth_oauthlib.flow ^
    --hidden-import=googleapiclient.discovery ^
    --collect-all=google ^
    --collect-all=googleapiclient ^
    --collect-all=google_auth_oauthlib ^
    glosas_file_downloader.py

echo.
if exist "dist\GlosasAutomatizador.exe" (
    echo ========================================
    echo   COMPILACION EXITOSA!
    echo ========================================
    echo.
    echo El ejecutable se encuentra en:
    echo   dist\GlosasAutomatizador.exe
    echo.
    echo IMPORTANTE: Para distribuir el programa:
    echo 1. Copia el archivo dist\GlosasAutomatizador.exe
    echo 2. Copia el archivo credentials.json
    echo 3. Envia AMBOS archivos al usuario
    echo 4. El usuario debe tener ambos en la MISMA carpeta
    echo.
) else (
    echo ========================================
    echo   ERROR EN LA COMPILACION
    echo ========================================
    echo Revisa los mensajes de error arriba
)

echo.
pause
