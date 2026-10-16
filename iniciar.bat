@echo off
chcp 65001 >nul
title Cuentas a Cobrar
cd /d "%~dp0"

rem La primera vez se crea el entorno de Python e instalan las librerias.
if not exist ".venv\Scripts\python.exe" (
    echo Preparando la aplicacion por primera vez. Puede tardar unos minutos...
    python -m venv .venv
    if errorlevel 1 (
        echo.
        echo No se encontro Python. Instalalo desde https://www.python.org/downloads/
        echo marcando la opcion "Add Python to PATH", y volve a abrir este archivo.
        pause
        exit /b 1
    )
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo No se pudieron instalar las librerias. Revisa la conexion a internet.
        rmdir /s /q .venv
        pause
        exit /b 1
    )
)

rem Evita que Streamlit pida un email la primera vez que se ejecuta.
if not exist "%USERPROFILE%\.streamlit\credentials.toml" (
    mkdir "%USERPROFILE%\.streamlit" 2>nul
    (echo [general]& echo email = "") > "%USERPROFILE%\.streamlit\credentials.toml"
)

call ".venv\Scripts\activate.bat"
echo.
echo La aplicacion se abre en el navegador. Para cerrarla, cerra esta ventana.
streamlit run app.py --browser.gatherUsageStats false
pause
