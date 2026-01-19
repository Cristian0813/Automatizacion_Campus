@echo off
set "ROOT_PATH=C:\Users\Talento Tech 19\Desktop\Codigos Python\capturapantalla_campusttv"
cd /d "%ROOT_PATH%"

echo Iniciando Campus Talento Tech Valle...
echo Cargando entorno virtual...

:: Activar el entorno virtual
call "env\Scripts\activate"

:: Ejecutar la aplicacion
python campus_ttv_app.py

if %errorlevel% neq 0 (
    echo.
    echo Ocurrio un error al iniciar la aplicacion.
    pause
)