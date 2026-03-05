# Automatización del Campus

Aplicación de escritorio desarrollada en **Python**, que utiliza **Tkinter** para la interfaz gráfica y **Selenium** para la automatización web.
Su objetivo es automatizar el ingreso al **Campus Talento Tech Valle**, recorrer los bootcamps asignados a cada usuario y generar evidencias mediante capturas de pantalla.
Los usuarios se cargan desde un archivo **Excel**, lo que permite procesar múltiples accesos de forma automática.

---

## Funcionalidades principales

* Interfaz gráfica desarrollada con Tkinter
* Lectura automática de usuarios desde Excel (`CORREO` y `DOCUMENTO`)
* Inicio de sesión automatizado en el campus
* Acceso automático al bootcamp correspondiente
* Scroll automático de la página para cargar contenido
* Captura de pantalla por cada usuario procesado
* Registro detallado del proceso de ejecución
* Generación de archivo de resultados (`resultados_procesamiento.xlsx`)
* Ejecución en segundo plano mediante multihilo

---

## Estructura del proyecto

```
capturapantalla_campusttv/
│
├── campus_ttv_app.py
├── README.md
├── iniciar_app.bat
├── env/
├── screenshots/
└── resultados_procesamiento.xlsx
```

---

## Requisitos del archivo Excel

El archivo Excel debe contener las siguientes columnas:

| Columna   | Descripción                 |
| --------- | --------------------------- |
| CORREO    | Usuario de acceso al campus |
| DOCUMENTO | Contraseña de acceso        |

Los nombres de las columnas deben coincidir exactamente para que el sistema pueda leerlos correctamente.

---

## Requisitos del sistema

* Windows 10 o superior
* Conexión a Internet
* Google Chrome instalado y actualizado
* Python 3.9 o superior

---

## Instalación de Python

### 1. Descargar Python

1. Ingresa a:
   [https://www.python.org/downloads/](https://www.python.org/downloads/)

2. Descarga **Python 3.9 o superior**.

3. Ejecuta el instalador.

### 2. Configuración durante la instalación

Marca la opción:

```
Add Python to PATH
```

Luego haz clic en **Install Now**.

Esto permite ejecutar Python desde la consola del sistema.

### 3. Verificar la instalación

Abre **Símbolo del sistema (CMD)** y ejecuta:

```bash
python --version
```

Si aparece una versión de Python, la instalación fue exitosa.

---

## Creación del entorno virtual

1. Abre **CMD** o **PowerShell**.

2. Ubícate en la carpeta del proyecto:

```bash
cd C:\Ruta\Donde\Esta\capturapantalla_campusttv
```

3. Crea el entorno virtual:

```bash
python -m venv env
```

4. Activa el entorno virtual:

```bash
env\Scripts\activate
```

Si la activación fue correcta, aparecerá `(env)` al inicio de la línea de comandos.

---

## Instalación de dependencias

Con el entorno virtual activo, instala las librerías necesarias:

```bash
pip install selenium webdriver-manager
pip install pandas openpyxl
```

Tkinter viene incluido con Python, por lo que no requiere instalación adicional.

---

## Ejecución de la aplicación

### Uso mediante archivo `.bat` (recomendado)

Ejemplo de archivo `iniciar_app.bat`:

```bat
@echo off
set "ROOT_PATH=C:\Users\Cambia a la carpeta donde esta el proyecto"
cd /d "%ROOT_PATH%"

echo Iniciando Campus Talento Tech Valle...
echo Cargando entorno virtual...

call "env\Scripts\activate"
python Mia_app.py

if %errorlevel% neq 0 (
    echo.
    echo Ocurrio un error al iniciar la aplicacion.
    pause
)
```

Ajusta:

* `ROOT_PATH` con la ruta correcta del proyecto.
* `Mia_app.py` con el nombre del archivo principal de tu aplicación.

---

## Resultados generados

Durante la ejecución del sistema se generan los siguientes archivos:

* Capturas de pantalla almacenadas en la carpeta `screenshots/`
* Archivo `resultados_procesamiento.xlsx` con el estado del procesamiento de cada usuario

---

## Advertencias

* Este sistema automatiza accesos reales al campus.
* Debe utilizarse únicamente dentro del entorno de **Talento Tech Valle**.
* No se recomienda distribuir públicamente sin autorización institucional.
* Utilizar únicamente con permisos de uso adecuados.

---

## Licencia

Este proyecto está bajo la licencia **MIT**.

Desarrollado para **Talento Tech Valle**.

Copyright (c) 2026 Cristian Arias.
