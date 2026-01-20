# 🎓 Campus Talento Tech Valle – Automatización

Aplicación de escritorio desarrollada en **Python**, utilizando **Tkinter** para la interfaz gráfica y **Selenium** para la automatización web.
Su objetivo es **automatizar el ingreso al Campus Talento Tech Valle**, recorrer los bootcamps asignados a cada usuario y generar **evidencias mediante capturas de pantalla**, tomando los usuarios desde un archivo **Excel**.

---

## 🚀 Funcionalidades principales

* Interfaz gráfica amigable (Tkinter)
* Lectura automática de usuarios desde Excel (`CORREO` y `DOCUMENTO`)
* Inicio de sesión automatizado en el campus
* Acceso al bootcamp correspondiente
* Scroll automático de la página
* Captura de pantalla por usuario
* Registro detallado del proceso
* Generación de archivo de resultados (`resultados_procesamiento.xlsx`)
* Ejecución segura en segundo plano (multihilo)

---

## 📁 Estructura del proyecto

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

## 📊 Requisitos del archivo Excel

| Columna   | Descripción                 |
| --------- | --------------------------- |
| CORREO    | Usuario de acceso al campus |
| DOCUMENTO | Contraseña                  |

📌 Los nombres de las columnas deben coincidir exactamente.

---

## 🧰 Requisitos del sistema

* Windows 10 o superior
* Conexión a Internet
* Google Chrome instalado y actualizado

---

## 🐍 Instalación de Python (PASO A PASO)

### 1️⃣ Descargar Python

1. Ingresa a: [https://www.python.org/downloads/](https://www.python.org/downloads/)
2. Descarga **Python 3.9 o superior**
3. Ejecuta el instalador

### 2️⃣ Configuración IMPORTANTE durante la instalación

✔ Marca la opción **Add Python to PATH**
✔ Luego haz clic en **Install Now**

📌 Esto es obligatorio para que Python funcione correctamente desde la consola.

### 3️⃣ Verificar instalación

1. Abre **Símbolo del sistema (CMD)**
2. Ejecuta:

```bash
python --version
```

Si ves algo como `Python 3.x.x`, la instalación fue exitosa.

---

## 📦 Creación del entorno virtual

1️⃣ Abre **CMD** o **PowerShell**
2️⃣ Ubícate en la carpeta del proyecto:

```bash
cd C:\Ruta\Donde\Esta\capturapantalla_campusttv
```

3️⃣ Crea el entorno virtual:

```bash
python -m venv env
```

4️⃣ Activa el entorno virtual:

```bash
env\Scripts\activate
```

Si todo está correcto, verás `(env)` al inicio de la línea.

---

## 📚 Instalación de librerías necesarias

Con el entorno virtual activo, ejecuta:

```bash
# Instalación de Selenium y el manejador de drivers
pip install selenium webdriver-manager

# Instalación de Pandas (para manejo de archivos Excel/CSV)
pip install pandas openpyxl
```

📌 Tkinter viene incluido con Python, no requiere instalación adicional.

---

## ▶️ Ejecución de la aplicación

### Usando archivo `.bat` (recomendado)

Ejemplo de `iniciar_app.bat`:

```bat
@echo off
set "ROOT_PATH=C:\Users\Cambia a la carpeta donde esta el proyecto"
cd /d "%ROOT_PATH%"

echo Iniciando Campus Talento Tech Valle...
echo Cargando entorno virtual...

call "env\Scripts\activate"
python campus_ttv_app.py

if %errorlevel% neq 0 (
    echo.
    echo Ocurrio un error al iniciar la aplicacion.
    pause
)
```

📌 Ajusta la ruta `ROOT_PATH` según tu equipo.

---

## 📂 Resultados generados

* Capturas almacenadas en `screenshots/`
* Archivo `resultados_procesamiento.xlsx` con el estado del proceso

---

## ⚠️ Advertencias importantes

* Automatiza accesos reales
* Uso exclusivo Talento Tech Valle
* No distribuir públicamente
* Usar solo con autorización institucional

---

## 👨‍💻 Autor

**CRISTIAN J ARIAS O**
Automatización académica y control de evidencias

---

Si quieres, puedo ayudarte a:

✔ Crear una versión **para usuarios finales (solo doble clic)**
✔ Generar un **instalador o .exe**
✔ Añadir **capturas al README**
✔ Documentar errores comunes y soluciones

Solo dime 👍
