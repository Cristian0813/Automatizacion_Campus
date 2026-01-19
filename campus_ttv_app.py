import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import time
import os
from selenium import webdriver
import pandas as pd
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException

class CampusTTVApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Campus Talento Tech Valle - Automatización")
        self.root.geometry("1000x800")
        self.root.resizable(True, True)
        
        # Variables
        self.archivo_excel = tk.StringVar()
        self.carpeta_screenshots = tk.StringVar()
        self.procesando = False
        
        # Configurar estilo
        self.configurar_estilos()
        
        # Crear interfaz
        self.crear_interfaz()
        
    def configurar_estilos(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        # Colores
        COLOR_PRIMARIO = "#4A90E2"
        COLOR_SECUNDARIO = "#50C878"
        COLOR_FONDO = "#F5F7FA"
        COLOR_TEXTO = "#2C3E50"
        
        self.root.configure(bg=COLOR_FONDO)
        
        # Estilo para botones
        style.configure("Primary.TButton",
                       font=("Segoe UI", 10, "bold"),
                       background=COLOR_PRIMARIO,
                       foreground="white",
                       padding=10)
        
        style.configure("Success.TButton",
                       font=("Segoe UI", 10, "bold"),
                       background=COLOR_SECUNDARIO,
                       foreground="white",
                       padding=15)
        
    def crear_interfaz(self):
        # Header
        header_frame = tk.Frame(self.root, bg="#4A90E2", height=80)
        header_frame.pack(fill="x")
        
        tk.Label(header_frame,
                text="🎓 Campus Talento Tech Valle",
                font=("Segoe UI", 20, "bold"),
                bg="#4A90E2",
                fg="white").pack(pady=20)
        
        # Contenedor principal
        main_frame = tk.Frame(self.root, bg="#F5F7FA", padx=40, pady=30)
        main_frame.pack(fill="both", expand=True)
        
        # Sección: Archivo Excel
        excel_frame = tk.LabelFrame(main_frame,
                                   text="📊 Archivo Excel con Usuarios",
                                   font=("Segoe UI", 11, "bold"),
                                   bg="#F5F7FA",
                                   fg="#2C3E50",
                                   padx=20,
                                   pady=15)
        excel_frame.pack(fill="x", pady=(0, 20))
        
        tk.Entry(excel_frame,
                textvariable=self.archivo_excel,
                font=("Segoe UI", 10),
                width=60,
                state="readonly").pack(side="left", padx=(0, 10))
        
        ttk.Button(excel_frame,
                  text="Seleccionar Excel",
                  command=self.seleccionar_excel,
                  style="Primary.TButton").pack(side="left")
        
        # Sección: Carpeta Screenshots
        carpeta_frame = tk.LabelFrame(main_frame,
                                     text="📁 Carpeta para Screenshots",
                                     font=("Segoe UI", 11, "bold"),
                                     bg="#F5F7FA",
                                     fg="#2C3E50",
                                     padx=20,
                                     pady=15)
        carpeta_frame.pack(fill="x", pady=(0, 20))
        
        tk.Entry(carpeta_frame,
                textvariable=self.carpeta_screenshots,
                font=("Segoe UI", 10),
                width=60,
                state="readonly").pack(side="left", padx=(0, 10))
        
        ttk.Button(carpeta_frame,
                  text="Seleccionar Carpeta",
                  command=self.seleccionar_carpeta,
                  style="Primary.TButton").pack(side="left")
        
        # Botón Iniciar
        self.btn_iniciar = ttk.Button(main_frame,
                                     text="▶ Iniciar Procesamiento",
                                     command=self.iniciar_procesamiento,
                                     style="Success.TButton")
        self.btn_iniciar.pack(pady=20)
        
        # Barra de progreso
        self.progress = ttk.Progressbar(main_frame,
                                       mode='indeterminate',
                                       length=600)
        self.progress.pack(pady=(0, 10))
        
        # Área de log
        log_frame = tk.LabelFrame(main_frame,
                                 text="📋 Registro de Actividad",
                                 font=("Segoe UI", 11, "bold"),
                                 bg="#F5F7FA",
                                 fg="#2C3E50",
                                 padx=10,
                                 pady=10)
        log_frame.pack(fill="both", expand=True)
        
        # Scrollbar para el log
        scrollbar = tk.Scrollbar(log_frame)
        scrollbar.pack(side="right", fill="y")
        
        self.log_text = tk.Text(log_frame,
                               height=20,
                               font=("Consolas", 9),
                               bg="#FFFFFF",
                               fg="#2C3E50",
                               yscrollcommand=scrollbar.set,
                               wrap="word")
        self.log_text.pack(fill="both", expand=True)
        scrollbar.config(command=self.log_text.yview)
        
        # Footer
        footer = tk.Label(self.root,
                         text="Desarrollado para Talento Tech Valle 💻",
                         font=("Segoe UI", 9),
                         bg="#E8EBF0",
                         fg="#7F8C8D",
                         pady=10)
        footer.pack(side="bottom", fill="x")
        
    def seleccionar_excel(self):
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        if archivo:
            self.archivo_excel.set(archivo)
            self.log(f"✅ Excel seleccionado: {os.path.basename(archivo)}")
    
    def seleccionar_carpeta(self):
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta para screenshots")
        if carpeta:
            self.carpeta_screenshots.set(carpeta)
            self.log(f"✅ Carpeta seleccionada: {carpeta}")
    
    def log(self, mensaje):
        self.log_text.insert("end", f"{mensaje}\n")
        self.log_text.see("end")
        self.root.update()
    
    def iniciar_procesamiento(self):
        # Validaciones
        if not self.archivo_excel.get():
            messagebox.showerror("Error", "Por favor selecciona un archivo Excel")
            return
        
        if not self.carpeta_screenshots.get():
            messagebox.showerror("Error", "Por favor selecciona una carpeta para los screenshots")
            return
        
        if self.procesando:
            messagebox.showwarning("Advertencia", "Ya hay un procesamiento en curso")
            return
        
        # Iniciar procesamiento en hilo separado
        self.procesando = True
        self.btn_iniciar.config(state="disabled")
        self.progress.start()
        
        thread = threading.Thread(target=self.procesar_usuarios, daemon=True)
        thread.start()
    
    def procesar_usuarios(self):
        try:
            self.log("\n" + "="*80)
            self.log("🚀 INICIANDO PROCESAMIENTO")
            self.log("="*80)
            
            # Leer Excel
            self.log(f"\n📂 Leyendo archivo: {os.path.basename(self.archivo_excel.get())}")
            df = pd.read_excel(self.archivo_excel.get())
            
            if "CORREO" not in df.columns or "DOCUMENTO" not in df.columns:
                self.log("❌ ERROR: El Excel debe tener las columnas CORREO y DOCUMENTO")
                messagebox.showerror("Error", "El Excel debe tener las columnas CORREO y DOCUMENTO")
                return
            
            df = df.dropna(subset=["CORREO", "DOCUMENTO"])
            total_usuarios = len(df)
            self.log(f"✅ Usuarios encontrados: {total_usuarios}")
            
            # Crear carpeta si no existe
            if not os.path.exists(self.carpeta_screenshots.get()):
                os.makedirs(self.carpeta_screenshots.get())
                self.log(f"📁 Carpeta creada: {self.carpeta_screenshots.get()}")
            
            # Constantes
            URL_LOGIN = "https://campus.talentotechvalle.co/auth/login"
            BOOTCAMPS = [
                "Análisis de Datos",
                "Inteligencia Artificial",
                "Programación",
                "Ciberseguridad",
                "Arquitectura en la Nube",
                "Blockchain"
            ]
            
            XPATH_BOOTCAMPS_H2 = (
                "//div[contains(@class,'card')]//h2["
                + " or ".join([f"contains(normalize-space(.), '{b}')" for b in BOOTCAMPS])
                + "]"
            )
            
            resultados = []
            
            # Procesar cada usuario
            for index, row in df.iterrows():
                EMAIL = row["CORREO"]
                PASSWORD = str(row["DOCUMENTO"])
                
                self.log(f"\n{'='*80}")
                self.log(f"🔄 PROCESANDO USUARIO {index + 1}/{total_usuarios}")
                self.log(f"📧 Email: {EMAIL}")
                self.log(f"🆔 Documento: {PASSWORD}")
                self.log(f"{'='*80}")
                
                resultado = {
                    "email": EMAIL,
                    "documento": PASSWORD,
                    "estado": "pendiente",
                    "mensaje": ""
                }
                
                options = Options()
                options.add_argument("--start-maximized")
                options.add_argument("--disable-blink-features=AutomationControlled")
                
                driver = webdriver.Chrome(
                    service=Service(ChromeDriverManager().install()),
                    options=options
                )
                
                wait = WebDriverWait(driver, 60)
                
                try:
                    self.log("🔎 Abriendo página de login...")
                    driver.get(URL_LOGIN)
                    
                    # LOGIN
                    self.log("📝 Ingresando credenciales...")
                    wait.until(EC.presence_of_element_located((By.NAME, "email"))).send_keys(EMAIL)
                    driver.find_element(By.NAME, "password").send_keys(PASSWORD)
                    
                    btn_login = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Ingresar']"))
                    )
                    driver.execute_script("arguments[0].click();", btn_login)
                    
                    self.log("⏳ Esperando autenticación...")
                    wait.until(lambda d: "auth/login" not in d.current_url)
                    
                    # Validación
                    html_lang = driver.execute_script("return document.documentElement.lang")
                    if html_lang != "es":
                        raise Exception("HTML base no cargado")
                    
                    self.log("✅ Login exitoso")
                    
                    # Navegación al bootcamp
                    self.log("🎯 Buscando bootcamp...")
                    wait.until(EC.presence_of_element_located((By.XPATH, XPATH_BOOTCAMPS_H2)))
                    
                    try:
                        btn_ingresar = wait.until(
                            EC.element_to_be_clickable((
                                By.XPATH,
                                (
                                    "//div[contains(@class,'card')]"
                                    "//h2["
                                    + " or ".join([f"contains(normalize-space(.), '{b}')" for b in BOOTCAMPS])
                                    + "]"
                                    "/ancestor::div[contains(@class,'card')]"
                                    "//button[not(@disabled) and contains(., 'Ingresar')]"
                                )
                            ))
                        )
                    except TimeoutException:
                        btn_ingresar = wait.until(
                            EC.element_to_be_clickable((
                                By.XPATH,
                                XPATH_BOOTCAMPS_H2 +
                                "/ancestor::div[contains(@class,'card')]"
                                "//button[not(@disabled) and contains(., 'Ingresar')]"
                            ))
                        )
                    
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_ingresar)
                    driver.execute_script("arguments[0].click();", btn_ingresar)
                    
                    self.log("⏳ Cargando bootcamp...")
                    wait.until(lambda d: "bootcamp" in d.current_url or "inicio" not in d.current_url)
                    
                    # Scroll
                    self.log("📜 Scrolleando página...")
                    last_height = driver.execute_script("return document.body.scrollHeight")
                    scroll_count = 0
                    
                    while scroll_count < 10:
                        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                        time.sleep(1)
                        new_height = driver.execute_script("return document.body.scrollHeight")
                        
                        if new_height == last_height:
                            break
                        
                        last_height = new_height
                        scroll_count += 1
                    
                    self.log("✅ Proceso completado exitosamente")
                    resultado["estado"] = "exitoso"
                    resultado["mensaje"] = "Completado"
                    
                except Exception as e:
                    self.log(f"❌ ERROR: {str(e)}")
                    resultado["estado"] = "error"
                    resultado["mensaje"] = str(e)
                
                finally:
                    # Screenshot
                    nombre_archivo = f"{PASSWORD}.png"
                    ruta_completa = os.path.join(self.carpeta_screenshots.get(), nombre_archivo)
                    
                    try:
                        time.sleep(1)
                        driver.save_screenshot(ruta_completa)
                        self.log(f"📸 Screenshot guardado: {nombre_archivo}")
                    except:
                        self.log("⚠️ No se pudo guardar screenshot")
                    
                    driver.quit()
                    resultados.append(resultado)
                    
                    # Pausa entre usuarios
                    if index < total_usuarios - 1:
                        time.sleep(2)
            
            # Resumen
            exitosos = sum(1 for r in resultados if r["estado"] == "exitoso")
            self.log(f"\n{'='*80}")
            self.log("📊 RESUMEN FINAL")
            self.log(f"{'='*80}")
            self.log(f"✅ Exitosos: {exitosos}/{total_usuarios}")
            self.log(f"❌ Fallidos: {total_usuarios - exitosos}/{total_usuarios}")
            
            # Guardar resultados
            df_resultados = pd.DataFrame(resultados)
            archivo_resultados = os.path.join(self.carpeta_screenshots.get(), "resultados_procesamiento.xlsx")
            df_resultados.to_excel(archivo_resultados, index=False)
            self.log(f"\n💾 Resultados guardados: {archivo_resultados}")
            
            messagebox.showinfo("Completado", f"Procesamiento finalizado\n✅ Exitosos: {exitosos}/{total_usuarios}")
            
        except Exception as e:
            self.log(f"\n❌ ERROR CRÍTICO: {str(e)}")
            messagebox.showerror("Error", f"Error crítico: {str(e)}")
        
        finally:
            self.procesando = False
            self.btn_iniciar.config(state="normal")
            self.progress.stop()
            self.log("\n🧪 PROCESAMIENTO FINALIZADO\n")

if __name__ == "__main__":
    root = tk.Tk()
    app = CampusTTVApp(root)
    root.mainloop()