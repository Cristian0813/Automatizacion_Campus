import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import time
import os
from datetime import datetime, timedelta
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
        
        # Variables para el cronómetro
        self.tiempo_inicio = None
        self.cronometro_activo = False
        
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
        
        # Frame para cronómetro
        cronometro_frame = tk.Frame(main_frame, bg="#F5F7FA")
        cronometro_frame.pack(pady=(0, 10))
        
        tk.Label(cronometro_frame,
                text="⏱️ Tiempo transcurrido:",
                font=("Segoe UI", 11, "bold"),
                bg="#F5F7FA",
                fg="#2C3E50").pack(side="left", padx=(0, 10))
        
        self.label_cronometro = tk.Label(cronometro_frame,
                                         text="00:00:00",
                                         font=("Consolas", 16, "bold"),
                                         bg="#F5F7FA",
                                         fg="#E74C3C")
        self.label_cronometro.pack(side="left")
        
        # Barra de progreso
        self.progress = ttk.Progressbar(main_frame,
                                       mode='indeterminate',
                                       length=600)
        self.progress.pack(pady=(10, 10))
        
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
    
    def actualizar_cronometro(self):
        """Actualiza el cronómetro cada segundo"""
        if self.cronometro_activo and self.tiempo_inicio:
            tiempo_transcurrido = datetime.now() - self.tiempo_inicio
            horas = int(tiempo_transcurrido.total_seconds() // 3600)
            minutos = int((tiempo_transcurrido.total_seconds() % 3600) // 60)
            segundos = int(tiempo_transcurrido.total_seconds() % 60)
            
            tiempo_formateado = f"{horas:02d}:{minutos:02d}:{segundos:02d}"
            self.label_cronometro.config(text=tiempo_formateado)
            
            # Llamar de nuevo después de 1 segundo
            self.root.after(1000, self.actualizar_cronometro)
    
    def iniciar_cronometro(self):
        """Inicia el cronómetro"""
        self.tiempo_inicio = datetime.now()
        self.cronometro_activo = True
        self.label_cronometro.config(fg="#E74C3C")
        self.actualizar_cronometro()
    
    def detener_cronometro(self):
        """Detiene el cronómetro"""
        self.cronometro_activo = False
        self.label_cronometro.config(fg="#27AE60")
        
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
        
        # Iniciar cronómetro
        self.iniciar_cronometro()
        
        thread = threading.Thread(target=self.procesar_usuarios, daemon=True)
        thread.start()
    
    def cerrar_sesion(self, driver, wait):
        """Cierra sesión y vuelve al login"""
        try:
            self.log("🚪 Cerrando sesión...")
            
            # Estrategia 1: Buscar el botón "Cerrar Sesión" en el drawer/sidebar
            try:
                btn_cerrar_sesion = wait.until(
                    EC.element_to_be_clickable((
                        By.XPATH,
                        "//button[contains(@class, 'text-error') and .//span[contains(., 'Cerrar Sesión')]]"
                    ))
                )
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_cerrar_sesion)
                # time.sleep(0.5)
                driver.execute_script("arguments[0].click();", btn_cerrar_sesion)
                self.log("   ✅ Click en 'Cerrar Sesión'")
            except TimeoutException:
                # Estrategia 2: Buscar por el SVG y texto
                btn_cerrar_sesion = wait.until(
                    EC.element_to_be_clickable((
                        By.XPATH,
                        "//button[contains(@class, 'border-error')]//span[contains(., 'Cerrar Sesión')]"
                    ))
                )
                driver.execute_script("arguments[0].click();", btn_cerrar_sesion)
                self.log("   ✅ Click en 'Cerrar Sesión' (estrategia 2)")
            
            # Esperar a que redirija al login
            wait.until(lambda d: "auth/login" in d.current_url)
            self.log("   ✅ Sesión cerrada exitosamente")
            # time.sleep(1)
            return True
            
        except Exception as e:
            self.log(f"   ⚠️ Error cerrando sesión: {str(e)}")
            self.log("   🔄 Navegando manualmente al login...")
            driver.get("https://campus.talentotechvalle.co/auth/login")
            # time.sleep(2)
            return False
    
    def procesar_usuarios(self):
        driver = None
        
        try:
            self.log("\n" + "="*80)
            self.log("🚀 INICIANDO PROCESAMIENTO")
            self.log(f"⏰ Hora de inicio: {self.tiempo_inicio.strftime('%H:%M:%S')}")
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
                "Análisis y Visualización de Datos",
                "Inteligencia Artificial",
                "Programación",
                "Desarrollo Web Full Stack",                
                "Ciberseguridad",
                "Arquitectura en la Nube",
                "Blockchain"
            ]
            
            XPATH_BOOTCAMPS_H2 = (
                "//div[contains(@class,'card')]//h2["
                + " or ".join([f"contains(normalize-space(.), '{b}')" for b in BOOTCAMPS])
                + "]"
            )
            
            # ========== INICIALIZAR CHROME UNA SOLA VEZ ==========
            self.log("\n🌐 Iniciando Chrome (se reutilizará para todos los usuarios)...")
            options = Options()
            options.add_argument("--start-maximized")
            options.add_argument("--disable-blink-features=AutomationControlled")
            
            driver = webdriver.Chrome(
                service=Service(ChromeDriverManager().install()),
                options=options
            )
            
            wait = WebDriverWait(driver, 60)
            self.log("✅ Chrome iniciado correctamente")
            
            resultados = []
            
            # ========== PROCESAR CADA USUARIO ==========
            for index, row in df.iterrows():
                PASSWORD = str(row["DOCUMENTO"])
                NOMBRE = str(row.get("NOMBRE COMPLETO", "")).strip()
                EMAIL = row["CORREO"]
                TELEFONO = str(row.get("TELÉFONO", "")).strip()
                
                self.log(f"\n{'='*80}")
                self.log(f"🔄 PROCESANDO USUARIO {index + 1}/{total_usuarios}")
                self.log(f"📧 Email: {EMAIL}")
                self.log(f"🆔 Documento: {PASSWORD}")
                self.log(f"{'='*80}")
                
                resultado = {
                    "DOCUMENTO": PASSWORD,
                    "NOMBRE COMPLETO": NOMBRE,
                    "CORREO": EMAIL,
                    "TELÉFONO": TELEFONO,
                    "ESTADO": "Pendiente"
                }
                
                try:
                    # Si no es el primer usuario, ya estamos en login por el logout anterior
                    if index == 0:
                        self.log("🔎 Navegando a login...")
                        driver.get(URL_LOGIN)
                        # time.sleep(2)
                    
                    # LOGIN
                    self.log("📝 Ingresando credenciales...")
                    
                    # Limpiar campos antes de escribir
                    email_input = wait.until(EC.presence_of_element_located((By.NAME, "email")))
                    email_input.clear()
                    email_input.send_keys(EMAIL)
                    
                    password_input = driver.find_element(By.NAME, "password")
                    password_input.clear()
                    password_input.send_keys(PASSWORD)
                    
                    btn_login = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Ingresar']"))
                    )
                    driver.execute_script("arguments[0].click();", btn_login)
                    
                    self.log("⏳ Esperando autenticación...")
                    wait.until(lambda d: "auth/login" not in d.current_url)
                    # time.sleep(2)
                    
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
                    # time.sleep(0.5)
                    driver.execute_script("arguments[0].click();", btn_ingresar)
                    
                    self.log("⏳ Cargando bootcamp...")
                    wait.until(lambda d: "bootcamp" in d.current_url or "inicio" not in d.current_url)
                    # time.sleep(2)
                    
                    # Scroll hasta el final
                    self.log("📜 Scrolleando hasta el final...")
                    last_height = driver.execute_script("return document.body.scrollHeight")
                    scroll_count = 0
                    scroll_increment = 800
                    
                    while True:
                        driver.execute_script(f"window.scrollBy(0, {scroll_increment});")
                        time.sleep(1)
                        
                        new_height = driver.execute_script("return document.body.scrollHeight")
                        current_position = driver.execute_script("return window.pageYOffset + window.innerHeight;")
                        
                        scroll_count += 1
                        
                        if current_position >= new_height - 10:
                            break
                        
                        if new_height == last_height:
                            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                            time.sleep(1)
                            final_height = driver.execute_script("return document.body.scrollHeight")
                            if final_height == new_height:
                                break
                        
                        last_height = new_height
                        
                        if scroll_count > 20:
                            break
                    
                    self.log(f"   ✅ Scroll completado ({scroll_count} scrolls)")
                    
                    # Screenshot
                    nombre_archivo = f"{PASSWORD}.png"
                    ruta_completa = os.path.join(self.carpeta_screenshots.get(), nombre_archivo)
                    
                    # time.sleep(1)
                    driver.save_screenshot(ruta_completa)
                    self.log(f"📸 Screenshot guardado: {nombre_archivo}")
                    
                    self.log("✅ Proceso completado exitosamente")
                    resultado["ESTADO"] = "Exitoso"
                    
                except Exception as e:
                    self.log(f"❌ ERROR: {str(e)}")
                    resultado["ESTADO"] = f"Error: {str(e)[:100]}"
                    
                    # Screenshot de error
                    try:
                        error_screenshot = os.path.join(
                            self.carpeta_screenshots.get(), 
                            f"{PASSWORD}_ERROR.png"
                        )
                        driver.save_screenshot(error_screenshot)
                        self.log(f"📸 Screenshot de error guardado")
                    except:
                        pass
                
                finally:
                    resultados.append(resultado)
                    
                    # CERRAR SESIÓN (excepto en el último usuario)
                    if index < total_usuarios - 1:
                        self.cerrar_sesion(driver, wait)
                        # time.sleep(2)
            
            # Calcular tiempo total
            tiempo_fin = datetime.now()
            tiempo_total = tiempo_fin - self.tiempo_inicio
            horas = int(tiempo_total.total_seconds() // 3600)
            minutos = int((tiempo_total.total_seconds() % 3600) // 60)
            segundos = int(tiempo_total.total_seconds() % 60)
            
            # Resumen
            exitosos = sum(1 for r in resultados if r["ESTADO"] == "Exitoso")
            self.log(f"\n{'='*80}")
            self.log("📊 RESUMEN FINAL")
            self.log(f"{'='*80}")
            self.log(f"⏰ Hora de inicio: {self.tiempo_inicio.strftime('%H:%M:%S')}")
            self.log(f"⏰ Hora de fin: {tiempo_fin.strftime('%H:%M:%S')}")
            self.log(f"⏱️ Tiempo total: {horas:02d}:{minutos:02d}:{segundos:02d}")
            self.log(f"✅ Exitosos: {exitosos}/{total_usuarios}")
            self.log(f"❌ Fallidos: {total_usuarios - exitosos}/{total_usuarios}")
            
            # Guardar resultados
            df_resultados = pd.DataFrame(resultados)
            archivo_resultados = os.path.join(
                self.carpeta_screenshots.get(), 
                f"resultados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )
            df_resultados.to_excel(archivo_resultados, index=False)
            self.log(f"\n💾 Resultados guardados: {archivo_resultados}")
            
            messagebox.showinfo("Completado", 
                              f"Procesamiento finalizado\n"
                              f"⏱️ Tiempo: {horas:02d}:{minutos:02d}:{segundos:02d}\n"
                              f"✅ Exitosos: {exitosos}/{total_usuarios}")
            
        except Exception as e:
            self.log(f"\n❌ ERROR CRÍTICO: {str(e)}")
            messagebox.showerror("Error", f"Error crítico: {str(e)}")
        
        finally:
            # CERRAR CHROME SOLO AL FINAL
            if driver:
                self.log("\n🔒 Cerrando navegador...")
                driver.quit()
                self.log("✅ Navegador cerrado")
            
            self.procesando = False
            self.btn_iniciar.config(state="normal")
            self.progress.stop()
            self.detener_cronometro()
            self.log("\n🏁 PROCESAMIENTO FINALIZADO\n")

if __name__ == "__main__":
    root = tk.Tk()
    app = CampusTTVApp(root)
    root.mainloop()