import os
import re
import shutil
import fitz
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from datetime import datetime
import win32clipboard
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import TkinterDnD
from PIL import Image, ImageTk
import json
import time

"""
Este programa observa una carpeta específica en busca de convenios en archivos PDF.
Cuando se detecta un nuevo archivo PDF, extrae información relevante del mismo,crea 
una carpeta con el nombre del estudiante y el tipo de convenio y mueve el PDFa esa
carpeta.
Además, copia los datos relevantes al portapapeles en un formato que puede ser pegado
directamente en el Excel de seguimiento.

Esta versión en conreto funciona con una interfaz gráfica de usuario (GUI) que permite al usuario
seleccionar las carpetas de origen y destino para la monitorización y el guardado de los archivos.

 ______________________________________________________________________________________
|                                                                                      |
| !!! Se debe modificar la ruta base y la carpeta monitoreada según sea necesario. !!! |
|______________________________________________________________________________________|

La ruta base corresponde a la carpeta donde se crearán las subcarpetas para cada
estudiante.
El programa usa la biblioteca watchdog para monitorear la carpeta y Fitz para extraer
texto de los archivos PDF.
"""

# Carpeta de descargas del usuario por defecto
descargas_path = os.path.join(os.environ["USERPROFILE"], "Downloads")

# Ruta base donde se crearán las carpetas
RUTA_BASE = descargas_path
MONITORED_FOLDER = descargas_path

CONFIG_FILE = "creacarp_config.json"

def cargar_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
                return config.get("monitorizacion", descargas_path), config.get("salida", descargas_path)
        except Exception:
            pass
    return descargas_path, descargas_path

def guardar_config(monitorizacion, salida):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"monitorizacion": monitorizacion, "salida": salida}, f)
    except Exception:
        pass

# --- INICIO INTERFAZ GRÁFICA ---
class App:
    def __init__(self, root):
        self.root = root
        self.root.iconbitmap("./assets/upv.ico")
        self.root.title("CreaCarpListener UPE INF")
        self.root.geometry("650x500")

        self.img = Image.open("./assets/titulo_etsinf.png")
        self.img = self.img.resize((600, 100), Image.LANCZOS)  # Ajustar el tamaño de la imagen
        self.photo = ImageTk.PhotoImage(self.img)
        self.label_img = tk.Label(root, image=self.photo)
        self.label_img.pack(pady=10)

        # Cargar rutas desde config
        self.monitored_folder, self.ruta_base = cargar_config()

        tk.Label(root, text="Ruta de monitorización (origen):", font=("Arial", 11)).pack(pady=8)
        self.entry_origen = tk.Entry(root, width=60)
        self.entry_origen.pack(pady=2)
        self.entry_origen.insert(0, self.monitored_folder)
        tk.Button(root, text="Seleccionar carpeta", command=self.seleccionar_origen).pack(pady=2)

        tk.Label(root, text="Ruta de destino (carpetas creadas):", font=("Arial", 11)).pack(pady=8)
        self.entry_destino = tk.Entry(root, width=60)
        self.entry_destino.pack(pady=2)
        self.entry_destino.insert(0, self.ruta_base)
        tk.Button(root, text="Seleccionar carpeta", command=self.seleccionar_destino).pack(pady=2)

        self.button_iniciar = tk.Button(root, text="Iniciar monitorización", command=self.iniciar_observador, bg="green", fg="white", font=("Arial", 12, "bold"))
        self.button_iniciar.pack(pady=15)
        self.button_detener = tk.Button(root, text="Detener monitorización", command=self.detener_observador, state=tk.DISABLED, bg="red", fg="white", font=("Arial", 12, "bold"))
        self.button_detener.pack(pady=5)

        self.observer = None

    def seleccionar_origen(self):
        carpeta = filedialog.askdirectory()
        if carpeta:
            self.monitored_folder = carpeta
            self.entry_origen.delete(0, tk.END)
            self.entry_origen.insert(0, carpeta)
            guardar_config(self.monitored_folder, self.ruta_base)  # <-- monitorizacion, salida

    def seleccionar_destino(self):
        carpeta = filedialog.askdirectory()
        if carpeta:
            self.ruta_base = carpeta
            self.entry_destino.delete(0, tk.END)
            self.entry_destino.insert(0, carpeta)
            guardar_config(self.monitored_folder, self.ruta_base)  # <-- monitorizacion, salida

    def iniciar_observador(self):
        global RUTA_BASE, MONITORED_FOLDER
        RUTA_BASE = self.ruta_base
        MONITORED_FOLDER = self.monitored_folder

        guardar_config(self.monitored_folder, self.ruta_base)  # <-- monitorizacion, salida

        if not RUTA_BASE or not MONITORED_FOLDER:
            messagebox.showerror("Error", "Debe seleccionar ambas carpetas.")
            return

        event_handler = PDFHandler()
        self.observer = Observer()
        self.observer.schedule(event_handler, MONITORED_FOLDER, recursive=False)
        self.observer.start()
        messagebox.showinfo("Monitorización", "Monitorización iniciada.")
        self.button_iniciar.config(state=tk.DISABLED)
        self.button_detener.config(state=tk.NORMAL)

    def detener_observador(self):
        if self.observer:
            self.observer.stop()
            self.observer.join()
            messagebox.showinfo("Monitorización", "Monitorización detenida.")
            self.button_iniciar.config(state=tk.NORMAL)
            self.button_detener.config(state=tk.DISABLED)

# --- FIN INTERFAZ GRÁFICA ---

def copiar_a_portapapeles_excel(contenido):
    """
    Copia texto formateado como celdas adyacentes de Excel (Calibri 10)
    params: contenido: lista de cadenas que representan las celdas a copiar
    returns: None
    """
    try:
        # Formato para simular Excel (tab como separador de celdas, \n para filas)
        texto_excel = "\t".join(contenido)
        
        print(texto_excel)  # Para depuración, imprime el contenido que se copiará al portapapeles
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(texto_excel)
        win32clipboard.CloseClipboard()

        # Leer y mostrar el contenido del portapapeles
        win32clipboard.OpenClipboard()
        texto_leido = win32clipboard.GetClipboardData()
        win32clipboard.CloseClipboard()
        print("Contenido actual del portapapeles:")
        print(repr(texto_leido))

    except Exception as e:
        print(f"Error al copiar al portapapeles: {e}")

def extraer_datos(pdf_path):
    """
    Extrae el nombre del estudiante, tipo de convenio y otros datos del PDF,
    basándose en el contenido del texto del PDF, según las líneas esperadas.
    params: pdf_path: ruta del archivo PDF a procesar
    returns: una tupla con los datos extraídos o un mensaje de error
    """
    try:
        with fitz.open(pdf_path) as doc:
            texto = doc[0].get_text("text")  # Extrae texto de la primera página
        lineas = texto.split("\n")
        
        # Buscar nombre del estudiante
        nombre_completo = None
        cp_index = -1
        for i, linea in enumerate(lineas):
            if "Centro Docente:" in linea or "Centre Docent:" in linea:
                centro_docente_index = i
                if i + 1 < len(lineas):
                    nombre_completo = lineas[i + 1].strip()
            if "CP:" in linea:
                cp_index = i

        if not nombre_completo:
            raise ValueError("No se encontró el nombre del estudiante en el PDF.")
        if cp_index == -1:
            raise ValueError("No se encontró el campo 'CP:' en el PDF.")

        # Extraer DNI (línea justo debajo del nombre)
        dni = lineas[lineas.index(nombre_completo) + 1].strip() if nombre_completo in lineas else ""

        # Extraer datos de la empresa
        nombre_empresa = lineas[cp_index + 11].strip() if cp_index + 11 < len(lineas) else ""
        cif_empresa = lineas[cp_index + 3].strip() if cp_index + 3 < len(lineas) else ""
        correo_empresa = lineas[cp_index + 9].strip() if cp_index + 9 < len(lineas) else ""

        # Verificar si la entidad es de la UPV y formatear el nombre
        practica_UPV = False
        if "Q4618002B" in cif_empresa:
            nombre_empresa_sin_upv = nombre_empresa.replace("UPV", "").replace("Universitat Politècnica de València", "").replace("Universitat Politecnica de Valencia", "").replace("Universidad Politécnica de Valencia", "").replace("Universidad Politecnica de Valencia", "").replace("ETSINF", "").replace("-", "")
            nombre_empresa = f"UPV - {nombre_empresa_sin_upv}"
            practica_UPV = True

        # Extraer fechas
        fecha_inicio = ""
        fecha_fin = ""
        index_fecha_fin = -1
        for i, linea in enumerate(lineas):
            if re.match(r'\d{2}/\d{2}/\d{4}', linea):
                fecha_inicio = linea
                if i + 1 < len(lineas) and re.match(r'\d{2}/\d{2}/\d{4}', lineas[i + 1]):
                    fecha_fin = lineas[i + 1]
                    index_fecha_fin = i + 1
                break

        # Extraer titulación
        titulacion = lineas[centro_docente_index + 3] if centro_docente_index + 3 < len(lineas) else ""

        # Separar apellidos y nombre
        partes_nombre = nombre_completo.split()
        if len(partes_nombre) < 3:
            if len(partes_nombre) == 2:
                nombre = partes_nombre[0].capitalize()
                apellido1 = partes_nombre[1].capitalize()
                apellido2 = ""
            else:
                raise ValueError("Nombre del estudiante no reconocido.")
        else:
            palabras_excepcionales = {"de", "la", "del", "los", "las"}
            apellido1 = []
            apellido2 = []
            nombre = []
            apellido1_pointer = False

            i = len(partes_nombre) - 1
            while i >= 0:
                if i == len(partes_nombre) - 1:
                    apellido2.append(partes_nombre[i].capitalize())
                elif partes_nombre[i].lower() in palabras_excepcionales:
                    if apellido1_pointer:
                        apellido1.insert(0, partes_nombre[i].capitalize())
                    else:
                        apellido2.insert(0, partes_nombre[i].capitalize())
                elif not apellido1_pointer:
                    apellido1_pointer = True
                    apellido1.insert(0, partes_nombre[i].capitalize())
                else:
                    nombre.insert(0, partes_nombre[i].capitalize())
                i -= 1

            apellido1 = " ".join(apellido1).strip()
            apellido2 = " ".join(apellido2).strip()
            nombre = " ".join(nombre).strip()

        apellido1 = apellido1.upper()
        apellido2 = apellido2.upper()
        nombre = nombre.title()

        # Buscar tipo de convenio
        tipo_convenio = "Extra"
        if "curriculares" in texto:
            tipo_convenio = "Curr"

        tipo_convenio_excel = "Curr" if tipo_convenio == "Curr" else "Extrac"
        codigo_titulacion = ""
        if "DOBLE GRADO EN ADMINISTRACION Y DIRECCIÓN DE EMPRESAS + INGENIERIA" in titulacion:
            codigo_titulacion = f"DG ADE-GII {tipo_convenio_excel}."
        elif "GRADO EN INGENIERÍA INFORMÁTICA" in titulacion:
            codigo_titulacion = f"GII {tipo_convenio_excel}."
        elif "GRADO EN CIENCIA DE DATOS" in titulacion:
            codigo_titulacion = f"GCD {tipo_convenio_excel}."
        elif "MÁSTER EN INGENIERÍA INFORMÁTICA" in titulacion:
            codigo_titulacion = f"MUIINF {tipo_convenio_excel}."
        elif "MÁSTER EN CIBERSEGURIDAD Y CIBERINTELIGENCIA" in titulacion:
            codigo_titulacion = f"MUCC {tipo_convenio_excel}."
        elif "MÁSTER EN HUMANIDADES DIGITALES" in titulacion:
            codigo_titulacion = f"MUHD {tipo_convenio_excel}."
        elif "GRADO EN INFORMÁTICA INDUSTRIAL Y ROBÓTICA" in titulacion:
            codigo_titulacion = f"GIIR {tipo_convenio_excel}."
        
        

        if tipo_convenio == "Curr":
            num_c = lineas[index_fecha_fin + 1] if index_fecha_fin + 1 < len(lineas) else ""
            numero_creditos = (float(num_c) / 25) if num_c else ""
            numero_creditos = str(numero_creditos).replace('.', ',')
        else:
            numero_creditos = ""

        horas = lineas[index_fecha_fin + 1] if index_fecha_fin + 1 < len(lineas) else ""
        if horas:
            horas = horas.replace('.', ',')

        cantidad_dinero = ""
        match = re.search(r"El estudiante recibirá la cantidad de (\d+(?:\.\d+)?)", texto)
        if match:
            cantidad_dinero = match.group(1).replace('.', ',')
        else:
            cantidad_dinero = "No especificado"

        pract_upv_suffix = " - Práctica UPV" if practica_UPV else ""

        if len(partes_nombre) == 2:
            nombre_carpeta = f"{apellido1.upper()}, {nombre} - Conv {tipo_convenio}{pract_upv_suffix}"
            return (nombre_carpeta, nombre_completo.strip(), tipo_convenio, dni, nombre_empresa,
                    cif_empresa, correo_empresa, fecha_inicio, fecha_fin, codigo_titulacion,
                    numero_creditos, horas, cantidad_dinero, practica_UPV)
        else:
            nombre_carpeta = f"{apellido1.upper()} {apellido2.upper()}, {nombre} - Conv {tipo_convenio}{pract_upv_suffix}"
            return (nombre_carpeta, nombre_completo.strip(), tipo_convenio, dni, nombre_empresa,
                    cif_empresa, correo_empresa, fecha_inicio, fecha_fin, codigo_titulacion,
                    numero_creditos, horas, cantidad_dinero, practica_UPV)

    except Exception as e:
        # Asegurarse de devolver 13 valores incluso en caso de error
        return (f"Error: {e}", None, None, None, None, None, None, None, None, None, None, None, None, None)

def procesar_pdf(pdf_path):
    """
    Crea una carpeta con el nombre basado en los datos del PDF y mueve el archivo allí. Luego copia los datos al portapapeles.
    params: pdf_path: ruta del archivo PDF a procesar
    returns: un mensaje de éxito o error
    """
    (nombre_carpeta, nombre_completo, tipo_convenio, dni, 
     nombre_empresa, cif_empresa, correo_empresa,
     fecha_inicio, fecha_fin, codigo_titulacion,
     numero_creditos, horas, cantidad_dinero, practica_upv) = extraer_datos(pdf_path)
    
    if nombre_carpeta.startswith("Error"):
        return nombre_carpeta

    # Reemplazar signo de interrogación por Ñ en el nombre de la carpeta
    nombre_carpeta = nombre_carpeta.replace("?", "Ñ")
    
    ruta_destino = os.path.join(RUTA_BASE, nombre_carpeta)
    
    try:
        os.makedirs(ruta_destino, exist_ok=True)
        nuevo_pdf_path = os.path.join(ruta_destino, os.path.basename(pdf_path))
        shutil.move(pdf_path, nuevo_pdf_path)
        nuevo_nombre_pdf = os.path.join(ruta_destino, f"{nombre_carpeta}.pdf")
        os.rename(nuevo_pdf_path, nuevo_nombre_pdf)

        nombre_formateado = nombre_carpeta.split(" - Conv")[0]

        # Preparar datos para el portapapeles
        hoy = datetime.now().strftime("%d/%m/%Y")
        pract_upv = "Pract UPV" if practica_upv else ""
        celdas = [
            hoy,                    # Fecha actual
            "",                     # Celda vacía
            "Correo electrónico",   # Texto fijo
            "Convenio",             # Texto fijo
            "",                     # Celda vacía
            nombre_formateado,      # Nombre completo
            dni,                    # DNI
            "!",                    # Celda vacía
            nombre_empresa,         # Empresa
            cif_empresa,            # CIF
            correo_empresa,         # Correo empresa
            fecha_inicio,           # Fecha inicio
            fecha_fin,              # Fecha fin
            codigo_titulacion,      # Código titulación
            numero_creditos,        # Número de créditos (si aplica)
            horas,                  # Horas
            cantidad_dinero,        # Cantidad de dinero
            "",                     # **Online?
            "",                     # **Num Convenio
            "Carlos",               # Nombre del gestor
            "",                     # Celda vacía
            "",                     # Número de teléfono
            "",                     # Celda vacía
            "",                     # Celda vacía
            "",                     # Celda vacía
            pract_upv               # Practica UPV si aplica
        ]
        
        copiar_a_portapapeles_excel(celdas)

        if pract_upv == "":
            pract_upv = "No"

        return (f"✅ PDF procesado correctamente\n"
                f"📂 Ruta: {nuevo_nombre_pdf}\n"
                f"📋 Datos copiados:\n"
                f"  - Fecha: {hoy}\n"
                f"  - Nombre: {nombre_formateado}\n"
                f"  - DNI: {dni}\n"
                f"  - Empresa: {nombre_empresa}\n"
                f"  - CIF: {cif_empresa}\n"
                f"  - Correo: {correo_empresa}\n"
                f"  - Inicio: {fecha_inicio}\n"
                f"  - Fin: {fecha_fin}\n"
                f"  - Titulación: {codigo_titulacion}\n"
                f"  - Créditos: {numero_creditos}\n"
                f"  - Horas: {horas}\n"
                f"  - Cantidad: {cantidad_dinero} €\n"
                f"  - Practica UPV: {pract_upv}\n")

    except Exception as e:
        return f"Error al procesar PDF: {e}"

class PDFHandler(FileSystemEventHandler):
    def __init__(self):
        self.info_win = None  # Ventana de resultado

    def on_created(self, event):
        if not event.is_directory and event.src_path.lower().endswith(".pdf"):
            # Esperar a que el archivo tenga tamaño > 0 y esté desbloqueado
            for _ in range(10):
                if os.path.getsize(event.src_path) > 0:
                    try:
                        with open(event.src_path, "rb") as f:
                            f.read(1)
                        break
                    except Exception:
                        time.sleep(0.5)
                else:
                    time.sleep(0.5)
            else:
                print(f"Archivo vacío o bloqueado: {event.src_path}")
                messagebox.showerror("Error", f"El archivo está vacío o bloqueado: {event.src_path}")
                return

            resultado = procesar_pdf(event.src_path)
            print(resultado)
            try:
                if str(resultado).lower().startswith("error"):
                    root.lift()
                    root.attributes('-topmost', True)
                    root.after(1000, lambda: root.attributes('-topmost', False))
                    messagebox.showerror("Error", f"{resultado}\n\nEste igual toca hacerlo manualmente.")
                else:
                    # Cerrar ventana anterior si existe
                    if self.info_win is not None and self.info_win.winfo_exists():
                        self.info_win.destroy()
                    self.info_win = tk.Toplevel(root)
                    self.info_win.title("Resultado")
                    self.info_win.geometry("400x365")
                    self.info_win.attributes('-topmost', True)
                    tk.Label(self.info_win, text=resultado, justify="left", wraplength=380, font=("Arial", 11)).pack(padx=15, pady=15)
                    tk.Button(self.info_win, text="Cerrar", command=self.info_win.destroy).pack(pady=10)
            except:
                pass  # Por si no hay ventana activa

# Al crear el observer, usa: event_handler = PDFHandler()
if __name__ == "__main__":
    # Lanzar interfaz gráfica
    root = TkinterDnD.Tk()
    app = App(root)
    root.mainloop()