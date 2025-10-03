import os
import fitz  # PyMuPDF
import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
from tkinter import messagebox, filedialog

# Carpeta de descargas del usuario
descargas_path = os.path.join(os.environ["USERPROFILE"], "Downloads")

root = TkinterDnD.Tk()
root.title("Arrastra y suelta el PDF aquí")
root.geometry("450x300")

# Variable global para la ruta de la carpeta (después de crear root)
carpeta_guardado = tk.StringVar(root)
carpeta_guardado.set(descargas_path)

def extraer_texto_pdf(pdf_path):
    try:
        with fitz.open(pdf_path) as doc:
            texto = "\n".join(page.get_text("text") for page in doc)  # Extrae todo el texto del PDF
        return texto
    except Exception as e:
        return f"Error: {e}"

def procesar_pdf(pdf_path):
    texto = extraer_texto_pdf(pdf_path)
    
    if texto.startswith("Error"):
        return False, texto  # Retorna False y el mensaje de error

    txt_file_path = f"{carpeta_guardado.get()}/Texto_pdf.txt"
    try:
        with open(txt_file_path, "w", encoding="utf-8") as txtdoc:
            txtdoc.write(texto)
        return True, f"Extracción completada con éxito.\nGuardado en:\n{txt_file_path}"
    except Exception as e:
        return False, f"Error al escribir el archivo TXT: {e}"

def seleccionar_carpeta():
    carpeta = filedialog.askdirectory()
    if carpeta:
        carpeta_guardado.set(carpeta)

def on_drop(event):
    pdf_path = event.data.strip().strip("{}")  # Limpia la ruta del archivo (a veces viene entre llaves)
    print(f"Archivo PDF recibido: {pdf_path}")

    success, mensaje = procesar_pdf(pdf_path)

    if success:
        messagebox.showinfo("Éxito", mensaje)
    else:
        messagebox.showerror("Error", mensaje)

frame = tk.Frame(root)
root.iconbitmap("./assets/pdf_txt.ico")
frame.pack(fill="x", padx=10, pady=10)

tk.Label(frame, text="Carpeta de guardado:", font=("Arial", 10)).pack(side="left")
tk.Entry(frame, textvariable=carpeta_guardado, width=35).pack(side="left", padx=5)
tk.Button(frame, text="Seleccionar...", command=seleccionar_carpeta).pack(side="left")

label = tk.Label(root, text="Arrastra el archivo PDF aquí", font=("Arial", 12), padx=10, pady=10, borderwidth=2, relief="solid")
label.pack(expand=True, fill="both", padx=10, pady=10)

# Habilitar arrastrar y soltar
root.drop_target_register(DND_FILES)
root.dnd_bind("<<Drop>>", on_drop)

# Iniciar ventana
root.mainloop()