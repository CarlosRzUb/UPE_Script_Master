# filepath: script-launcher-app/scripts/blog/PDFFromMsgDownloader.py
import os
import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import filedialog

import extract_msg

def process_files(files):
    files = root.tk.splitlist(files)
    for file in files:
        if file.endswith('.msg'):
            msg = extract_msg.Message(file)
            for attachment in msg.attachments:
                if attachment.longFilename.endswith('.pdf'):
                    save_path = os.path.join(carpeta_guardado.get(), attachment.longFilename)
                    os.makedirs(os.path.dirname(save_path), exist_ok=True)
                    with open(save_path, 'wb') as f:
                        f.write(attachment.data)
            msg.close()
            os.remove(file)

def drop(event):
    process_files(event.data)

def seleccionar_carpeta():
    carpeta = filedialog.askdirectory()
    if carpeta:
        carpeta_guardado.set(carpeta)

root = TkinterDnD.Tk()
root.iconbitmap("./assets/msg_pdf.ico")
root.title("PDF Downloader")
root.geometry("450x300")

# Carpeta de descargas del usuario por defecto
carpeta_guardado = tk.StringVar(root)
carpeta_guardado.set(os.path.join(os.environ["USERPROFILE"], "Downloads"))

frame = tk.Frame(root)
frame.pack(fill="x", padx=10, pady=10)

tk.Label(frame, text="Carpeta de guardado:", font=("Arial", 10)).pack(side="left")
tk.Entry(frame, textvariable=carpeta_guardado, width=35).pack(side="left", padx=5)
tk.Button(frame, text="Seleccionar...", command=seleccionar_carpeta).pack(side="left")

label = tk.Label(root, text="Arrastra los archivos .msg aquí", font=("Arial", 12), padx=10, pady=10, borderwidth=2, relief="solid")
label.pack(expand=True, fill="both", padx=10, pady=10)

root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', drop)

root.mainloop()