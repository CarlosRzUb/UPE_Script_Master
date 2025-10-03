import tkinter as tk
from tkinter import messagebox

def calcular_letra_dni():
    dni = entry_dni.get()
    if not dni.isdigit() or len(dni) != 8:
        messagebox.showerror("Error", "Introduce un DNI válido de 8 cifras.")
        return
    letras = "TRWAGMYFPDXBNJZSQVHLCKE"
    letra = letras[int(dni) % 23]
    label_resultado.config(text=f"La letra del DNI es: {letra}")

root = tk.Tk()
root.iconbitmap("./assets/dni.ico")
root.title("Calculadora Letra DNI")

tk.Label(root, text="Introduce el DNI (8 cifras, sin letra):").pack(padx=10, pady=5)
entry_dni = tk.Entry(root)
entry_dni.pack(padx=10, pady=5)

tk.Button(root, text="Calcular Letra", command=calcular_letra_dni).pack(padx=10, pady=5)
label_resultado = tk.Label(root, text="")
label_resultado.pack(padx=10, pady=10)

root.bind('<Return>', lambda event: calcular_letra_dni())

root.mainloop()