import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import sys
from pathlib import Path
import os
import main as main_module  # Importa el módulo main.py

def seleccionar_carpeta():
    carpeta = filedialog.askdirectory()
    if carpeta:
        entry_carpeta.delete(0, tk.END)
        entry_carpeta.insert(0, carpeta)

def ejecutar():
    carpeta_str = entry_carpeta.get()
    año = entry_año.get()
    mes = entry_mes.get()

    if not carpeta_str or not año.isdigit() or not mes.isdigit():
        messagebox.showerror("Error", "Por favor completa todos los campos correctamente.")
        return

    try:
        # Configurar las variables de entorno para que main.py pueda leerlas
        os.environ["RUTA_CARPETA"] = carpeta_str
        os.environ["AÑO"] = año
        os.environ["MES"] = mes
        
        # Llama a la función principal de main.py
        main_module.run_main() 
        
        messagebox.showinfo("Éxito", f"✅ Proceso finalizado para {mes}/{año}")

    except Exception as e:
        messagebox.showerror("Error en ejecución", f"Ocurrió un error: {str(e)}")

# ---- Interfaz Tkinter ----
root = tk.Tk()
root.title("Automatización Planeación SEO")
root.geometry("400x250")

# Selección carpeta
tk.Label(root, text="Carpeta de archivos origen:").pack(pady=5)
frame_carpeta = tk.Frame(root)
frame_carpeta.pack()
entry_carpeta = tk.Entry(frame_carpeta, width=30)
entry_carpeta.pack(side=tk.LEFT, padx=5)
btn_carpeta = tk.Button(frame_carpeta, text="Seleccionar", command=seleccionar_carpeta)
btn_carpeta.pack(side=tk.LEFT)

# Año
tk.Label(root, text="Año (ej: 2026):").pack(pady=5)
entry_año = tk.Entry(root)
entry_año.pack()

# Mes
tk.Label(root, text="Mes (1-12):").pack(pady=5)
entry_mes = tk.Entry(root)
entry_mes.pack()

# Botón ejecutar
btn_ejecutar = tk.Button(root, text="Ejecutar", command=ejecutar, bg="green", fg="white")
btn_ejecutar.pack(pady=15)

root.mainloop()


