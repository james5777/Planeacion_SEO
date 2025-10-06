# import tkinter as tk
# from tkinter import filedialog, messagebox
# import subprocess
# import sys
# from pathlib import Path
# import os
# import main as main_module  # Importa el módulo main.py
# from datetime import date # Para obtener el año actual

# meses_es = main_module.meses_es  # Reutiliza el diccionario de meses en español



# def seleccionar_carpeta():
#     carpeta = filedialog.askdirectory()
#     if carpeta:
#         entry_carpeta.delete(0, tk.END)
#         entry_carpeta.insert(0, carpeta)

# def ejecutar():
#     carpeta_str = entry_carpeta.get()
#     nombre_mes = mes_seleccionado.get()
#     año_str = año_seleccionado.get()

#     #Validaciones
#     if not carpeta_str or not año_str:
#         messagebox.showerror("Error", "Por favor selecciona una carpeta, un mes y un año.")
#         return

#     try:
#         # Mapear el nombre del mes con su numero (1_12)
#         # invertir valor del diccionario a la clave
#         mapeo_mes_a_num = {v: k for k, v in meses_es.items()}
#         mes_int = mapeo_mes_a_num[nombre_mes]
#         año_int = int(año_str)
        
#         #Configuracion de variable para (RUTA CARPETA)
#         os.environ["RUTA_CARPETA"] = carpeta_str

#         #Se llama a la funcion desde main.py pasando el mes y año como argumentos
#         main_module.run_main(año_int, mes_int)
#         messagebox.showinfo("Exito", f"Proceso finalizado para {nombre_mes} de {año_str}.")

#     except Exception as e:
#         # Se capturan todos los errores
#         messagebox.showerror("Error en ejecutcion", f"Ocurrió un error: {str(e)}")


# # ---- Interfaz Tkinter ----
# root = tk.Tk()
# root.title("Automatización Planeación SEO")
# root.geometry("400x300")
# root.resizable(False, False)

# # Frame principal para centrar los elementos
# frame_main = tk.Frame(root, padx=20, pady=20)
# frame_main.pack(expand=True, fill='both')

# # Selección carpeta
# tk.Label(frame_main, text="Carpeta de archivos origen:").pack(pady=5)
# frame_carpeta = tk.Frame(frame_main)
# frame_carpeta.pack(fill='x')
# entry_carpeta = tk.Entry(frame_carpeta, width=30)
# entry_carpeta.pack(side=tk.LEFT, padx=5, fill='x', expand=True)
# btn_carpeta = tk.Button(frame_carpeta, text="Seleccionar", command=seleccionar_carpeta)
# btn_carpeta.pack(side=tk.LEFT)

# # Año (Desplegable)
# tk.Label(frame_main, text="Seleccione el Año:").pack()
# entry_año = tk.Entry(root)
# entry_año.pack()

# # Generacion de lista de años
# año_actual = date.today().year
# años_disponibles = [str(a) for a in range(año_actual, año_actual + 3)]

# año_seleccionado = tk.StringVar(root)
# año_seleccionado.set(str(año_actual)) #Valor inicial
# menu_año = tk.OptionMenu(frame_main, año_seleccionado, *años_disponibles)
# menu_año.config(width=15)
# menu_año.pack(pady=5)

# # Mes(Desplegable)
# tk.Label(frame_main, text="Seleccione el mes:").pack(pady=5)


# # Usar nombres de los meses del diccionario
# opciones_mes = list(meses_es.values())

# mes_seleccionado = tk.StringVar(root)
# # se intenta establecer el mes actual como valor inicial

# try:
#     mes_inicial = meses_es[date.today().month]
# except KeyError:
#     mes_inicial = "enero" # Por defecto si hay error
# mes_seleccionado.set(mes_inicial)

# menu_mes = tk.OptionMenu(frame_main, mes_seleccionado, *opciones_mes)
# menu_mes.config(width=15)
# menu_mes.pack(pady=5)

# # Botón ejecutar
# btn_ejecutar = tk.Button(frame_main, text="Ejecutar", command=ejecutar, bg="green", fg="white", font=("Arial", 12, "bold"))
# btn_ejecutar.pack(pady=15, ipady=5)

# root.mainloop()


import tkinter as tk
from tkinter import filedialog, messagebox
# Importamos ttkbootstrap en lugar de tkinter.ttk
import ttkbootstrap as tkb 
from ttkbootstrap.constants import * # Importa constantes como PRIMARY, SUCCESS, etc.

import os
from pathlib import Path
from datetime import date
import main as main_module 	# Importa el módulo main.py

# Reutiliza el diccionario de meses en español
meses_es = main_module.meses_es 


def seleccionar_carpeta():
	"""Abre un diálogo para seleccionar una carpeta y actualiza el campo de entrada."""
	carpeta = filedialog.askdirectory()
	if carpeta:
		entry_carpeta.delete(0, tk.END)
		entry_carpeta.insert(0, carpeta)

def ejecutar():
	"""
	Recoge los datos de la interfaz, valida y ejecuta la lógica principal 
	en main.py, manejando errores.
	"""
	carpeta_str = entry_carpeta.get()
	nombre_mes = mes_seleccionado.get()
	año_str = año_seleccionado.get()

	# Validaciones de campos
	if not carpeta_str or not año_str or not nombre_mes:
		messagebox.showerror("Error", "Por favor, completa todos los campos (Carpeta, Mes y Año).")
		return

	try:
		# Mapear el nombre del mes con su número (1-12)
		mapeo_mes_a_num = {v: k for k, v in meses_es.items()}
		mes_int = mapeo_mes_a_num[nombre_mes]
		año_int = int(año_str)
		
		# Configuración de variable de entorno para RUTA_CARPETA
		os.environ["RUTA_CARPETA"] = carpeta_str

		# Llamada a la función principal de main.py
		main_module.run_main(año_int, mes_int)
		
		messagebox.showinfo("Éxito", f"✅ Proceso finalizado para {nombre_mes} de {año_str}.")

	except Exception as e:
		# Se capturan todos los errores que puedan ocurrir
		messagebox.showerror("Error en ejecución", f"Ocurrió un error: {str(e)}")


# ---- Interfaz con ttkbootstrap ----
# Usamos tkb.Window en lugar de tk.Tk para aplicar el tema Bootstrap.
root = tkb.Window(themename="flatly") # Puedes probar otros temas como 'cosmo', 'journal', 'litera'
root.title("Automatización Planeación SEO")
root.geometry("500x450") 
root.resizable(False, False)

# El estilo se maneja automáticamente por el tema, no necesitamos ttk.Style manual.

# Frame principal para centrar los elementos y aplicar padding general
# Usamos Frame de ttkbootstrap
frame_main = tkb.Frame(root, padding="25 25 25 25")
frame_main.pack(expand=True, fill='both')

# --- 1. Sección de Carpeta ---
# Usamos Label de ttkbootstrap
tkb.Label(frame_main, text="1. Carpeta de archivos origen:", anchor="w").pack(pady=(0, 5), fill='x')
frame_carpeta = tkb.Frame(frame_main)
frame_carpeta.pack(fill='x')
entry_carpeta = tkb.Entry(frame_carpeta)
entry_carpeta.pack(side=tk.LEFT, padx=(0, 10), fill='x', expand=True)

# Usamos tkb.Button con un estilo 'secondary' para el botón de selección
btn_carpeta = tkb.Button(frame_carpeta, text="Seleccionar", command=seleccionar_carpeta, bootstyle=SECONDARY)
btn_carpeta.pack(side=tk.LEFT, ipadx=5, ipady=3)


# --- 2. Sección de Año (Desplegable) ---
tkb.Label(frame_main, text="2. Seleccione el Año:", anchor="w").pack(pady=(15, 5), fill='x')
frame_año = tkb.Frame(frame_main)
frame_año.pack(fill='x')

año_actual = date.today().year
años_disponibles = [str(a) for a in range(año_actual, año_actual + 3)]
año_seleccionado = tk.StringVar(root)
año_seleccionado.set(str(año_actual)) 

# Usamos tkb.Combobox con un estilo moderno
menu_año = tkb.Combobox(frame_año, textvariable=año_seleccionado, values=años_disponibles, state='readonly', bootstyle=PRIMARY)
menu_año.pack(pady=5, fill='x')
menu_año.set(str(año_actual)) 


# --- 3. Sección de Mes (Desplegable) ---
tkb.Label(frame_main, text="3. Seleccione el Mes:", anchor="w").pack(pady=(15, 5), fill='x')
frame_mes = tkb.Frame(frame_main)
frame_mes.pack(fill='x')

opciones_mes = list(meses_es.values())
mes_seleccionado = tk.StringVar(root)
try:
	mes_inicial = meses_es[date.today().month]
except KeyError:
	mes_inicial = "enero" 

# Usamos tkb.Combobox con un estilo moderno
menu_mes = tkb.Combobox(frame_mes, textvariable=mes_seleccionado, values=opciones_mes, state='readonly', bootstyle=PRIMARY)
menu_mes.pack(pady=5, fill='x')
menu_mes.set(mes_inicial) 

# --- Botón ejecutar ---
# Usamos tkb.Button y el estilo predefinido 'SUCCESS' (botón verde)
btn_ejecutar = tkb.Button(
	frame_main, 
	text="EJECUTAR", 
	command=ejecutar,
	bootstyle=SUCCESS # Usa el color verde de Bootstrap
)
btn_ejecutar.pack(pady=30, fill='x', ipady=5) 

root.mainloop()
