# ----------------------------------------------------------------------
# Autor: James Palacio
#        Jacobo Chica
#
# Fecha: 04/09/2025
#
# Descripción: POR DETALLAR 
# ----------------------------------------------------------------------

### ---------- Librerias del codigo ---------- ###

import pandas as pd
import sqlite3
from pathlib import Path


# Ruta de la carpeta de los excels
carpeta = Path("Archivos/archivos_origen")

### ---------- Columnas necesarias ---------- ###


name_column_archivo_origen = 'Archivo_Origen'
name_column_tiempo_estimado_produc = 'Tiempo estimado produccion'
name_column_tiempo_estimado_public = 'Tiempo estimado publicacion'
name_column_tema_blog = 'Tema del blog'
name_column_responsable_produccion = 'Responsable de produccion'
name_column_responsable_publicacion = 'Responsable de publicacion'
name_column_fecha_produccion = 'Fecha de producción'
name_column_fecha_publicacion = 'Fecha de publicación '



# Lista para guardar todos los DataFrames
dataframes = []

# Recorremos los archivos Excel de la carpeta
for archivo in carpeta.glob("*.xlsx"):
    # Obtenemos todas las hojas del archivo
    xls = pd.ExcelFile(archivo)
    
    for hoja in xls.sheet_names:
        # Leemos la hoja
        df = pd.read_excel(xls, sheet_name=hoja)
        
        # Agregamos columnas de trazabilidad
        df["Archivo_Origen"] = archivo.name
        df["Hoja_Origen"] = hoja
        
        # Guardamos en la lista
        dataframes.append(df)
        

# Concatenamos todos los DataFrames
df_generalizado = pd.concat(dataframes, ignore_index=True)

print("Total filas:", len(df_generalizado))
print("Columnas:", df_generalizado.columns.tolist())
print(df_generalizado.head())

df_generalizado.to_excel("dataframe_completo.xlsx", index=False)

columnas_necesarias = [name_column_archivo_origen, name_column_tema_blog, name_column_fecha_produccion, name_column_responsable_produccion, name_column_tiempo_estimado_produc, name_column_fecha_publicacion, name_column_responsable_publicacion, name_column_tiempo_estimado_public]

df_col_necesarias = df_generalizado[columnas_necesarias].copy()

df_col_necesarias.to_excel("dataframe_col_necesarias.xlsx", index=False)

print(df_generalizado.columns.tolist())






