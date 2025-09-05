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

### ---------- Ruta de la base de datos ---------- ###
rutadb = Path("Archivos/Archivos_base_de_datos/Archivo_base_de_datos.db")

### ---------- Nombre de las tablas en la base de datos ---------- ###
name_tabla_general = "Datos_generales"

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

# print("Total filas:", len(df_generalizado))
# print("Columnas:", df_generalizado.columns.tolist())
# print(df_generalizado.head())

df_generalizado.to_excel("dataframe_completo.xlsx", index=False)

columnas_necesarias = [name_column_archivo_origen, name_column_tema_blog, name_column_fecha_produccion, name_column_responsable_produccion, name_column_tiempo_estimado_produc, name_column_fecha_publicacion, name_column_responsable_publicacion, name_column_tiempo_estimado_public]

df_col_necesarias = df_generalizado[columnas_necesarias].copy()

df_col_necesarias.to_excel("dataframe_col_necesarias.xlsx", index=False)


def convertir_columna_fecha(df, col):
    if col not in df.columns:
        print(f"⚠️ La columna '{col}' no existe.")
        return df
    
    # Si es numérica (serial Excel)
    if pd.api.types.is_numeric_dtype(df[col]):
        df[col] = pd.to_datetime(df[col], unit="d", origin="1899-12-30", errors="coerce")
    else:
        # Si ya es string o datetime
        df[col] = pd.to_datetime(df[col], dayfirst=True, errors="coerce")
    
    return df

# 🔹 Ejemplo con varias columnas
columnas_fecha = [name_column_fecha_produccion, name_column_fecha_publicacion]

for col in columnas_fecha:
    df_col_necesarias = convertir_columna_fecha(df_col_necesarias, col)

# df_col_necesarias[name_column_fecha_produccion] = df_col_necesarias[name_column_fecha_produccion].dt.strftime('%d/%m/%Y')
# df_col_necesarias[name_column_fecha_publicacion] = df_col_necesarias[name_column_fecha_publicacion].dt.strftime('%d/%m/%Y')

print(df_col_necesarias[columnas_fecha].head())
print(df_col_necesarias[columnas_fecha].dtypes)

print(df_col_necesarias[columnas_fecha].isna().sum())


def guardar_en_sqlite(df: pd.DataFrame, nombre_tabla: str, ruta_db: Path, if_exists: str = "replace") -> None:



    """
    Guarda un DataFrame en una base de datos SQLite, creando o actualizando la tabla según se especifique.

    Parámetros:
    ----------
    df : pd.DataFrame
        El DataFrame que se desea guardar en la base de datos.
    
    nombre_tabla : str
        El nombre de la tabla en la base de datos SQLite.
    
    ruta_db : Path
        Ruta al archivo `.sqlite` o `.db` donde se guardarán los datos.
    
    if_exists : str, opcional
        Comportamiento si la tabla ya existe. Valores permitidos:
        - 'replace' (por defecto): elimina la tabla y la vuelve a crear.
        - 'append': agrega los datos sin eliminar la tabla.
        - 'fail': lanza una excepción si la tabla ya existe.

    Retorna:
    -------
    None
        Esta función no retorna un valor. Inserta los datos directamente en la base de datos.
    """
    # Validar que el DataFrame no esté vacío
    if df.empty:
        print(f"\n ⚠️ El DataFrame está vacío. No se insertaron datos en la tabla '{nombre_tabla}'.\n ")
        return
    try:
        # Conexión a SQLite
        with sqlite3.connect(ruta_db) as conn:
            df.to_sql(nombre_tabla, conn, if_exists=if_exists, index=False)
        print(f"\n ✅ Se insertarón los datos con la tabla: '{nombre_tabla}' en la base de datos '{ruta_db.name}'.\n ")
    
    except Exception as e:
        print(f"\n ❌ Error al guardar en SQLite: {e}")

guardar_en_sqlite(df_col_necesarias,name_tabla_general, rutadb)
