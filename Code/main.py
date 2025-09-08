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

### ---------- Nombre de archivos ---------- ###
name_archivo_partner_aciertala = 'Planeación contenido blog Aciertala 2025.xlsx'
name_archivo_partner_camanbet = 'Planeación contenido blog CamanBet 2025.xlsx'
name_archivo_partner_doradobet_cr = 'Planeación contenido blog Doradobet CR 2025.xlsx'
name_archivo_partner_doradobet_gt = 'Planeación contenido blog Doradobet GT 2025.xlsx'
name_archivo_partner_doradobet_pe = 'Planeación contenido blog Doradobet PE 2025.xlsx'
name_archivo_partner_doradobet_sv = 'Planeación Doradobet El Salvador 2025.xlsx'
name_archivo_partner_ecuabet = 'Planeación contenido blog Ecuabet 2025.xlsx'
name_archivo_partner_ganaplay_gt = 'Planeación contenido blog GanaPlay GT 2025.xlsx'
name_archivo_partner_ganaplay_sv = 'Planeación contenido blog GanaPlay SV 2025.xlsx'
name_archivo_partner_paniplay = 'Planeación contenido blog PaniPlay 2025.xlsx'

### ---------- Nombre de partners ---------- ###
name_partner_aciertala = 'Aciertala'
name_partner_camanbet = 'Camanbet'
name_partner_doradobet_cr = 'Doradobet CR'
name_partner_doradobet_gt = 'Doradobet GT'
name_partner_doradobet_pe = 'Doradobet PE'
name_partner_doradobet_sv = 'Doradobet SV'
name_partner_ecuabet = 'Ecuabet'
name_partner_ganaplay_gt = 'Ganaplay GT'
name_partner_ganaplay_sv = 'Ganaplay SV'
name_partner_paniplay = 'Paniplay'


### ---------- Columnas necesarias ---------- ###
name_column_archivo_origen = 'Archivo_Origen'
name_column_tiempo_estimado_produc = 'Tiempo estimado produccion'
name_column_tiempo_estimado_public = 'Tiempo estimado publicacion'
name_column_tema_blog = 'Tema del blog'
name_column_responsable_produccion = 'Responsable de produccion'
name_column_responsable_publicacion = 'Responsable de publicacion'
name_column_fecha_produccion = 'Fecha de producción'
name_column_fecha_publicacion = 'Fecha de publicación '
name_column_partner = 'Partner'
name_column_blog_plataforma = 'Blog/Plataforma'

# Lista para guardar todos los DataFrames
dataframes = []

# Recorremos los archivos Excel de la carpeta
for archivo in carpeta.glob("*.xlsx"):
    xls = pd.ExcelFile(archivo)
    for hoja in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=hoja)
        df["Archivo_Origen"] = archivo.name
        df["Hoja_Origen"] = hoja
        dataframes.append(df)

# Concatenamos todos los DataFrames
df_generalizado = pd.concat(dataframes, ignore_index=True)
df_generalizado.to_excel("dataframe_completo.xlsx", index=False)

#  Mapeo de Archivo_Origen, para asignar partners
mapeo_partners = {
    name_archivo_partner_aciertala : name_partner_aciertala,
    name_archivo_partner_camanbet : name_partner_camanbet,
    name_archivo_partner_doradobet_cr : name_partner_doradobet_cr,
    name_archivo_partner_doradobet_gt : name_partner_doradobet_gt,
    name_archivo_partner_doradobet_pe : name_partner_doradobet_pe,
    name_archivo_partner_doradobet_sv : name_partner_doradobet_sv,
    name_archivo_partner_ecuabet : name_partner_ecuabet,
    name_archivo_partner_ganaplay_gt : name_partner_ganaplay_gt,
    name_archivo_partner_ganaplay_sv : name_partner_ganaplay_sv,
    name_archivo_partner_paniplay : name_partner_paniplay
}

df_generalizado[name_column_partner] = df_generalizado[name_column_archivo_origen].map(mapeo_partners)

# Filtramos columnas necesarias
columnas_necesarias = [
    name_column_archivo_origen, name_column_partner, name_column_blog_plataforma, name_column_tema_blog,
    name_column_fecha_produccion, name_column_responsable_produccion,
    name_column_tiempo_estimado_produc, name_column_fecha_publicacion,
    name_column_responsable_publicacion, name_column_tiempo_estimado_public
]

df_col_necesarias = df_generalizado[columnas_necesarias].copy()

### ---------- Normalización de fechas ---------- ###
def convertir_columna_fecha(df, col):
    if col not in df.columns:
        print(f"⚠️ La columna '{col}' no existe.")
        return df

    if pd.api.types.is_numeric_dtype(df[col]):  # Serial Excel
        df[col] = pd.to_datetime(
            df[col], unit="d", origin="1899-12-30", errors="coerce"
        )
    else:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    return df

columnas_fecha = [name_column_fecha_produccion, name_column_fecha_publicacion]

for col in columnas_fecha:
    df_col_necesarias = convertir_columna_fecha(df_col_necesarias, col)

# 🔹 Forzamos formato día/mes/año sin ambigüedad
for col in columnas_fecha:
    df_col_necesarias[col] = pd.to_datetime(
        df_col_necesarias[col], format="%d/%m/%Y", errors="coerce"
    ).dt.date



print(df_col_necesarias[columnas_fecha].head())
print(df_col_necesarias[columnas_fecha].dtypes)
print(df_col_necesarias[columnas_fecha].isna().sum())

### ---------- Guardar en SQLite ---------- ###
def guardar_en_sqlite(df: pd.DataFrame, nombre_tabla: str, ruta_db: Path, if_exists: str = "replace") -> None:
    if df.empty:
        print(f"\n ⚠️ El DataFrame está vacío. No se insertaron datos en la tabla '{nombre_tabla}'.\n ")
        return
    try:
        with sqlite3.connect(ruta_db) as conn:
            df.to_sql(nombre_tabla, conn, if_exists=if_exists, index=False)
        print(f"\n ✅ Se insertaron los datos en la tabla: '{nombre_tabla}' en la base de datos '{ruta_db.name}'.\n ")
    except Exception as e:
        print(f"\n ❌ Error al guardar en SQLite: {e}")

guardar_en_sqlite(df_col_necesarias, name_tabla_general, rutadb)


