# ----------------------------------------------------------------------
# Autor: James Palacio
#        Jacobo Chica
#
# Fecha: 04/09/2025
#
# Descripción: Integración de planeación SEO con base de datos y calendario
# ----------------------------------------------------------------------

### ---------- Librerias del codigo ---------- ###
import pandas as pd
import sqlite3
from pathlib import Path
from openpyxl import load_workbook

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
name_column_fecha_publicacion = 'Fecha de publicación'
name_column_partner = 'Partner'
name_column_blog_plataforma = 'Blog/Plataforma'

# Lista para guardar todos los DataFrames
dataframes = []
hojas_a_ignorar = ['Contenido futuro', 'Rendimiento contenido']

# Recorremos los archivos Excel de la carpeta
for archivo in carpeta.glob("*.xlsx"):
    xls = pd.ExcelFile(archivo)
    for hoja in xls.sheet_names:
        if hoja in hojas_a_ignorar:
            continue # Ignorar hojas no necesarias

        df = pd.read_excel(xls, sheet_name=hoja)

        # Normalizamos nombres de columnas desde el inicio
        df.columns = df.columns.str.strip()

        #Si el archivo es ecuabet, reemplazar los valores que hay en la columna tema, enviandolos a la columna tema del blog
        if name_archivo_partner_ecuabet in archivo.name and 'Tema' in df.columns:
            df[name_column_tema_blog] = df['Tema']
            df = df.drop(columns=['Tema'])

        #Eliminar filas con NULL en columnas clave
        columnas_obligatorias = [
            name_column_blog_plataforma,
            name_column_tema_blog,
            name_column_fecha_produccion,
            name_column_responsable_produccion,
            name_column_tiempo_estimado_produc,
            name_column_fecha_publicacion,
            name_column_responsable_publicacion,
            name_column_tiempo_estimado_public
        ]

        cols_existentes = [c for c in columnas_obligatorias if c in df.columns]
        if cols_existentes:
            df = df.dropna(subset=cols_existentes, how="all")

        df["Archivo_Origen"] = archivo.name
        df["Hoja_Origen"] = hoja
        dataframes.append(df)

# Concatenamos todos los DataFrames
df_generalizado = pd.concat(dataframes, ignore_index=True)
df_generalizado.columns = df_generalizado.columns.str.strip()

# Mapeo de partners
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

# Filtrar columnas necesarias
columnas_necesarias = [
    name_column_archivo_origen, name_column_partner, name_column_blog_plataforma, name_column_tema_blog,
    name_column_fecha_produccion, name_column_responsable_produccion,
    name_column_tiempo_estimado_produc, name_column_fecha_publicacion,
    name_column_responsable_publicacion, name_column_tiempo_estimado_public
]
df_col_necesarias = df_generalizado[columnas_necesarias].copy()
df_col_necesarias.columns = df_col_necesarias.columns.str.strip()

### ---------- Normalización de fechas ---------- ###
def convertir_columna_fecha(df, col):
    if col not in df.columns:
        print(f"⚠️ La columna '{col}' no existe.")
        return df
    if pd.api.types.is_numeric_dtype(df[col]):  # Serial Excel
        df[col] = pd.to_datetime(df[col], unit="d", origin="1899-12-30", errors="coerce")
    else:
        df[col] = pd.to_datetime(df[col], errors="coerce")
    return df

for col in [name_column_fecha_produccion, name_column_fecha_publicacion]:
    df_col_necesarias = convertir_columna_fecha(df_col_necesarias, col)
    df_col_necesarias[col] = pd.to_datetime(df_col_necesarias[col], errors="coerce").dt.date

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

#---------------------------------------------------------------------------------------------------------------------#
# --- Parte del calendario ---
nombre_plantilla = "septiembre.xlsx"
nombre_salida = "septiembre_lleno.xlsx"

try:
    with sqlite3.connect(rutadb) as conn:
        query = f"""
        SELECT *
        FROM "{name_tabla_general}"
        WHERE STRFTIME('%Y-%m', "Fecha de producción") = '2025-09'
           OR STRFTIME('%Y-%m', "Fecha de publicación") = '2025-09';
        """
        df_septiembre = pd.read_sql_query(query, conn)

    if df_septiembre.empty:
        print("⚠️ No se encontraron tareas para septiembre. El archivo de salida estará vacío.")
        exit()
    print(f"✅ Se encontraron {len(df_septiembre)} tareas para el mes de septiembre.")

except Exception as e:
    print(f"❌ Error al conectar a la base de datos o leer los datos: {e}")
    exit()

def llenar_hoja_calendario(wb, df, nombre_hoja, mapa_celdas):
    if not mapa_celdas:
        print(f"⚠️ El mapa de celdas para la hoja '{nombre_hoja}' está vacío. Se omitirá.")
        return
    if nombre_hoja not in wb.sheetnames:
        print(f"❌ Error: La hoja '{nombre_hoja}' no se encuentra en el archivo.")
        return

    ws = wb[nombre_hoja]

    for _, row in df.iterrows():
        blog_plataforma = row.get('Blog/Plataforma', '')
        partner = row.get('Partner', '')
        tema = row.get('Tema del blog', '')
        resp_prod = row.get('Responsable de produccion', '')
        resp_pub = row.get('Responsable de publicacion', '')
        tiempo_prod = row.get('Tiempo estimado produccion', '')
        tiempo_pub = row.get('Tiempo estimado publicacion', '')

        # --- Producción ---
        fecha_prod = str(row.get('Fecha de producción', ''))
        if fecha_prod in mapa_celdas:
            celda_destino = mapa_celdas[fecha_prod]
            col = ws[celda_destino].column  # columna fija
            row_start = ws[celda_destino].row  # fila inicial

            # buscar la siguiente fila vacía en esa columna
            fila_actual = row_start
            while ws.cell(row=fila_actual, column=col).value:
                fila_actual += 1

            contenido_tarea = (
                f"HACER {blog_plataforma} {partner} {tiempo_prod}\n"
                f"{tema}\n"
                f"Responsable hacer: {resp_prod}\n"
                f"Responsable montar: {resp_pub}"
            )
            ws.cell(row=fila_actual, column=col, value=contenido_tarea)

        # --- Publicación ---
        fecha_pub = str(row.get('Fecha de publicación', ''))
        if fecha_pub in mapa_celdas:
            celda_destino = mapa_celdas[fecha_pub]
            col = ws[celda_destino].column
            row_start = ws[celda_destino].row

            fila_actual = row_start
            while ws.cell(row=fila_actual, column=col).value:
                fila_actual += 1

            contenido_tarea = (
                f"PUBLICAR {blog_plataforma} {partner} {tiempo_pub}\n"
                f"{tema}\n"
                f"Responsable hacer: {resp_prod}\n"
                f"Responsable montar: {resp_pub}"
            )
            ws.cell(row=fila_actual, column=col, value=contenido_tarea)

# --- Ejecutar llenado de calendario ---
try:
    wb = load_workbook(nombre_plantilla)

    df_manuela = df_septiembre[df_septiembre['Responsable de produccion'].str.lower().fillna('').str.contains('manuela')]
    df_juan_manuel = df_septiembre[df_septiembre['Responsable de produccion'].str.lower().fillna('').str.contains('juan manuel')]
    df_santiago = df_septiembre[df_septiembre['Responsable de produccion'].str.lower().fillna('').str.contains('santiago')]

    # Definir mapa_manuela antes de usarlo
    mapa_manuela = {
        '2025-09-01': 'C4', 
        '2025-09-02': 'E4', 
        '2025-09-03': 'G4', 
        '2025-09-04': 'I4', 
        '2025-09-05': 'K4',
        '2025-09-08': 'C22', 
        '2025-09-09': 'E22', 
        '2025-09-10': 'G22', 
        '2025-09-11': 'I22', 
        '2025-09-12': 'K22',
        '2025-09-15': 'C39', 
        '2025-09-16': 'E39', 
        '2025-09-17': 'G39', 
        '2025-09-18': 'I39', 
        '2025-09-19': 'K39',
        '2025-09-22': 'C56', 
        '2025-09-23': 'E56', 
        '2025-09-24': 'G56', 
        '2025-09-25': 'I56', 
        '2025-09-26': 'K56',
        '2025-09-29': 'C74', 
        '2025-09-30': 'E74',
    }
    # Definir mapa_juan_manuel antes de usarlo
    mapa_juan_manuel = {
        '2025-09-01': 'C4', 
        '2025-09-02': 'E4', 
        '2025-09-03': 'G4', 
        '2025-09-04': 'I4', 
        '2025-09-05': 'K4',
        '2025-09-08': 'C22', 
        '2025-09-09': 'E22', 
        '2025-09-10': 'G22', 
        '2025-09-11': 'I22', 
        '2025-09-12': 'K22',
        '2025-09-15': 'C39', 
        '2025-09-16': 'E39', 
        '2025-09-17': 'G39', 
        '2025-09-18': 'I39', 
        '2025-09-19': 'K39',
        '2025-09-22': 'C56', 
        '2025-09-23': 'E56', 
        '2025-09-24': 'G56', 
        '2025-09-25': 'I56', 
        '2025-09-26': 'K56',
        '2025-09-29': 'C74', 
        '2025-09-30': 'E74',
    }
    # Definir mapa_santiago antes de usarlo
    mapa_santiago = {
        '2025-09-01': 'C4', 
        '2025-09-02': 'E4', 
        '2025-09-03': 'G4', 
        '2025-09-04': 'I4', 
        '2025-09-05': 'K4',
        '2025-09-08': 'C22', 
        '2025-09-09': 'E22', 
        '2025-09-10': 'G22', 
        '2025-09-11': 'I22', 
        '2025-09-12': 'K22',
        '2025-09-15': 'C39', 
        '2025-09-16': 'E39', 
        '2025-09-17': 'G39', 
        '2025-09-18': 'I39', 
        '2025-09-19': 'K39',
        '2025-09-22': 'C56', 
        '2025-09-23': 'E56', 
        '2025-09-24': 'G56', 
        '2025-09-25': 'I56', 
        '2025-09-26': 'K56',
        '2025-09-29': 'C74', 
        '2025-09-30': 'E74',
    }

    llenar_hoja_calendario(wb, df_manuela, 'tareas manuela septiembre', mapa_manuela)
    llenar_hoja_calendario(wb, df_juan_manuel, 'tareas juan manuel septiembre', mapa_juan_manuel)
    llenar_hoja_calendario(wb, df_santiago, 'tareas de santiago septiembre', mapa_santiago)

    wb.save(nombre_salida)
    print(f"\n✅ El calendario se ha llenado y guardado como '{nombre_salida}'.")

except FileNotFoundError:
    print(f"❌ Error: El archivo de plantilla '{nombre_plantilla}' no se encuentra.")
except Exception as e:
    print(f"\n❌ Ocurrió un error al procesar el archivo Excel: {e}")


