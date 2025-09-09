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
hojas_a_ignorar = ['Contenido futuro', 'Rendimiento contenido']

# Recorremos los archivos Excel de la carpeta
for archivo in carpeta.glob("*.xlsx"):
    xls = pd.ExcelFile(archivo)
    for hoja in xls.sheet_names:
        if hoja in hojas_a_ignorar:
            continue # Si existe una hoja en la lista de hojas, la saltamos y la ignoramos.

        df = pd.read_excel(xls, sheet_name=hoja)
        
        #Si el archivo es ecuabet, reemplazar los valores que hay en la columna tema, enviandolos a la columna tema del blog
        if name_archivo_partner_ecuabet in archivo.name and 'Tema' in df.columns:
            #Si existe "Tema del blog" , sobre escribirlo con los valores de "Tema"
            df[name_column_tema_blog] = df['Tema']
            #Eliminamos la columna "Tema"
            df = df.drop(columns=['Tema'])

        #Eliminar filas que sean NULL o NaN en las columnas 'Blog/Plataforma', 'Tema del blog', 'Fecha de produccion, 'Responsable de produccion', 'Tiempo estimado produccion', 'Fecha de publicacion', 'Responsable de publicación', 'Tiempo estimado publicacion'
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

        # Usar solo las que realmente existen en este DataFrame
        cols_existentes = [c for c in columnas_obligatorias if c in df.columns]

        # Eliminar filas nulas en esas columnas
        if cols_existentes:  # solo aplica si hay columnas válidas
            df = df.dropna(subset=cols_existentes, how="all")

        

        df["Archivo_Origen"] = archivo.name
        df["Hoja_Origen"] = hoja
        dataframes.append(df)

        


    

# Concatenamos todos los DataFrames
df_generalizado = pd.concat(dataframes, ignore_index=True)


# 🔎 Diagnóstico de columnas únicas
print("\n🔎 Columnas detectadas en df_generalizado:")
print(sorted(df_generalizado.columns.tolist()))

print("\n📊 Registros totales concatenados:", len(df_generalizado))

# 🔎 Registros por archivo
print("\n📊 Registros por archivo:")
print(df_generalizado["Archivo_Origen"].value_counts())

# 🔎 Revisar qué columnas tiene cada archivo
for archivo, df in df_generalizado.groupby("Archivo_Origen"):
    print(f"\n📂 Archivo: {archivo}")
    print("   Columnas:", sorted(df.columns.tolist()))
    print("   Registros:", len(df))

# Si ya tienes mapeo de partners:
if "partner" in df_generalizado.columns:
    print("\n📊 Registros por partner:")
    print(df_generalizado["partner"].value_counts())

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

#---------------------------------------------------------------------------------------------------------------------#
import pandas as pd
import sqlite3
from pathlib import Path
from openpyxl import load_workbook

# --- Tus rutas de archivo ---
rutadb = Path("Archivos/Archivos_base_de_datos/Archivo_base_de_datos.db")
name_tabla_general = "Datos_generales"
nombre_plantilla = "septiembre.xlsx"
nombre_salida = "septiembre_lleno.xlsx"

# --- CONEXIÓN Y CONSULTA A LA BD ---
try:
    with sqlite3.connect(rutadb) as conn:
        query = f"""
        SELECT *
        FROM "{name_tabla_general}"
        WHERE STRFTIME('%Y-%m', "Fecha de producción" ) = '2025-09';
        """
        df_septiembre = pd.read_sql_query(query, conn)
    
    if df_septiembre.empty:
        print("⚠️ No se encontraron tareas para septiembre. El archivo de salida estará vacío.")
        exit()
        
    print(f"✅ Se encontraron {len(df_septiembre)} tareas para el mes de septiembre.")

except Exception as e:
    print(f"❌ Error al conectar a la base de datos o leer los datos: {e}")
    exit()

# --- MAPAS DE CELDAS POR RESPONSABLE ---
# ⚠️ IMPORTANTE: DEBES COMPLETAR LOS MAPAS DE JUAN MANUEL Y SANTIAGO
# La plantilla de Manuela ya está completa
mapa_manuela = {
    '2025-09-01': 'C4', 
    '2025-09-02': 'E4', 
    '2025-09-03': 'G4', 
    '2025-09-04': 'I4', 
    '2025-09-05': 'K4',
    '2025-09-08': 'C16', 
    '2025-09-09': 'E16', 
    '2025-09-10': 'G16', 
    '2025-09-11': 'I16', 
    '2025-09-12': 'K16',
    '2025-09-15': 'C29', 
    '2025-09-16': 'E29', 
    '2025-09-17': 'G29', 
    '2025-09-18': 'I29', 
    '2025-09-19': 'K29',
    '2025-09-22': 'C41', 
    '2025-09-23': 'E41', 
    '2025-09-24': 'G41', 
    '2025-09-25': 'I41', 
    '2025-09-26': 'K41',
    '2025-09-29': 'C55', 
    '2025-09-30': 'E55',
}

mapa_juan_manuel = {
    '2025-09-01': 'C4', 
    '2025-09-02': 'E4', 
    '2025-09-03': 'G4', 
    '2025-09-04': 'I4', 
    '2025-09-05': 'K4',
    '2025-09-08': 'C16', 
    '2025-09-09': 'E16', 
    '2025-09-10': 'G16', 
    '2025-09-11': 'I16', 
    '2025-09-12': 'K16',
    '2025-09-15': 'C29', 
    '2025-09-16': 'E29', 
    '2025-09-17': 'G29', 
    '2025-09-18': 'I29', 
    '2025-09-19': 'K29',
    '2025-09-22': 'C41', 
    '2025-09-23': 'E41', 
    '2025-09-24': 'G41', 
    '2025-09-25': 'I41', 
    '2025-09-26': 'K41',
    '2025-09-29': 'C55', 
    '2025-09-30': 'E55',
}

mapa_santiago = {
    '2025-09-01': 'C4', 
    '2025-09-02': 'E4', 
    '2025-09-03': 'G4', 
    '2025-09-04': 'I4', 
    '2025-09-05': 'K4',
    '2025-09-08': 'C16', 
    '2025-09-09': 'E16', 
    '2025-09-10': 'G16', 
    '2025-09-11': 'I16', 
    '2025-09-12': 'K16',
    '2025-09-15': 'C29', 
    '2025-09-16': 'E29', 
    '2025-09-17': 'G29', 
    '2025-09-18': 'I29', 
    '2025-09-19': 'K29',
    '2025-09-22': 'C41', 
    '2025-09-23': 'E41', 
    '2025-09-24': 'G41', 
    '2025-09-25': 'I41', 
    '2025-09-26': 'K41',
    '2025-09-29': 'C55', 
    '2025-09-30': 'E55',
}

# --- FUNCIÓN PARA LLENAR LA HOJA ---
def llenar_hoja_calendario(wb, df, nombre_hoja, mapa_celdas):
    if not mapa_celdas:
        print(f"⚠️ El mapa de celdas para la hoja '{nombre_hoja}' está vacío. Se omitirá.")
        return
        
    if nombre_hoja not in wb.sheetnames:
        print(f"❌ Error: La hoja '{nombre_hoja}' no se encuentra en el archivo.")
        return

    ws = wb[nombre_hoja]
    
    for _, row in df.iterrows():
        fecha = str(row['Fecha de producción'])
        
        if fecha in mapa_celdas:
            celda_destino = mapa_celdas[fecha]
            contenido_tarea = f"✅ {row['Tema del blog']} - {row['Partner']}"
            
            celda_actual = ws[celda_destino]
            
            if celda_actual.value:
                celda_actual.value += f"\n\n{contenido_tarea}"
            else:
                celda_actual.value = contenido_tarea
        else:
            print(f"⚠️ No se encontró una celda de destino para la fecha {fecha} en la hoja de {ws.title}.")


# --- CARGAR LA PLANTILLA Y LLAMAR A LA FUNCIÓN ---
try:
    wb = load_workbook(nombre_plantilla)

    # Identifica las hojas de cada responsable
    df_manuela = df_septiembre[df_septiembre['Responsable de produccion'].str.lower().fillna('').str.contains('manuela')]
    df_juan_manuel = df_septiembre[df_septiembre['Responsable de produccion'].str.lower().fillna('').str.contains('juan manuel')]
    df_santiago = df_septiembre[df_septiembre['Responsable de produccion'].str.lower().fillna('').str.contains('santiago')]
    
    # Llama a la función para cada responsable
    llenar_hoja_calendario(wb, df_manuela, 'tareas manuela septiembre', mapa_manuela)
    llenar_hoja_calendario(wb, df_juan_manuel, 'tareas juan manuel septiembre', mapa_juan_manuel)
    llenar_hoja_calendario(wb, df_santiago, 'tareas de santiago septiembre', mapa_santiago)

    # --- GUARDAR LOS CAMBIOS ---
    wb.save(nombre_salida)
    print(f"\n✅ El calendario se ha llenado y guardado como '{nombre_salida}'.")

except FileNotFoundError:
    print(f"❌ Error: El archivo de plantilla '{nombre_plantilla}' no se encuentra.")
except Exception as e:
    print(f"\n❌ Ocurrió un error al procesar el archivo Excel: {e}")


