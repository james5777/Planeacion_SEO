# ----------------------------------------------------------------------
# Autor: James Palacio
#        Jacobo Chica
#
# Fecha: 09/09/2025
#
# Descripción: Integración de planeación SEO con base de datos y calendario
#              -> corregido: coincidencia responsable (contains) y semana
# ----------------------------------------------------------------------

import pandas as pd
import sqlite3
from pathlib import Path
from openpyxl import load_workbook
import calendar
from datetime import date
from openpyxl.utils import range_boundaries

### ---------- Parámetros globales ---------- ###
año = 2025
mes = 8  # Cambia aquí el mes que quieras (1-12)

# Ruta de la carpeta de los excels
carpeta = Path("Archivos/archivos_origen")

### ---------- Ruta de la base de datos ---------- ###
rutadb = Path("Archivos/Archivos_base_de_datos/Archivo_base_de_datos.db")

### ---------- Nombre de la tabla ---------- ###
name_tabla_general = "Datos_generales"

### ---------- Columnas ---------- ###
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

# ---------- Lectura y normalización de archivos fuente (igual que tenías) ----------
dataframes = []
hojas_a_ignorar = ['Contenido futuro', 'Rendimiento contenido']

for archivo in carpeta.glob("*.xlsx"):
    xls = pd.ExcelFile(archivo)
    for hoja in xls.sheet_names:
        if hoja in hojas_a_ignorar:
            continue
        df = pd.read_excel(xls, sheet_name=hoja)
        df.columns = df.columns.str.strip()

        # caso Ecuabet si tiene columna 'Tema'
        if 'Planeación contenido blog Ecuabet' in archivo.name and 'Tema' in df.columns:
            df[name_column_tema_blog] = df['Tema']
            df = df.drop(columns=['Tema'])

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

df_generalizado = pd.concat(dataframes, ignore_index=True)
df_generalizado.columns = df_generalizado.columns.str.strip()

# Mapeo de archivos -> partner (igual que tenías)
mapeo_partners = {
    'Planeación contenido blog Aciertala 2025.xlsx': 'Aciertala',
    'Planeación contenido blog CamanBet 2025.xlsx': 'Camanbet',
    'Planeación contenido blog Doradobet CR 2025.xlsx': 'Doradobet CR',
    'Planeación contenido blog Doradobet GT 2025.xlsx': 'Doradobet GT',
    'Planeación contenido blog Doradobet PE 2025.xlsx': 'Doradobet PE',
    'Planeación Doradobet El Salvador 2025.xlsx': 'Doradobet SV',
    'Planeación contenido blog Ecuabet 2025.xlsx': 'Ecuabet',
    'Planeación contenido blog GanaPlay GT 2025.xlsx': 'Ganaplay GT',
    'Planeación contenido blog GanaPlay SV 2025.xlsx': 'Ganaplay SV',
    'Planeación contenido blog PaniPlay 2025.xlsx': 'Paniplay'
}
df_generalizado[name_column_partner] = df_generalizado[name_column_archivo_origen].map(mapeo_partners)

# Filtrar columnas que usaremos
columnas_necesarias = [
    name_column_archivo_origen, name_column_partner, name_column_blog_plataforma, name_column_tema_blog,
    name_column_fecha_produccion, name_column_responsable_produccion,
    name_column_tiempo_estimado_produc, name_column_fecha_publicacion,
    name_column_responsable_publicacion, name_column_tiempo_estimado_public
]
df_col_necesarias = df_generalizado[columnas_necesarias].copy()
df_col_necesarias.columns = df_col_necesarias.columns.str.strip()

# Normalizar fechas: convertimos tanto serial Excel como strings
def convertir_columna_fecha(df, col):
    if col not in df.columns:
        return df
    if pd.api.types.is_numeric_dtype(df[col]):
        df[col] = pd.to_datetime(df[col], unit="d", origin="1899-12-30", errors="coerce")
    else:
        df[col] = pd.to_datetime(df[col], errors="coerce")
    return df

for col in [name_column_fecha_produccion, name_column_fecha_publicacion]:
    df_col_necesarias = convertir_columna_fecha(df_col_necesarias, col)
    df_col_necesarias[col] = pd.to_datetime(df_col_necesarias[col], errors="coerce").dt.date

# Guardar en sqlite
def guardar_en_sqlite(df: pd.DataFrame, nombre_tabla: str, ruta_db: Path, if_exists: str = "replace") -> None:
    if df.empty:
        print(f"\n ⚠️ El DataFrame está vacío. No se insertaron datos en la tabla '{nombre_tabla}'.\n ")
        return
    with sqlite3.connect(ruta_db) as conn:
        df.to_sql(nombre_tabla, conn, if_exists=if_exists, index=False)
    print(f"\n ✅ Se insertaron los datos en la tabla: '{nombre_tabla}' en la base de datos '{ruta_db.name}'.\n ")

guardar_en_sqlite(df_col_necesarias, name_tabla_general, rutadb)

# ------------------ Parte del calendario (dinámica por mes/año) ------------------
meses_es = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
    5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
    9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
}
nombre_mes = meses_es[mes]
nombre_plantilla = f"{nombre_mes}.xlsx"
nombre_salida = f"{nombre_mes}_lleno.xlsx"

# Traer datos del mes desde sqlite
mes_str = f"{año}-{mes:02d}"
with sqlite3.connect(rutadb) as conn:
    query = f"""
    SELECT *
    FROM "{name_tabla_general}"
    WHERE STRFTIME('%Y-%m', "Fecha de producción") = '{mes_str}'
       OR STRFTIME('%Y-%m', "Fecha de publicación") = '{mes_str}';
    """
    df_mes = pd.read_sql_query(query, conn)

# convertir fechas a date por si acaso (si vienen como strings)
for col in [name_column_fecha_produccion, name_column_fecha_publicacion]:
    if col in df_mes.columns:
        df_mes[col] = pd.to_datetime(df_mes[col], errors="coerce").dt.date

if df_mes.empty:
    print(f"⚠️ No se encontraron tareas para {nombre_mes}.")
else:
    print(f"✅ Se encontraron {len(df_mes)} tareas para el mes {nombre_mes}.")

# Mapa encabezados (tu estructura)
mapa_encabezados = {
    1: {"Lunes": "C2", "Martes": "E2", "Miércoles": "G2", "Jueves": "I2", "Viernes": "K2"},
    2: {"Lunes": "C21", "Martes": "E21", "Miércoles": "G21", "Jueves": "I21", "Viernes": "K21"},
    3: {"Lunes": "C38", "Martes": "E38", "Miércoles": "G38", "Jueves": "I38", "Viernes": "K38"},
    4: {"Lunes": "C55", "Martes": "E55", "Miércoles": "G55", "Jueves": "I55", "Viernes": "K55"},
    5: {"Lunes": "C72", "Martes": "E72", "Miércoles": "G72", "Jueves": "I72", "Viernes": "K72"},
}

def llenar_encabezados_calendario(wb, nombre_hoja, año, mes, mapa_encabezados):
    if nombre_hoja not in wb.sheetnames:
        return
    ws = wb[nombre_hoja]
    dias_semana = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
    num_dias = calendar.monthrange(año, mes)[1]
    semana = 1
    for dia in range(1, num_dias + 1):
        f = date(año, mes, dia)
        weekday = f.weekday()  # lunes=0
        if weekday < 5:
            nombre_dia = dias_semana[weekday]
            if semana in mapa_encabezados and nombre_dia in mapa_encabezados[semana]:
                ws[mapa_encabezados[semana][nombre_dia]] = f"{nombre_dia} {dia}"
        if weekday == 6:  # domingo -> siguiente semana
            semana += 1

# abrir plantilla
wb = load_workbook(nombre_plantilla)

# Detectar hojas por persona buscando substring en nombre de hoja ( tolerant )
personas = ["manuela", "juan manuel", "santiago"]
hojas_persona = {}
for p in personas:
    encontrada = next((s for s in wb.sheetnames if p in s.lower()), None)
    if encontrada:
        hojas_persona[p] = encontrada
    else:
        print(f"⚠️ No se encontró hoja para '{p}' en la plantilla (buscando substring).")

# Llenar encabezados en las hojas encontradas
for p, hoja in hojas_persona.items():
    llenar_encabezados_calendario(wb, hoja, año, mes, mapa_encabezados)
print(f"✅ Encabezados de {nombre_mes} (en las hojas detectadas) rellenados.")

# Función para calcular semana usando la misma regla de encabezados (aumenta cuando aparece domingo)
def semana_por_domingos(fecha):
    """Devuelve 1 + número de domingos ocurridos entre el día 1 y `fecha` (inclusive)."""
    if not isinstance(fecha, date):
        raise TypeError("fecha debe ser datetime.date")
    conteo_domingos = 0
    for d in range(1, fecha.day + 1):
        if date(fecha.year, fecha.month, d).weekday() == 6:
            conteo_domingos += 1
    return 1 + conteo_domingos

# Rango de celdas donde escribir tareas por semana/día
mapa_tareas = {
    1: {"Lunes": "C4:C19", "Martes": "E4:E19", "Miércoles": "G4:G19", "Jueves": "I4:I19", "Viernes": "K4:K19"},
    2: {"Lunes": "C22:C36", "Martes": "E22:E36", "Miércoles": "G22:G36", "Jueves": "I22:I36", "Viernes": "K22:K36"},
    3: {"Lunes": "C39:C53", "Martes": "E39:E53", "Miércoles": "G39:G53", "Jueves": "I39:I53", "Viernes": "K39:K53"},
    4: {"Lunes": "C56:C70", "Martes": "E56:E70", "Miércoles": "G56:G70", "Jueves": "I56:I70", "Viernes": "K56:K70"},
    5: {"Lunes": "C74:C88", "Martes": "E74:E88", "Miércoles": "G74:G88", "Jueves": "I74:I88", "Viernes": "K74:K88"},
}

def escribir_tarea(ws, semana, dia_semana, texto, mapa_tareas):
    if semana not in mapa_tareas or dia_semana not in mapa_tareas[semana]:
        return
    rango = mapa_tareas[semana][dia_semana]
    min_col, min_row, max_col, max_row = range_boundaries(rango)
    for row in range(min_row, max_row + 1):
        cel = ws.cell(row=row, column=min_col)
        if cel.value is None:
            cel.value = texto
            return

def llenar_tareas_calendario(wb, nombre_hoja, df, año, mes, mapa_tareas, persona_substr):
    """Escribe las tareas en la hoja correspondiente a la persona.
       persona_substr debe ser en minúsculas (ej. 'manuela') y se compara por substring con los campos de responsables.
    """
    if nombre_hoja not in wb.sheetnames:
        print(f"⚠️ La hoja '{nombre_hoja}' no existe (saltando).")
        return
    ws = wb[nombre_hoja]
    dias_semana = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]

    for _, row in df.iterrows():
        # Normalizar responsables como string en minúsculas
        resp_prod = str(row.get(name_column_responsable_produccion, "")).strip().lower()
        resp_pub = str(row.get(name_column_responsable_publicacion, "")).strip().lower()

        # PRODUCCIÓN: si la persona aparece en el campo responsable de producción
        fecha_prod = row.get(name_column_fecha_produccion)
        if pd.notna(fecha_prod) and isinstance(fecha_prod, date) and persona_substr in resp_prod:
            if fecha_prod.month == mes and fecha_prod.year == año:
                semana = semana_por_domingos(fecha_prod)
                if semana > 5:
                    # Si por cualquier razón excede 5, lo colocamos en la última (5)
                    semana = 5
                weekday = fecha_prod.weekday()
                if weekday < 5:
                    dia_semana = dias_semana[weekday]
                    texto = (
                        f"Hacer - {row.get(name_column_blog_plataforma, '')} - {row.get(name_column_partner, '')} - "
                        f"{row.get(name_column_tiempo_estimado_produc, '')} \n {row.get(name_column_tema_blog, '')}  \n"
                        f"Responsable hacer: {row.get(name_column_responsable_produccion, '')} \n Responsable publicar: {row.get(name_column_responsable_publicacion, '')}"
                    )
                    escribir_tarea(ws, semana, dia_semana, texto, mapa_tareas)

        # PUBLICACIÓN: si la persona aparece en el campo responsable de publicación
        fecha_pub = row.get(name_column_fecha_publicacion)
        if pd.notna(fecha_pub) and isinstance(fecha_pub, date) and persona_substr in resp_pub:
            if fecha_pub.month == mes and fecha_pub.year == año:
                semana = semana_por_domingos(fecha_pub)
                if semana > 5:
                    semana = 5
                weekday = fecha_pub.weekday()
                if weekday < 5:
                    dia_semana = dias_semana[weekday]
                    texto = (
                        f"Publicar - {row.get(name_column_blog_plataforma, '')} - {row.get(name_column_partner, '')} - "
                        f"{row.get(name_column_tiempo_estimado_public, '')} \n {row.get(name_column_tema_blog, '')} \n"
                        f"Responsable hacer: {row.get(name_column_responsable_produccion, '')} \n Responsable publicar: {row.get(name_column_responsable_publicacion, '')}"
                    )
                    escribir_tarea(ws, semana, dia_semana, texto, mapa_tareas)

# Aplicar a todas las hojas/personas detectadas
for persona_substr, hoja_name in hojas_persona.items():
    llenar_tareas_calendario(wb, hoja_name, df_mes, año, mes, mapa_tareas, persona_substr)
    print(f"✅ Se llenaron tareas para '{persona_substr}' en hoja '{hoja_name}'")

# Guardar resultado
wb.save(nombre_salida)
print(f"\n✅ Calendario de {nombre_mes} guardado en '{nombre_salida}'")
