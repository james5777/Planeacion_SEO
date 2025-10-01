# ----------------------------------------------------------------------
# Autor: James Palacio
#        Jacobo Chica
#
# Fecha: 09/09/2025
#
# Descripción: Integración de planeación SEO con base de datos y calendario
# ----------------------------------------------------------------------

### ---------- Librerias importadas ---------- ###
import pandas as pd
import sqlite3
from pathlib import Path
from openpyxl import load_workbook
import calendar
from datetime import date
from openpyxl.utils import range_boundaries
from openpyxl.styles import PatternFill
from openpyxl.styles import Font
import re
import os
import sys


### ---------- Parámetros globales ---------- ###
año = int(os.getenv("AÑO", "2025"))
mes = int(os.getenv("MES", "9"))  # Seleccionar el mes (1-12)

def get_resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def run_main():

    ### ---------- Parámetros globales ---------- ###
    año = int(os.getenv("AÑO", "2025"))
    mes = int(os.getenv("MES", "9"))  # Seleccionar el mes (1-12)

    carpeta_base = Path(os.getenv("RUTA_CARPETA", ""))
    carpeta = carpeta_base / "Archivos" / "archivos_origen"

    ### ---------- Validaciones ---------- ###
    ### ---------- Verificar que la carpeta existe y contiene archivos ---------- ###
    if not carpeta.is_dir():
        raise FileNotFoundError(f"la carpeta de origen '{carpeta}' no se encontró. verifique la ruta seleccionada")

    archivos = list(carpeta.glob("*.xlsx"))
    if not archivos:
        raise FileNotFoundError(f"la carpeta de origen '{carpeta}' no contiene archivos. verifique la ruta seleccionada")
    

    


    ### ---------- Ruta de la base de datos ---------- ###
    rutadb = carpeta_base / "Archivos" / "Archivos_base_de_datos" / "Archivo_base_de_datos.db"

    ### ---------- Nombre de la tabla ---------- ###
    name_tabla_general = "Datos_generales"

    ### ---------- Nombres de las Columnas ---------- ###
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

    ### ---------- Hojas que no se van a recorrer ---------- ###
    hojas_a_ignorar = ['Contenido futuro', 'Rendimiento contenido']

    ### ---------- Bucle que recorre todos los archivos excel y sus hojas ---------- ###
    for archivo in carpeta.glob("*.xlsx"):
        try:
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

               
                ### ---------- Columnas que vamos a utilizar en el dataframe generalizado ---------- ###
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

                # #Validacion de formato fecha
                # for col in [name_column_fecha_produccion, name_column_fecha_publicacion]:
                #     df_col_necesarias = convertir_columna_fecha(df_col_necesarias, col)
                #     if df_col_necesarias[col].isna().any():
                #         raise ValueError(f"⚠️ Advertencia: En el archivo '{archivo.name}', hoja '{hoja}', columna '{col}' hay fechas inválidas o mal formateadas.")

                dataframes.append(df)
            xls.close()
        except PermissionError:
            raise PermissionError(f"El archivo '{archivo.name}' está abierto. por favor, cierrelo e intente de nuevo.")
        except Exception as e:
            raise RuntimeError(f"Ocurrió un error al procesar el archivo '{archivo.name}': {e}")
        
        

    df_generalizado = pd.concat(dataframes, ignore_index=True)
    df_generalizado.columns = df_generalizado.columns.str.strip()

    ### ---------- Mapaeo de nombres de partners en relacion a los archivos ---------- ###
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
    ### ---------- Se agrega una nueva columna con los nombres de los parners ---------- ###
    df_generalizado[name_column_partner] = df_generalizado[name_column_archivo_origen].map(mapeo_partners)

    # Filtrar columnas que usaremos
    columnas_necesarias = [
        name_column_archivo_origen, name_column_partner, name_column_blog_plataforma, name_column_tema_blog,
        name_column_fecha_produccion, name_column_responsable_produccion,
        name_column_tiempo_estimado_produc, name_column_fecha_publicacion,
        name_column_responsable_publicacion, name_column_tiempo_estimado_public
    ]

    ### ---------- Se crea un dataframe nuevo con la lista de columnas que vamos a utilizar ---------- ###
    df_col_necesarias = df_generalizado[columnas_necesarias].copy()
    df_col_necesarias.columns = df_col_necesarias.columns.str.strip()

    ### ---------- Normalizacion de fechas, se cambia de Serial de excel a yyyy/mm/dd ---------- ###
    def convertir_columna_fecha(df, col):
        if col not in df.columns:
            return df
        if pd.api.types.is_numeric_dtype(df[col]):
            df[col] = pd.to_datetime(df[col], unit="d", origin="1899-12-30", errors="coerce")
        else:
            df[col] = pd.to_datetime(df[col], errors="coerce")
        return df

    ### ---------- Se convierte la columna a tipo date ---------- ###
    for col in [name_column_fecha_produccion, name_column_fecha_publicacion]:
        df_col_necesarias = convertir_columna_fecha(df_col_necesarias, col)
        df_col_necesarias[col] = pd.to_datetime(df_col_necesarias[col], errors="coerce").dt.date

    ### ---------- Se convierte los minutos escritos de diferentes formas a numero ---------- ###
    def convertir_a_minutos(valor):
        if pd.isna(valor):
            return 0

        # Normalizar a string para trabajar
        s = str(valor).strip().lower().replace(",", ".")  # "30,0" -> "30.0"

        # Si es número puro (int o float en string)
        if re.fullmatch(r"\d+(\.\d+)?", s):
            return int(float(s))  # 60.0 -> 60

        # Buscar horas (ej: "1 hora", "2h", "1.5 horas")
        horas_match = re.search(r"(\d+(\.\d+)?)\s*(h|hora|horas)", s)
        minutos_match = re.search(r"(\d+(\.\d+)?)\s*(m|min|minuto|minutos)", s)

        total = 0
        if horas_match:
            total += int(float(horas_match.group(1)) * 60)
        if minutos_match:
            total += int(float(minutos_match.group(1)))

        return total

    # ----------Se aplica la funcion de "Convertir a minutos" para normalizar columnas de tiempos ----------
    for col in [name_column_tiempo_estimado_produc, name_column_tiempo_estimado_public]:
        if col in df_col_necesarias.columns:
            df_col_necesarias[col] = df_col_necesarias[col].apply(convertir_a_minutos).fillna(0).astype(int)

    # Guardar en sqlite
    def guardar_en_sqlite(df: pd.DataFrame, nombre_tabla: str, ruta_db: Path, if_exists: str = "replace") -> None:
        if df.empty:
            print(f"\n ⚠️ El DataFrame está vacío. No se insertaron datos en la tabla '{nombre_tabla}'.\n ")
            return
        with sqlite3.connect(ruta_db) as conn:
            df.to_sql(nombre_tabla, conn, if_exists=if_exists, index=False)
        print(f"\n ✅ Se insertaron los datos en la tabla: '{nombre_tabla}' en la base de datos '{ruta_db.name}'.\n ")

    guardar_en_sqlite(df_col_necesarias, name_tabla_general, rutadb)

    # ------------------ Diccionario de mapeo Mes - numero mes ------------------
    meses_es = {
        1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
        5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
        9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
    }

    nombre_mes = meses_es[mes]

    ### ---------- Nombre de archivo de plantilla ---------- ###
    nombre_plantilla_base = get_resource_path("plantilla.xlsx")
    nombre_salida_x = Path(f"Calendarios obtenidos/Calendario {nombre_mes}.xlsx")
    nombre_plantilla_x = nombre_plantilla_base


    ### ---------- Conexion con SQLite para traer columnas ---------- ###
    # Traer tareas de mes actual y mes anterior desde SQLite
    mes_str = f"{año}-{mes:02d}"
    mes_ant = mes - 1 if mes > 1 else 12
    año_ant = año if mes > 1 else año - 1
    mes_ant_str = f"{año_ant}-{mes_ant:02d}"

    with sqlite3.connect(rutadb) as conn:
        query = f"""
        SELECT *
        FROM "{name_tabla_general}"
        WHERE STRFTIME('%Y-%m', "Fecha de producción") IN ('{mes_str}', '{mes_ant_str}')
        OR STRFTIME('%Y-%m', "Fecha de publicación") IN ('{mes_str}', '{mes_ant_str}');
        """
        df_mes = pd.read_sql_query(query, conn)


    ### ---------- convertir fechas a date por si acaso (si vienen como strings) ---------- ###
    for col in [name_column_fecha_produccion, name_column_fecha_publicacion]:
        if col in df_mes.columns:
            df_mes[col] = pd.to_datetime(df_mes[col], errors="coerce").dt.date

    if df_mes.empty:
        print(f"⚠️ No se encontraron tareas para {nombre_mes}.")
    else:
        print(f"✅ Se encontraron {len(df_mes)} tareas para el mes {nombre_mes}.")

    ### ---------- Diccionario de encabezados para calendario dinamico ---------- ###
    mapa_encabezados = {
        1: {"Lunes": "C2", "Martes": "E2", "Miércoles": "G2", "Jueves": "I2", "Viernes": "K2"},
        2: {"Lunes": "C21", "Martes": "E21", "Miércoles": "G21", "Jueves": "I21", "Viernes": "K21"},
        3: {"Lunes": "C38", "Martes": "E38", "Miércoles": "G38", "Jueves": "I38", "Viernes": "K38"},
        4: {"Lunes": "C55", "Martes": "E55", "Miércoles": "G55", "Jueves": "I55", "Viernes": "K55"},
        5: {"Lunes": "C72", "Martes": "E72", "Miércoles": "G72", "Jueves": "I72", "Viernes": "K72"},
        6: {"Lunes": "C90", "Martes": "E90", "Miércoles": "G90", "Jueves": "I90", "Viernes": "K90"}
    }

    ### ---------- Se define estilo para color de alerta por carga laboral ---------- ###
    fill_rojo = PatternFill(start_color="FD5D5D", end_color="FD5D5D", fill_type="solid")

    ### ---------- Función para llenar los encabezados de forma dinamica con la fecha - mes correspondiente ---------- ###
    def llenar_encabezados_calendario(wb, nombre_hoja, año, mes, mapa_encabezados, df, persona_substr):
        if nombre_hoja not in wb.sheetnames:
            return

        ws = wb[nombre_hoja]
        dias_semana = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
        num_dias = calendar.monthrange(año, mes)[1]
        semana = 1

        for dia in range(1, num_dias + 1):
            f = date(año, mes, dia)
            weekday = f.weekday()

            if weekday < 5:  # solo lunes a viernes
                nombre_dia = dias_semana[weekday]

                # ---- FILTRAR SOLO TAREAS DE ESA PERSONA ----
                mask_prod = (df[name_column_fecha_produccion] == f) & \
                            (df[name_column_responsable_produccion].str.lower().str.contains(persona_substr))
                mask_pub  = (df[name_column_fecha_publicacion] == f) & \
                            (df[name_column_responsable_publicacion].str.lower().str.contains(persona_substr))

                total_minutos = 0
                if mask_prod.any():
                    total_minutos += int(df.loc[mask_prod, name_column_tiempo_estimado_produc].sum())
                if mask_pub.any():
                    total_minutos += int(df.loc[mask_pub, name_column_tiempo_estimado_public].sum())

                # ---- convertir a horas y minutos ----
                horas, minutos = divmod(total_minutos, 60)

                if horas > 0 and minutos > 0:
                    tiempo_str = f"{horas}h {minutos}m"
                elif horas > 0:
                    tiempo_str = f"{horas}h"
                elif minutos > 0:
                    tiempo_str = f"{minutos}m"
                else:
                    tiempo_str = "0h"

                # ---- escribir en el calendario ----
                if semana in mapa_encabezados and nombre_dia in mapa_encabezados[semana]:
                    celda = ws[mapa_encabezados[semana][nombre_dia]]
                    celda.value = f"{nombre_dia} {dia} ({tiempo_str})"

                    # Si pasa de 8 horas, pintamos en rojo
                    if total_minutos > 480:
                        celda.fill = fill_rojo

            # domingo -> siguiente semana
            if weekday == 6:
                semana += 1

    # abrir plantilla
    wb = load_workbook(nombre_plantilla_base)

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
    for persona_substr, hoja in hojas_persona.items():
        llenar_encabezados_calendario(wb, hoja, año, mes, mapa_encabezados, df_mes, persona_substr)


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
        6: {"Lunes": "C92:C106", "Martes": "E92:E106", "Miércoles": "G92:G106", "Jueves": "I92:I106", "Viernes": "K92:K106"},
    }
    ### ---------- Colores de fondo para cada tarea en relacion con cada partner ---------- ###
    partner_fills = {
        "ACIERTALA": PatternFill(start_color="D0D0F7", end_color="D0D0F7", fill_type="solid"),  # rojo claro
        "CAMANBET": PatternFill(start_color="E2FECD", end_color="E2FECD", fill_type="solid"),      # azul claro
        "DORADOBET CR": PatternFill(start_color="F8D3D3", end_color="F8D3D3", fill_type="solid"),     # verde claro#A8E6FA
        "DORADOBET GT": PatternFill(start_color="F8D3D3", end_color="F8D3D3", fill_type="solid"),     # morado claro#EED247
        "DORADOBET PE": PatternFill(start_color="F8D3D3", end_color="F8D3D3", fill_type="solid"),     # morado claro
        "DORADOBET SV": PatternFill(start_color="F8D3D3", end_color="F8D3D3", fill_type="solid"),     # morado claro
        "ECUABET": PatternFill(start_color="FFED8C", end_color="FFED8C", fill_type="solid"),     # morado claro
        "GANAPLAY GT": PatternFill(start_color="C9C9C9", end_color="C9C9C9", fill_type="solid"),     # morado claro
        "GANAPLAY SV": PatternFill(start_color="C9C9C9", end_color="C9C9C9", fill_type="solid"),     # morado claro
        "PANIPLAY": PatternFill(start_color="ABFBFE", end_color="ABFBFE", fill_type="solid"),     # morado claro
    }


    ### ---------- Funcion para escribir tarea en cada una de las celdas seleccionadas para cada dia ---------- ###
    def escribir_tarea(ws, semana, dia_semana, texto, mapa_tareas, partner):
        if semana not in mapa_tareas or dia_semana not in mapa_tareas[semana]:
            return
        rango = mapa_tareas[semana][dia_semana]
        min_col, min_row, max_col, max_row = range_boundaries(rango)
        for row in range(min_row, max_row + 1):
            cel = ws.cell(row=row, column=min_col)
            if cel.value is None:
                cel.value = texto
                # Aplicar color según partner
                partner_key = partner.strip().upper()
                if partner_key in partner_fills:
                    cel.fill = partner_fills[partner_key]
                return

    ### ---------- Funcion para escribir tarea en base a la plantilla dinamica dependiendo del mes ---------- ###
    def llenar_tareas_calendario(wb, nombre_hoja, df, año, mes, mapa_tareas, persona_substr):
        if nombre_hoja not in wb.sheetnames:
            print(f"⚠️ La hoja '{nombre_hoja}' no existe (saltando).")
            return
        ws = wb[nombre_hoja]
        dias_semana = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]

        for _, row in df.iterrows():
            resp_prod = str(row.get(name_column_responsable_produccion, "")).strip().lower()
            resp_pub = str(row.get(name_column_responsable_publicacion, "")).strip().lower()
            partner = str(row.get(name_column_partner, "")).strip().upper()

            # PRODUCCIÓN
            fecha_prod = row.get(name_column_fecha_produccion)
            if pd.notna(fecha_prod) and isinstance(fecha_prod, date) and persona_substr in resp_prod:
                if fecha_prod.month == mes and fecha_prod.year == año:
                    semana = semana_por_domingos(fecha_prod)
                    if semana > 5:
                        semana = 5
                    weekday = fecha_prod.weekday()
                    if weekday < 5:
                        dia_semana = dias_semana[weekday]
                        texto = (
                            f"Hacer - {row.get(name_column_blog_plataforma, '')} - {partner} - "
                            f"{row.get(name_column_tiempo_estimado_produc, '')} \n {row.get(name_column_tema_blog, '')}  \n"
                            f"Responsable hacer: {row.get(name_column_responsable_produccion, '')} \n Responsable publicar: {row.get(name_column_responsable_publicacion, '')}"
                        )
                        escribir_tarea(ws, semana, dia_semana, texto, mapa_tareas, partner)

            # PUBLICACIÓN
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
                            f"Publicar - {row.get(name_column_blog_plataforma, '')} - {partner} - "
                            f"{row.get(name_column_tiempo_estimado_public, '')} \n {row.get(name_column_tema_blog, '')} \n"
                            f"Responsable hacer: {row.get(name_column_responsable_produccion, '')} \n Responsable publicar: {row.get(name_column_responsable_publicacion, '')}"
                        )
                        escribir_tarea(ws, semana, dia_semana, texto, mapa_tareas, partner)

    # ---------------- GENERACIÓN DE CALENDARIOS PARA TODOS LOS MESES IMPLICADOS ---------------- #

    # Crear carpeta "Calendarios obtenidos" si no existe
    carpeta_salida = Path("Calendarios obtenidos")
    carpeta_salida.mkdir(parents=True, exist_ok=True)

    # ---------------- GENERACIÓN DE CALENDARIOS ---------------- #

    meses_a_generar = set()
    meses_a_generar.add((año, mes))  # mes principal

    # Calcular mes anterior
    mes_ant = mes - 1 if mes > 1 else 12
    año_ant = año if mes > 1 else año - 1
    meses_a_generar.add((año_ant, mes_ant))

    for (año_x, mes_x) in sorted(meses_a_generar):
        nombre_mes_x = meses_es[mes_x]
        nombre_plantilla_x = nombre_plantilla_base
        nombre_salida_x = Path(f"Calendarios obtenidos/Calendario {nombre_mes_x}.xlsx")

        # Normalizar fechas a datetime64 para filtrados seguros
        fechas_prod = pd.to_datetime(df_mes[name_column_fecha_produccion], errors="coerce")
        fechas_pub = pd.to_datetime(df_mes[name_column_fecha_publicacion], errors="coerce")

        if (año_x, mes_x) == (año, mes):
            # Mes principal → todas las tareas que caen en este mes
            df_mes_x = df_mes[
                ((fechas_prod.dt.month == mes) & (fechas_prod.dt.year == año)) |
                ((fechas_pub.dt.month == mes) & (fechas_pub.dt.year == año))
            ]
        else:
        # Mes anterior →
        # a) producciones en mes anterior con publicación en mes actual
        # b) publicaciones reales que caen en mes anterior
        # c) tareas totalmente dentro del mes anterior (prod y pub en mes anterior)

            fechas_prod = pd.to_datetime(df_mes[name_column_fecha_produccion], errors="coerce")
            fechas_pub  = pd.to_datetime(df_mes[name_column_fecha_publicacion], errors="coerce")

            df_mes_x = df_mes[
                (
                    (fechas_prod.dt.month == mes_x) & (fechas_prod.dt.year == año_x) &
                    (fechas_pub.dt.month == mes) & (fechas_pub.dt.year == año)
                )
                |
                (
                    (fechas_pub.dt.month == mes_x) & (fechas_pub.dt.year == año_x)
                )
                |
                (
                    (fechas_prod.dt.month == mes_x) & (fechas_prod.dt.year == año_x) &
                    (fechas_pub.dt.month == mes_x) & (fechas_pub.dt.year == año_x)
                )
            ]
        

        if df_mes_x.empty:
            raise ValueError(f"No se encontraron tareas para {mes_x}/{año_x}.")


            print(f"⚠️ No se encontraron tareas para {nombre_mes_x}.")
            continue

        print(f"\n✅ Generando calendario para {nombre_mes_x} con {len(df_mes_x)} tareas...")

        wb = load_workbook(nombre_plantilla_x)

        # Detectar hojas por persona
        hojas_persona = {}
        for p in personas:
            encontrada = next((s for s in wb.sheetnames if p in s.lower()), None)
            if encontrada:
                hojas_persona[p] = encontrada
            else:
                print(f"⚠️ No se encontró hoja para '{p}' en la plantilla {nombre_plantilla_x} (buscando substring).")

        # Llenar encabezados y tareas
        for persona_substr, hoja_name in hojas_persona.items():
            llenar_encabezados_calendario(wb, hoja_name, año_x, mes_x, mapa_encabezados, df_mes_x, persona_substr)
            llenar_tareas_calendario(wb, hoja_name, df_mes_x, año_x, mes_x, mapa_tareas, persona_substr)
            print(f"✅ Se llenaron tareas para '{persona_substr}' en hoja '{hoja_name}' ({nombre_mes_x})")

        wb.save(nombre_salida_x)
        print(f"📂 Archivo guardado: {nombre_salida_x}")

if __name__ == "__main__":
    run_main()

