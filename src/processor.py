import datetime
import os
import re
from typing import Dict, List, Optional, Tuple

import pandas as pd
import xlrd

from src.config import OUTPUT_COLUMNS, SUMMARY_COLUMNS


def cargar_archivo_crudo(filepath: str) -> pd.DataFrame:
    """
    Lee de forma ultra-robusta archivos exportados por Leal Sistemas:
    1. Libros Excel antiguos BIFF2-BIFF8 (.xls) vía xlrd con soporte de codificación cp1252/latin-1
       y libros modernos (.xlsx) vía openpyxl.
    2. Tablas HTML guardadas con extensión .xls.
    3. Archivos delimitados CSV / TSV con autodetección de separador y codificación.
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"El archivo no existe: {filepath}")

    errores: List[str] = []

    # 1A. Intento directo con xlrd (ideal para archivos BIFF antiguos de Leal Sistemas sin CODEPAGE)
    for enc_override in ("cp1252", "latin-1", "utf-8"):
        try:
            with open(os.devnull, "w", encoding="utf-8") as devnull:
                wb = xlrd.open_workbook(
                    filepath,
                    encoding_override=enc_override,
                    logfile=devnull,
                )
                sheet = wb.sheet_by_index(0)
                if sheet.nrows > 0 and sheet.ncols >= 3:
                    filas = []
                    for r_idx in range(sheet.nrows):
                        fila_vals = []
                        for c_idx in range(sheet.ncols):
                            cell = sheet.cell(r_idx, c_idx)
                            val = cell.value
                            if cell.ctype == xlrd.XL_CELL_DATE:
                                try:
                                    dt_tuple = xlrd.xldate_as_tuple(val, wb.datemode)
                                    if dt_tuple[0] > 0:
                                        val = datetime.date(
                                            dt_tuple[0], dt_tuple[1], dt_tuple[2]
                                        ).strftime("%d/%m/%Y")
                                except Exception:
                                    val = str(val)
                            elif cell.ctype == xlrd.XL_CELL_NUMBER:
                                if float(val).is_integer():
                                    val = str(int(val))
                                else:
                                    val = str(val)
                            fila_vals.append(val)
                        filas.append(fila_vals)
                    df_xlrd = pd.DataFrame(filas)
                    if not df_xlrd.empty and df_xlrd.shape[1] >= 3:
                        return df_xlrd
        except Exception as exc:
            errores.append(f"xlrd({enc_override}): {exc}")

    # 1B. Intento con pd.read_excel (openpyxl / default)
    for engine in (None, "openpyxl", "xlrd"):
        try:
            kwargs = {"header": None}
            if engine is not None:
                kwargs["engine"] = engine
            df = pd.read_excel(filepath, **kwargs)
            if df is not None and not df.empty and df.shape[1] >= 3:
                return df
        except Exception as exc:
            errores.append(f"read_excel({engine}): {exc}")

    # 2. Intento con pd.read_html (para falsos .xls que son tablas HTML)
    for encoding in ("utf-8", "latin-1", "cp1252"):
        try:
            tablas = pd.read_html(filepath, encoding=encoding)
            if tablas:
                df_html = max(tablas, key=lambda t: t.shape[0] * t.shape[1])
                if not df_html.empty and df_html.shape[1] >= 3:
                    if not all(isinstance(col, int) for col in df_html.columns):
                        header_row = pd.DataFrame([list(df_html.columns)])
                        df_html.columns = range(df_html.shape[1])
                        df_html = pd.concat(
                            [header_row, df_html], ignore_index=True
                        )
                    return df_html
        except Exception as exc:
            errores.append(f"read_html({encoding}): {exc}")

    # 3. Intento con pd.read_csv (autodetección de separadores y fallback a latin-1)
    for encoding in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        for sep in (None, "\t", ";", ","):
            try:
                df_csv = pd.read_csv(
                    filepath,
                    sep=sep,
                    engine="python",
                    header=None,
                    encoding=encoding,
                    dtype=str,
                )
                if df_csv is not None and not df_csv.empty and df_csv.shape[1] >= 3:
                    return df_csv
            except Exception as exc:
                errores.append(f"read_csv({encoding}, sep={sep}): {exc}")

    detalle = "; ".join(errores[:4])
    raise ValueError(
        "No fue posible interpretar el formato del archivo seleccionado. "
        f"Verifique que sea un reporte válido de Leal Sistemas. Detalle: {detalle}"
    )


def limpiar_id(valor: object) -> str:
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if not texto or texto.lower() == "nan":
        return ""
    if re.fullmatch(r"\d+\.0+", texto):
        texto = texto.split(".")[0]
    return texto


def normalizar_fecha(valor: object) -> str:
    if pd.isna(valor):
        return ""

    if isinstance(valor, (datetime.datetime, datetime.date, pd.Timestamp)):
        return valor.strftime("%d/%m/%Y")

    texto = str(valor).strip()
    if not texto or texto.lower() in ("nan", "nat", "none"):
        return ""

    # Eliminar sufijo 00:00:00 si viene incluido
    texto = re.sub(r"\s+\d{1,2}:\d{2}(:\d{2})?(\.\d+)?$", "", texto).strip()

    # Patrón YYYY-MM-DD o YYYY/MM/DD
    m_iso = re.fullmatch(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", texto)
    if m_iso:
        anio, mes, dia = (
            int(m_iso.group(1)),
            int(m_iso.group(2)),
            int(m_iso.group(3)),
        )
        try:
            return datetime.date(anio, mes, dia).strftime("%d/%m/%Y")
        except ValueError:
            return f"{dia:02d}/{mes:02d}/{anio:04d}"

    # Patrón DD/MM/YYYY o DD-MM-YYYY o D/M/YYYY
    m_lat = re.fullmatch(r"(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})", texto)
    if m_lat:
        p1, p2, anio_raw = (
            int(m_lat.group(1)),
            int(m_lat.group(2)),
            int(m_lat.group(3)),
        )
        anio = 2000 + anio_raw if anio_raw < 100 else anio_raw
        if p1 <= 12 and p2 > 12:
            mes, dia = p1, p2
        else:
            dia, mes = p1, p2
        try:
            return datetime.date(anio, mes, dia).strftime("%d/%m/%Y")
        except ValueError:
            return f"{dia:02d}/{mes:02d}/{anio:04d}"

    return texto


def es_fila_encabezado_o_vacia(
    id_str: str, nombre_str: str, fecha_str: str
) -> bool:
    if not id_str and not nombre_str and not fecha_str:
        return True

    palabras_encabezado = {
        "nro. id",
        "nro id",
        "nro.id",
        "id",
        "legajo",
        "nro",
        "numero",
        "número",
        "codigo",
        "código",
        "empleado",
    }
    if id_str.lower() in palabras_encabezado:
        return True
    if nombre_str.lower() in {"nombre", "apellido y nombre", "empleado", "personal"}:
        return True
    if fecha_str.lower() in {"fecha", "dia", "día", "date", "hora"}:
        return True

    tiene_id_valido = bool(re.search(r"\d", id_str))
    tiene_fecha_valida = bool(
        re.search(r"\d{1,4}[-/]\d{1,2}[-/]\d{1,4}", fecha_str)
    )
    return not (tiene_id_valido or tiene_fecha_valida)


def normalizar_hora_hhmm(hora_raw: str) -> Optional[str]:
    partes = hora_raw.strip().split(":")
    if len(partes) != 2:
        return None
    try:
        hh = int(partes[0])
        mm = int(partes[1])
        if 0 <= hh <= 23 and 0 <= mm <= 59:
            return f"{hh:02d}:{mm:02d}"
    except ValueError:
        return None
    return None


def calcular_minutos_trabajados(ingreso: str, salida: str) -> int:
    h_in, m_in = (int(x) for x in ingreso.split(":"))
    h_out, m_out = (int(x) for x in salida.split(":"))

    minutos_ingreso = h_in * 60 + m_in
    minutos_salida = h_out * 60 + m_out

    diferencia = minutos_salida - minutos_ingreso
    if diferencia < 0:
        diferencia += 24 * 60
    return diferencia


def formatear_minutos_hhmm(total_minutos: int) -> str:
    horas = total_minutos // 60
    minutos = total_minutos % 60
    return f"{horas:02d}:{minutos:02d}"


def procesar_fichadas_df(df_crudo: pd.DataFrame) -> pd.DataFrame:
    """
    Transforma el DataFrame crudo en el DataFrame limpio con las 6 columnas exactas:
    Nro. ID, Nombre, Fecha, Ingreso, Salida, Tiempo Transcurrido.
    """
    if df_crudo is None or df_crudo.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    df = df_crudo.copy()
    df.columns = range(df.shape[1])
    while df.shape[1] < 4:
        df[df.shape[1]] = ""

    patron_horas = re.compile(r"\b\d{1,2}:\d{2}\b")
    registros_limpios = []

    for _, fila in df.iterrows():
        id_limpio = limpiar_id(fila[0])
        nombre_limpio = "" if pd.isna(fila[1]) else str(fila[1]).strip()
        if nombre_limpio.lower() == "nan":
            nombre_limpio = ""
        fecha_limpia = normalizar_fecha(fila[2])

        if es_fila_encabezado_o_vacia(id_limpio, nombre_limpio, fecha_limpia):
            continue

        fichadas_raw = fila[3]
        ingreso = ""
        salida = ""
        tiempo_transcurrido = ""

        if not pd.isna(fichadas_raw):
            texto_fichadas = str(fichadas_raw).strip()
            if texto_fichadas and texto_fichadas.lower() != "nan":
                coincidencias = patron_horas.findall(texto_fichadas)
                horas_normalizadas: List[str] = []
                for h_match in coincidencias:
                    h_norm = normalizar_hora_hhmm(h_match)
                    if h_norm is not None:
                        horas_normalizadas.append(h_norm)

                if horas_normalizadas:
                    primera_hora = horas_normalizadas[0]
                    ultima_hora = horas_normalizadas[-1]
                    ingreso = primera_hora

                    if primera_hora == ultima_hora:
                        salida = "Sin salida"
                        tiempo_transcurrido = ""
                    else:
                        salida = ultima_hora
                        mins = calcular_minutos_trabajados(ingreso, salida)
                        tiempo_transcurrido = formatear_minutos_hhmm(mins)

        registros_limpios.append(
            {
                "Nro. ID": id_limpio,
                "Nombre": nombre_limpio,
                "Fecha": fecha_limpia,
                "Ingreso": ingreso,
                "Salida": salida,
                "Tiempo Transcurrido": tiempo_transcurrido,
            }
        )

    return pd.DataFrame(registros_limpios, columns=OUTPUT_COLUMNS)


def filtrar_dataframe_según_modo(
    df_limpio: pd.DataFrame, modo: str, texto_busqueda: str = ""
) -> pd.DataFrame:
    """
    Filtra y organiza el DataFrame procesado según el modo seleccionado:
    - "con_fichadas": Solo filas donde el empleado fichó (Ingreso != "").
    - "empleados_activos": Todos los días del mes pero solo de empleados que tienen >=1 fichada en el periodo.
    - "todos": El 100% de las filas (ordenando primero los empleados activos para evitar ver filas vacías arriba).
    """
    if df_limpio is None or df_limpio.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    tiene_fichada = df_limpio["Ingreso"].astype(str).str.strip() != ""

    if modo == "con_fichadas":
        resultado = df_limpio[tiene_fichada].copy()
        # Si el archivo de prueba solo tenía filas vacías, mostrar todas
        if resultado.empty:
            resultado = df_limpio.copy()
    elif modo == "empleados_activos":
        ids_activos = set(df_limpio.loc[tiene_fichada, "Nro. ID"].unique())
        if ids_activos:
            resultado = df_limpio[df_limpio["Nro. ID"].isin(ids_activos)].copy()
        else:
            resultado = df_limpio.copy()
    else:
        # Modo "todos": ponemos primero a los empleados activos para que la tabla nunca arranque vacía
        ids_activos = set(df_limpio.loc[tiene_fichada, "Nro. ID"].unique())
        if ids_activos:
            df_activos = df_limpio[df_limpio["Nro. ID"].isin(ids_activos)]
            df_inactivos = df_limpio[~df_limpio["Nro. ID"].isin(ids_activos)]
            resultado = pd.concat([df_activos, df_inactivos], ignore_index=True)
        else:
            resultado = df_limpio.copy()

    busqueda = texto_busqueda.strip().lower()
    if busqueda:
        mask_id = resultado["Nro. ID"].astype(str).str.lower().str.contains(busqueda, na=False)
        mask_nom = resultado["Nombre"].astype(str).str.lower().str.contains(busqueda, na=False)
        mask_fec = resultado["Fecha"].astype(str).str.lower().str.contains(busqueda, na=False)
        resultado = resultado[mask_id | mask_nom | mask_fec].copy()

    return resultado.reset_index(drop=True)


def generar_resumen_por_empleado(df_limpio: pd.DataFrame) -> pd.DataFrame:
    """
    Genera una tabla gerencial de resumen por empleado activo con totales de días fichados,
    jornadas completas, días sin salida y suma total de horas trabajadas.
    """
    if df_limpio is None or df_limpio.empty:
        return pd.DataFrame(columns=SUMMARY_COLUMNS)

    tiene_fichada = df_limpio["Ingreso"].astype(str).str.strip() != ""
    df_activos = df_limpio[tiene_fichada]
    if df_activos.empty:
        return pd.DataFrame(columns=SUMMARY_COLUMNS)

    filas_resumen = []
    for (nro_id, nombre), grupo in df_activos.groupby(
        ["Nro. ID", "Nombre"], sort=False
    ):
        dias_fichados = len(grupo)
        sin_salida = int((grupo["Salida"] == "Sin salida").sum())
        jornadas_completas = dias_fichados - sin_salida

        minutos_totales = 0
        for _, row in grupo.iterrows():
            t_str = str(row["Tiempo Transcurrido"]).strip()
            if ":" in t_str:
                partes = t_str.split(":")
                try:
                    minutos_totales += int(partes[0]) * 60 + int(partes[1])
                except ValueError:
                    minutos_totales += 0

        total_hhmm = formatear_minutos_hhmm(minutos_totales)
        if jornadas_completas > 0:
            promedio_mins = minutos_totales // jornadas_completas
            promedio_hhmm = formatear_minutos_hhmm(promedio_mins)
        else:
            promedio_hhmm = "00:00"

        filas_resumen.append(
            {
                "Nro. ID": nro_id,
                "Nombre": nombre,
                "Días con Fichaje": dias_fichados,
                "Jornadas Completas": jornadas_completas,
                "Días Sin Salida": sin_salida,
                "Total Horas (HH:MM)": total_hhmm,
                "Promedio Diario (HH:MM)": promedio_hhmm,
            }
        )

    return pd.DataFrame(filas_resumen, columns=SUMMARY_COLUMNS)


def obtener_estadisticas_df(df_limpio: pd.DataFrame) -> Dict[str, int]:
    if df_limpio is None or df_limpio.empty:
        return {
            "total_filas": 0,
            "dias_con_fichadas": 0,
            "jornadas_completas": 0,
            "sin_salida": 0,
            "empleados_totales": 0,
            "empleados_activos": 0,
            "filas_empleados_activos": 0,
        }

    tiene_fichada = df_limpio["Ingreso"].astype(str).str.strip() != ""
    dias_con_fichadas = int(tiene_fichada.sum())
    sin_salida = int((df_limpio["Salida"] == "Sin salida").sum())
    jornadas_completas = max(0, dias_con_fichadas - sin_salida)

    empleados_totales = int(df_limpio["Nro. ID"].nunique())
    ids_activos = set(df_limpio.loc[tiene_fichada, "Nro. ID"].unique())
    empleados_activos = len(ids_activos)
    filas_empleados_activos = int(df_limpio["Nro. ID"].isin(ids_activos).sum())

    return {
        "total_filas": len(df_limpio),
        "dias_con_fichadas": dias_con_fichadas,
        "jornadas_completas": jornadas_completas,
        "sin_salida": sin_salida,
        "empleados_totales": empleados_totales,
        "empleados_activos": empleados_activos,
        "filas_empleados_activos": filas_empleados_activos,
    }
