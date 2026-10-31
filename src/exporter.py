import datetime
from typing import Dict, List, Optional, Set
import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
import pandas as pd

from src.config import FONT_FAMILY, OUTPUT_COLUMNS, SUMMARY_COLUMNS
from src.processor import (
    filtrar_dataframe_según_modo,
    generar_resumen_por_empleado,
)


def _parsear_hora_excel(valor: object) -> object:
    """
    Convierte cadenas 'HH:MM' en objetos datetime.time nativos de Excel para que
    las celdas de Ingreso y Salida sean editables y calculables con fórmulas nativas.
    """
    if valor is None or pd.isna(valor):
        return None
    texto = str(valor).strip()
    if not texto or texto.lower() == "nan":
        return None
    if texto == "Sin salida":
        return "Sin salida"
    partes = texto.split(":")
    if len(partes) == 2:
        try:
            hh = int(partes[0])
            mm = int(partes[1])
            if 0 <= hh <= 23 and 0 <= mm <= 59:
                return datetime.time(hh, mm)
        except ValueError:
            pass
    return texto


def _estilizar_hoja_fichadas_dinamica(
    ws: Worksheet,
    df_datos: pd.DataFrame,
    columnas: List[str],
    columnas_centradas: Set[int],
) -> None:
    """
    Genera una hoja de fichadas (Hoja 1 o Hoja 3) con:
    - Horas nativas editables en Ingreso (D) y Salida (E) con formato 'hh:mm' centrado.
    - Fórmula automática en vivo en Tiempo Transcurrido (F) que recalcula al instante
      cuando el usuario escribe o corrige un Ingreso o Salida en Excel.
    - Formato Condicional dinámico que ilumina en ámbar cuando falta la salida y se
      quita automáticamente apenas el usuario escribe la hora de salida.
    """
    header_fill = PatternFill(
        start_color="112239", end_color="112239", fill_type="solid"
    )
    even_row_fill = PatternFill(
        start_color="F8FAFC", end_color="F8FAFC", fill_type="solid"
    )
    odd_row_fill = PatternFill(
        start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"
    )
    warn_row_fill = PatternFill(
        start_color="FFFBEB", end_color="FFFBEB", fill_type="solid"
    )

    header_font = Font(name=FONT_FAMILY, size=11, bold=True, color="FFFFFF")
    cell_font = Font(name=FONT_FAMILY, size=10, bold=False, color="1E293B")
    warn_font = Font(name=FONT_FAMILY, size=10, bold=True, color="B45309")

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    thin_side = Side(border_style="thin", color="CBD5E1")
    cell_border = Border(
        left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
    )

    # 1. Encabezados en fila 1
    ws.append(columnas)
    ws.row_dimensions[1].height = 28

    for col_idx in range(1, len(columnas) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center
        cell.border = cell_border

    ws.freeze_panes = "A2"

    # 2. Filas de datos con fórmula dinámica en columna F (Tiempo Transcurrido)
    for row_idx, row_data in enumerate(
        df_datos[columnas].itertuples(index=False), start=2
    ):
        valores_raw = list(row_data)
        nro_id = valores_raw[0]
        nombre = valores_raw[1]
        fecha = valores_raw[2]
        ingreso_val = _parsear_hora_excel(valores_raw[3])
        salida_val = _parsear_hora_excel(valores_raw[4])

        # Fórmula Excel en vivo: funciona tanto con horas nativas como con texto "HH:MM"
        formula_tiempo = (
            f'=IF(OR(D{row_idx}="",E{row_idx}="",E{row_idx}="Sin salida"),"",'
            f'IFERROR(MOD((E{row_idx}+0)-(D{row_idx}+0),1),""))'
        )

        ws.append([nro_id, nombre, fecha, ingreso_val, salida_val, formula_tiempo])
        ws.row_dimensions[row_idx].height = 21

        row_fill = even_row_fill if row_idx % 2 == 0 else odd_row_fill

        for col_idx in range(1, len(columnas) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.fill = row_fill
            cell.border = cell_border
            cell.font = cell_font

            if col_idx in columnas_centradas:
                cell.alignment = align_center
            else:
                cell.alignment = align_left

            # Formato de hora para Ingreso (D), Salida (E) y Tiempo Transcurrido (F)
            if col_idx in (4, 5, 6):
                cell.number_format = "hh:mm"

    max_r = max(2, ws.max_row)

    # 3. Formato Condicional dinámico para filas "Sin salida" o con Ingreso pero sin Salida
    if ws.max_row >= 2:
        regla_fila_ambar = FormulaRule(
            formula=['OR($E2="Sin salida",AND($D2<>"",$E2=""))'],
            stopIfTrue=False,
            fill=warn_row_fill,
        )
        regla_texto_ambar = FormulaRule(
            formula=['OR($E2="Sin salida",AND($D2<>"",$E2=""))'],
            stopIfTrue=False,
            fill=warn_row_fill,
            font=warn_font,
        )
        ws.conditional_formatting.add(f"A2:F{max_r}", regla_fila_ambar)
        ws.conditional_formatting.add(f"E2:E{max_r}", regla_texto_ambar)

    # 4. Activar filtro automático de Excel
    if ws.max_row >= 1 and len(columnas) >= 1:
        ultima_col_letra = get_column_letter(len(columnas))
        ws.auto_filter.ref = f"A1:{ultima_col_letra}{max(1, ws.max_row)}"

    # 5. Auto-ajuste de ancho de columnas
    anchos_minimos = {1: 12, 2: 28, 3: 15, 4: 14, 5: 15, 6: 22}
    for col_idx, col_name in enumerate(columnas, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = len(str(col_name))
        if col_idx in (1, 2, 3):
            for row_idx in range(2, ws.max_row + 1):
                val = ws.cell(row=row_idx, column=col_idx).value
                if val is not None:
                    val_len = len(str(val))
                    if val_len > max_len:
                        max_len = val_len
        ws.column_dimensions[col_letter].width = max(
            max_len + 5, anchos_minimos.get(col_idx, 15)
        )


def _estilizar_hoja_resumen_dinamica(
    ws: Worksheet,
    df_resumen: pd.DataFrame,
    columnas: List[str],
    max_row_hoja1: int,
) -> None:
    """
    Genera la Hoja 2 ('Resumen por Empleado') conectada mediante fórmulas dinámicas
    a la Hoja 1 ('Fichadas Procesadas'), de modo que al agregar o editar horas en la
    Hoja 1 se recalculen solos los días trabajados, horas totales y promedios.
    """
    header_fill = PatternFill(
        start_color="112239", end_color="112239", fill_type="solid"
    )
    even_row_fill = PatternFill(
        start_color="F8FAFC", end_color="F8FAFC", fill_type="solid"
    )
    odd_row_fill = PatternFill(
        start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"
    )

    header_font = Font(name=FONT_FAMILY, size=11, bold=True, color="FFFFFF")
    cell_font = Font(name=FONT_FAMILY, size=10, bold=False, color="1E293B")

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    thin_side = Side(border_style="thin", color="CBD5E1")
    cell_border = Border(
        left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
    )

    ws.append(columnas)
    ws.row_dimensions[1].height = 28

    for col_idx in range(1, len(columnas) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center
        cell.border = cell_border

    ws.freeze_panes = "A2"

    h1_end = max(2, max_row_hoja1)

    for row_idx, row_data in enumerate(
        df_resumen[columnas].itertuples(index=False), start=2
    ):
        valores_raw = list(row_data)
        nro_id = valores_raw[0]
        nombre = valores_raw[1]

        # Fórmulas Excel enlazadas en vivo con 'Fichadas Procesadas'
        f_dias_fichaje = (
            f"=COUNTIFS('Fichadas Procesadas'!$A$2:$A${h1_end},A{row_idx},"
            f"'Fichadas Procesadas'!$D$2:$D${h1_end},\"<>\")"
        )
        f_jornadas_completas = (
            f"=COUNTIFS('Fichadas Procesadas'!$A$2:$A${h1_end},A{row_idx},"
            f"'Fichadas Procesadas'!$F$2:$F${h1_end},\">0\")"
        )
        f_dias_sin_salida = f"=MAX(0,C{row_idx}-D{row_idx})"
        f_total_horas = (
            f"=SUMIFS('Fichadas Procesadas'!$F$2:$F${h1_end},"
            f"'Fichadas Procesadas'!$A$2:$A${h1_end},A{row_idx})"
        )
        f_promedio_diario = f"=IF(D{row_idx}>0,F{row_idx}/D{row_idx},0)"

        ws.append(
            [
                nro_id,
                nombre,
                f_dias_fichaje,
                f_jornadas_completas,
                f_dias_sin_salida,
                f_total_horas,
                f_promedio_diario,
            ]
        )
        ws.row_dimensions[row_idx].height = 21

        row_fill = even_row_fill if row_idx % 2 == 0 else odd_row_fill

        for col_idx in range(1, len(columnas) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.fill = row_fill
            cell.border = cell_border
            cell.font = cell_font

            if col_idx == 2:
                cell.alignment = align_left
            else:
                cell.alignment = align_center

            if col_idx == 6:
                cell.number_format = "[h]:mm"
            elif col_idx == 7:
                cell.number_format = "hh:mm"

    if ws.max_row >= 1 and len(columnas) >= 1:
        ultima_col_letra = get_column_letter(len(columnas))
        ws.auto_filter.ref = f"A1:{ultima_col_letra}{max(1, ws.max_row)}"

    anchos_resumen = {1: 12, 2: 28, 3: 18, 4: 20, 5: 18, 6: 22, 7: 24}
    for col_idx, col_name in enumerate(columnas, start=1):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = anchos_resumen.get(
            col_idx, len(str(col_name)) + 6
        )


def exportar_excel_profesional(
    df_limpio: pd.DataFrame,
    ruta_destino: str,
    df_completo_original: Optional[pd.DataFrame] = None,
) -> Dict[str, int]:
    """
    Exporta siempre un libro de Excel con el 100% de los datos traducidos y fórmulas en vivo:
    - Hoja 1 ('Fichadas Procesadas'): El 100% de los registros del archivo original con cálculo
      automático de Tiempo Transcurrido y formato condicional al editar Ingreso o Salida.
    - Hoja 2 ('Resumen por Empleado'): Totalizador dinámico conectado a la Hoja 1.
    - Hoja 3 ('Solo Días con Fichaje'): Vista filtrada con fórmulas de cálculo automático.
    """
    wb = openpyxl.Workbook()

    base_completa = (
        df_completo_original if df_completo_original is not None else df_limpio
    )

    df_todos_original = filtrar_dataframe_según_modo(
        base_completa, modo="todos"
    )
    df_resumen = generar_resumen_por_empleado(base_completa)
    df_solo_fichadas = filtrar_dataframe_según_modo(
        base_completa, modo="con_fichadas"
    )

    # Hoja 1: Fichadas Procesadas (100% de los registros con fórmulas dinámicas)
    ws_principal = wb.active
    ws_principal.title = "Fichadas Procesadas"
    _estilizar_hoja_fichadas_dinamica(
        ws=ws_principal,
        df_datos=df_todos_original,
        columnas=OUTPUT_COLUMNS,
        columnas_centradas={1, 3, 4, 5, 6},
    )

    # Hoja 2: Resumen por Empleado (Con fórmulas COUNTIFS / SUMIFS en vivo hacia Hoja 1)
    ws_resumen = wb.create_sheet(title="Resumen por Empleado")
    _estilizar_hoja_resumen_dinamica(
        ws=ws_resumen,
        df_resumen=df_resumen,
        columnas=SUMMARY_COLUMNS,
        max_row_hoja1=ws_principal.max_row,
    )

    # Hoja 3: Solo Días con Fichaje (También con fórmulas dinámicas en vivo)
    ws_fichadas = wb.create_sheet(title="Solo Días con Fichaje")
    _estilizar_hoja_fichadas_dinamica(
        ws=ws_fichadas,
        df_datos=df_solo_fichadas,
        columnas=OUTPUT_COLUMNS,
        columnas_centradas={1, 3, 4, 5, 6},
    )

    wb.save(ruta_destino)

    return {
        "hoja_1_fichadas": len(df_todos_original),
        "hoja_2_resumen": len(df_resumen),
        "hoja_3_mensual": len(df_solo_fichadas),
    }
