from typing import Dict, List, Optional, Set
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
import pandas as pd

from src.config import FONT_FAMILY, OUTPUT_COLUMNS, SUMMARY_COLUMNS
from src.processor import (
    filtrar_dataframe_según_modo,
    generar_resumen_por_empleado,
)


def _estilizar_hoja_excel(
    ws: Worksheet,
    df_datos: pd.DataFrame,
    columnas: List[str],
    columnas_centradas: Set[int],
    col_alerta_idx: Optional[int] = None,
) -> None:
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

    # 2. Filas de datos
    for row_idx, row_data in enumerate(
        df_datos[columnas].itertuples(index=False), start=2
    ):
        valores_fila = list(row_data)
        ws.append(valores_fila)
        ws.row_dimensions[row_idx].height = 21

        es_sin_salida = (
            col_alerta_idx is not None
            and len(valores_fila) >= col_alerta_idx
            and str(valores_fila[col_alerta_idx - 1]) == "Sin salida"
        )

        if es_sin_salida:
            row_fill = warn_row_fill
        else:
            row_fill = even_row_fill if row_idx % 2 == 0 else odd_row_fill

        for col_idx in range(1, len(columnas) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.fill = row_fill
            cell.border = cell_border
            if col_idx in columnas_centradas:
                cell.alignment = align_center
            else:
                cell.alignment = align_left

            if col_alerta_idx is not None and col_idx == col_alerta_idx and es_sin_salida:
                cell.font = warn_font
            else:
                cell.font = cell_font

    # 3. Activar filtro automático de Excel
    if ws.max_row >= 1 and len(columnas) >= 1:
        ultima_col_letra = get_column_letter(len(columnas))
        ws.auto_filter.ref = f"A1:{ultima_col_letra}{max(1, ws.max_row)}"

    # 4. Auto-ajuste real del ancho de columnas
    for col_idx, col_name in enumerate(columnas, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = len(str(col_name))
        for row_idx in range(2, ws.max_row + 1):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val is not None:
                val_len = len(str(val))
                if val_len > max_len:
                    max_len = val_len
        ws.column_dimensions[col_letter].width = max(max_len + 6, 15)


def exportar_excel_profesional(
    df_limpio: pd.DataFrame,
    ruta_destino: str,
    df_completo_original: Optional[pd.DataFrame] = None,
) -> Dict[str, int]:
    """
    Exporta siempre un libro de Excel con el 100% de los datos traducidos sin descartar ningún registro:
    - Hoja 1 ('Fichadas Procesadas'): El 100% de los registros del archivo original en su mismo orden
      (incluyendo todos los legajos tengan o no fichadas, para que RRHH filtre libremente en Excel).
    - Hoja 2 ('Resumen por Empleado'): Totalizador de días trabajados y horas de todos los legajos del archivo.
    - Hoja 3 ('Solo Días con Fichaje'): Vista filtrada únicamente con las filas que registraron fichadas.
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

    # Hoja 1: Fichadas Procesadas (100% de los registros del archivo original sin ocultar ninguno)
    ws_principal = wb.active
    ws_principal.title = "Fichadas Procesadas"
    _estilizar_hoja_excel(
        ws=ws_principal,
        df_datos=df_todos_original,
        columnas=OUTPUT_COLUMNS,
        columnas_centradas={1, 3, 4, 5, 6},
        col_alerta_idx=5,
    )

    # Hoja 2: Resumen por Empleado (100% de los legajos del archivo)
    ws_resumen = wb.create_sheet(title="Resumen por Empleado")
    _estilizar_hoja_excel(
        ws=ws_resumen,
        df_datos=df_resumen,
        columnas=SUMMARY_COLUMNS,
        columnas_centradas={1, 3, 4, 5, 6, 7},
        col_alerta_idx=None,
    )

    # Hoja 3: Solo Días con Fichaje (Vista rápida auxiliar de días fichados)
    ws_fichadas = wb.create_sheet(title="Solo Días con Fichaje")
    _estilizar_hoja_excel(
        ws=ws_fichadas,
        df_datos=df_solo_fichadas,
        columnas=OUTPUT_COLUMNS,
        columnas_centradas={1, 3, 4, 5, 6},
        col_alerta_idx=5,
    )

    wb.save(ruta_destino)

    return {
        "hoja_1_fichadas": len(df_todos_original),
        "hoja_2_resumen": len(df_resumen),
        "hoja_3_mensual": len(df_solo_fichadas),
    }
