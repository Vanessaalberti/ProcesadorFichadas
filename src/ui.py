import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Dict, Optional

import customtkinter as ctk
import pandas as pd
from tkinterdnd2 import DND_FILES, TkinterDnD

from src.config import (
    APP_VERSION,
    COLOR_ATTACHED_BG,
    COLOR_ATTACHED_BORDER,
    COLOR_ATTACHED_SUBTEXT,
    COLOR_ATTACHED_TEXT,
    COLOR_BG_CONTAINER,
    COLOR_BG_OUTER,
    COLOR_BORDER,
    COLOR_BTN_BROWSE,
    COLOR_BTN_BROWSE_HOVER,
    COLOR_BTN_PROCESS,
    COLOR_BTN_PROCESS_HOVER,
    COLOR_BTN_PROCESS_TEXT,
    COLOR_BTN_SAVE,
    COLOR_BTN_SAVE_HOVER,
    COLOR_BTN_SAVE_READY,
    COLOR_CARD_WHITE,
    COLOR_DROPZONE_ACTIVE_BG,
    COLOR_DROPZONE_ACTIVE_BORDER,
    COLOR_DROPZONE_BORDER,
    COLOR_ERROR,
    COLOR_ROW_EVEN,
    COLOR_ROW_ODD,
    COLOR_ROW_WARN_BG,
    COLOR_ROW_WARN_FG,
    COLOR_SUCCESS,
    COLOR_TABLE_HEADER_BG,
    COLOR_TABLE_HEADER_FG,
    COLOR_TEXT_PRIMARY,
    COLOR_TEXT_SECONDARY,
    COLOR_UNATTACHED_BG,
    COLOR_UNATTACHED_BORDER,
    FONT_FAMILY,
    MAX_PREVIEW_ROWS,
    OUTPUT_COLUMNS,
)
from src.exporter import exportar_excel_profesional
from src.processor import (
    cargar_archivo_crudo,
    filtrar_dataframe_según_modo,
    obtener_estadisticas_df,
    procesar_fichadas_df,
)
from src.updater import verificar_actualizaciones_async


def _formatear_tamano_archivo(filepath: str) -> str:
    try:
        size_bytes = os.path.getsize(filepath)
        if size_bytes < 1024:
            return f"{size_bytes} B"
        if size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.1f} KB"
        return f"{size_bytes / (1024 * 1024):.2f} MB"
    except Exception:
        return ""


class ProcesadorFichadasApp(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self) -> None:
        super().__init__()

        ctk.set_appearance_mode("light")

        # Inicializar motor nativo TkinterDnD para arrastrar y soltar archivos
        self.TkdndVersion = None
        try:
            self.TkdndVersion = TkinterDnD._require(self)
        except Exception:
            self.TkdndVersion = None

        self.title(f"Procesador de Fichadas v{APP_VERSION} - Leal Sistemas")
        self.geometry("940x710")
        self.minsize(840, 620)
        self.configure(fg_color=COLOR_BG_OUTER)

        self.filepath_actual: Optional[str] = None
        self.df_crudo: Optional[pd.DataFrame] = None
        self.df_procesado_completo: Optional[pd.DataFrame] = None
        self.df_vista_actual: Optional[pd.DataFrame] = None
        self.estadisticas_actuales: Dict[str, int] = obtener_estadisticas_df(None)
        self._procesando: bool = False
        self._modo_filtro: str = "todos"

        self._construir_ui()
        self._configurar_estilo_treeview()
        self._registrar_drag_and_drop()

        # Verificación silenciosa de actualizaciones en segundo plano
        self.after(400, lambda: verificar_actualizaciones_async(self))

    def _construir_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=1)

        # ======================================================================
        # 1. TARJETA SUPERIOR: DROPZONE Y BANNER DE ARCHIVO ADJUNTO
        # ======================================================================
        self.top_card = ctk.CTkFrame(
            self,
            fg_color=COLOR_BG_CONTAINER,
            corner_radius=10,
            border_width=1,
            border_color=COLOR_BORDER,
        )
        self.top_card.grid(row=0, column=0, padx=20, pady=(16, 8), sticky="ew")
        self.top_card.grid_columnconfigure(0, weight=1)

        # Área de arrastrar y soltar (Dropzone)
        self.dropzone_frame = ctk.CTkFrame(
            self.top_card,
            fg_color=COLOR_CARD_WHITE,
            corner_radius=8,
            border_width=2,
            border_color=COLOR_DROPZONE_BORDER,
        )
        self.dropzone_frame.grid(
            row=0, column=0, padx=16, pady=(14, 10), sticky="ew"
        )
        self.dropzone_frame.grid_columnconfigure(0, weight=0)
        self.dropzone_frame.grid_columnconfigure(1, weight=1)

        # Ícono vectorial de documento (con indicador de estado)
        self.doc_canvas = tk.Canvas(
            self.dropzone_frame,
            width=66,
            height=74,
            bg=COLOR_CARD_WHITE,
            highlightthickness=0,
            bd=0,
            cursor="hand2",
        )
        self.doc_canvas.grid(
            row=0, column=0, rowspan=2, padx=(36, 16), pady=14, sticky="w"
        )
        self._dibujar_icono_documento(adjuntado=False, hover=False)

        # Contenedor central de Dropzone
        self.drop_center_frame = ctk.CTkFrame(
            self.dropzone_frame,
            fg_color="transparent",
        )
        self.drop_center_frame.grid(
            row=0, column=1, padx=(0, 50), pady=12, sticky="ew"
        )
        self.drop_center_frame.grid_columnconfigure(0, weight=1)

        self.lbl_drop_title = ctk.CTkLabel(
            self.drop_center_frame,
            text="Arrastra aquí tu archivo de fichadas (.xls, .xlsx, .csv) o",
            font=ctk.CTkFont(family=FONT_FAMILY, size=16, weight="bold"),
            text_color=COLOR_TEXT_PRIMARY,
        )
        self.lbl_drop_title.grid(row=0, column=0, pady=(2, 8))

        self.btn_browse = ctk.CTkButton(
            self.drop_center_frame,
            text="Examinar Archivo",
            font=ctk.CTkFont(family=FONT_FAMILY, size=14, weight="bold"),
            fg_color=COLOR_BTN_BROWSE,
            hover_color=COLOR_BTN_BROWSE_HOVER,
            text_color="#FFFFFF",
            corner_radius=6,
            width=184,
            height=36,
            command=self._accion_examinar_archivo,
        )
        self.btn_browse.grid(row=1, column=0, pady=(0, 4))

        self.lbl_drop_hint = ctk.CTkLabel(
            self.drop_center_frame,
            text="Compatible con reportes biométricos Leal Sistemas (.xls antiguos, tablas HTML y CSV)",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="normal"),
            text_color=COLOR_TEXT_SECONDARY,
        )
        self.lbl_drop_hint.grid(row=2, column=0, pady=(2, 0))

        # Permitir clic en el área de dropzone para examinar archivo
        for widget_clickable in (
            self.dropzone_frame,
            self.doc_canvas,
            self.lbl_drop_title,
            self.lbl_drop_hint,
        ):
            widget_clickable.bind(
                "<Button-1>", lambda _e: self._accion_examinar_archivo()
            )

        # Banner de Estado del Archivo Adjunto (Destacado visualmente)
        self.status_banner = ctk.CTkFrame(
            self.top_card,
            fg_color=COLOR_UNATTACHED_BG,
            corner_radius=8,
            border_width=1,
            border_color=COLOR_UNATTACHED_BORDER,
        )
        self.status_banner.grid(
            row=1, column=0, padx=16, pady=(0, 14), sticky="ew"
        )
        self.status_banner.grid_columnconfigure(1, weight=1)
        self.status_banner.grid_columnconfigure(2, weight=0)

        self.status_badge_canvas = tk.Canvas(
            self.status_banner,
            width=36,
            height=36,
            bg=COLOR_UNATTACHED_BG,
            highlightthickness=0,
            bd=0,
        )
        self.status_badge_canvas.grid(
            row=0, column=0, rowspan=2, padx=(14, 12), pady=10, sticky="w"
        )
        self._dibujar_badge_estado(cargado=False, error=False)

        self.lbl_filename = ctk.CTkLabel(
            self.status_banner,
            text="SIN ARCHIVO ADJUNTO",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            text_color=COLOR_TEXT_SECONDARY,
            anchor="w",
        )
        self.lbl_filename.grid(row=0, column=1, pady=(8, 0), sticky="ew")

        self.lbl_file_status = ctk.CTkLabel(
            self.status_banner,
            text="Arrastra un archivo al recuadro superior o presiona 'Examinar Archivo' para adjuntarlo.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="normal"),
            text_color=COLOR_TEXT_SECONDARY,
            anchor="w",
        )
        self.lbl_file_status.grid(row=1, column=1, pady=(0, 8), sticky="ew")

        self.lbl_attached_pill = ctk.CTkLabel(
            self.status_banner,
            text="ESPERANDO ARCHIVO",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            fg_color="#E2E8F0",
            text_color="#475569",
            corner_radius=6,
            padx=10,
            pady=4,
        )
        self.lbl_attached_pill.grid(
            row=0, column=2, rowspan=2, padx=(8, 14), pady=10, sticky="e"
        )

        # ======================================================================
        # 2. TARJETA INFERIOR: BOTONES, FILTROS, KPIS Y TABLA TREEVIEW
        # ======================================================================
        self.bottom_card = ctk.CTkFrame(
            self,
            fg_color=COLOR_BG_CONTAINER,
            corner_radius=10,
            border_width=1,
            border_color=COLOR_BORDER,
        )
        self.bottom_card.grid(
            row=1, column=0, padx=20, pady=(4, 18), sticky="nsew"
        )
        self.bottom_card.grid_columnconfigure(0, weight=1)
        self.bottom_card.grid_rowconfigure(3, weight=1)

        # Barra de Botones de Acción Principal
        self.actions_frame = ctk.CTkFrame(
            self.bottom_card,
            fg_color="transparent",
        )
        self.actions_frame.grid(
            row=0, column=0, padx=16, pady=(14, 8), sticky="ew"
        )
        self.actions_frame.grid_columnconfigure(0, weight=1)
        self.actions_frame.grid_columnconfigure(1, weight=1)

        self.btn_process = ctk.CTkButton(
            self.actions_frame,
            text="Procesar y Limpiar Datos   ⚙",
            font=ctk.CTkFont(family=FONT_FAMILY, size=15, weight="bold"),
            fg_color=COLOR_BTN_PROCESS,
            hover_color=COLOR_BTN_PROCESS_HOVER,
            text_color=COLOR_BTN_PROCESS_TEXT,
            corner_radius=8,
            height=42,
            command=self._accion_procesar_datos,
        )
        self.btn_process.grid(row=0, column=0, padx=(0, 8), sticky="ew")

        self.btn_save = ctk.CTkButton(
            self.actions_frame,
            text="↓   Guardar Reporte Organizado (.xlsx)",
            font=ctk.CTkFont(family=FONT_FAMILY, size=15, weight="bold"),
            fg_color=COLOR_BTN_SAVE,
            hover_color=COLOR_BTN_SAVE_HOVER,
            text_color="#FFFFFF",
            corner_radius=8,
            height=42,
            command=self._accion_guardar_reporte,
        )
        self.btn_save.grid(row=0, column=1, padx=(8, 0), sticky="ew")

        # Barra de Filtros de Vista y Buscador en Tiempo Real
        self.toolbar_frame = ctk.CTkFrame(
            self.bottom_card,
            fg_color=COLOR_CARD_WHITE,
            corner_radius=8,
            border_width=1,
            border_color=COLOR_BORDER,
        )
        self.toolbar_frame.grid(
            row=1, column=0, padx=16, pady=(2, 6), sticky="ew"
        )
        self.toolbar_frame.grid_columnconfigure(1, weight=0)
        self.toolbar_frame.grid_columnconfigure(2, weight=1)

        self.lbl_filter_caption = ctk.CTkLabel(
            self.toolbar_frame,
            text="Vista:",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_TEXT_PRIMARY,
        )
        self.lbl_filter_caption.grid(row=0, column=0, padx=(12, 6), pady=8, sticky="w")

        self.filter_btns_frame = ctk.CTkFrame(
            self.toolbar_frame,
            fg_color="#E2E8F0",
            corner_radius=6,
        )
        self.filter_btns_frame.grid(row=0, column=1, padx=(0, 12), pady=6, sticky="w")

        self._botones_filtro: Dict[str, ctk.CTkButton] = {}
        opciones_filtro = [
            ("todos", "Todos los registros"),
            ("con_fichadas", "Solo días con fichadas"),
            ("empleados_activos", "Empleados activos (mes)"),
        ]
        for idx_f, (clave_modo, etiqueta_modo) in enumerate(opciones_filtro):
            btn_f = ctk.CTkButton(
                self.filter_btns_frame,
                text=etiqueta_modo,
                font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
                height=28,
                corner_radius=5,
                command=lambda m=clave_modo: self._seleccionar_modo_filtro(m),
            )
            btn_f.grid(row=0, column=idx_f, padx=2, pady=2)
            self._botones_filtro[clave_modo] = btn_f

        self._actualizar_estilo_botones_filtro()

        # Buscador rápido por empleado / ID / fecha
        self.search_var = tk.StringVar(value="")
        self.search_var.trace_add("write", lambda *_: self._aplicar_filtro_y_busqueda())

        self.entry_search = ctk.CTkEntry(
            self.toolbar_frame,
            textvariable=self.search_var,
            placeholder_text="Buscar empleado por Nombre, Nro. ID o Fecha...",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            height=30,
            fg_color="#F8FAFC",
            border_color=COLOR_BORDER,
            text_color=COLOR_TEXT_PRIMARY,
        )
        self.entry_search.grid(row=0, column=2, padx=(0, 12), pady=8, sticky="ew")

        # Barra de Resumen Estadístico (KPIs)
        self.lbl_summary = ctk.CTkLabel(
            self.bottom_card,
            text="Vista previa de registros procesados — Adjunta un archivo para comenzar",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_TEXT_PRIMARY,
            anchor="w",
        )
        self.lbl_summary.grid(row=2, column=0, padx=18, pady=(0, 4), sticky="ew")

        # Contenedor de Tabla Treeview
        self.table_card = ctk.CTkFrame(
            self.bottom_card,
            fg_color=COLOR_CARD_WHITE,
            corner_radius=8,
            border_width=1,
            border_color=COLOR_BORDER,
        )
        self.table_card.grid(
            row=3, column=0, padx=16, pady=(0, 14), sticky="nsew"
        )
        self.table_card.grid_columnconfigure(0, weight=1)
        self.table_card.grid_rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            self.table_card,
            columns=OUTPUT_COLUMNS,
            show="headings",
            selectmode="browse",
        )
        self.scrollbar_y = ttk.Scrollbar(
            self.table_card,
            orient="vertical",
            command=self.tree.yview,
        )
        self.tree.configure(yscrollcommand=self.scrollbar_y.set)

        self.tree.grid(row=0, column=0, padx=(6, 0), pady=6, sticky="nsew")
        self.scrollbar_y.grid(row=0, column=1, padx=(0, 6), pady=6, sticky="ns")

        anchos_columnas = {
            "Nro. ID": (85, "center"),
            "Nombre": (230, "w"),
            "Fecha": (115, "center"),
            "Ingreso": (100, "center"),
            "Salida": (115, "center"),
            "Tiempo Transcurrido": (155, "center"),
        }
        for col in OUTPUT_COLUMNS:
            ancho, alineacion = anchos_columnas[col]
            self.tree.heading(col, text=col, anchor=alineacion)
            self.tree.column(
                col,
                width=ancho,
                minwidth=75,
                anchor=alineacion,
                stretch=True,
            )

        self.tree.tag_configure(
            "oddrow", background=COLOR_ROW_ODD, foreground=COLOR_TEXT_PRIMARY
        )
        self.tree.tag_configure(
            "evenrow", background=COLOR_ROW_EVEN, foreground=COLOR_TEXT_PRIMARY
        )
        self.tree.tag_configure(
            "warnrow",
            background=COLOR_ROW_WARN_BG,
            foreground=COLOR_ROW_WARN_FG,
        )

    def _configurar_estilo_treeview(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure(
            "Treeview",
            background=COLOR_CARD_WHITE,
            foreground=COLOR_TEXT_PRIMARY,
            fieldbackground=COLOR_CARD_WHITE,
            rowheight=28,
            font=(FONT_FAMILY, 10),
            borderwidth=0,
            relief="flat",
        )
        style.configure(
            "Treeview.Heading",
            background=COLOR_TABLE_HEADER_BG,
            foreground=COLOR_TABLE_HEADER_FG,
            font=(FONT_FAMILY, 10, "bold"),
            borderwidth=0,
            relief="flat",
            padding=(8, 6),
        )
        style.map(
            "Treeview.Heading",
            background=[("active", "#E2E8F0")],
        )
        style.map(
            "Treeview",
            background=[("selected", "#DBEAFE")],
            foreground=[("selected", "#1E3A8A")],
        )

    # ==========================================================================
    # SOPORTE DRAG & DROP CON FEEDBACK VISUAL DINÁMICO
    # ==========================================================================
    def _registrar_drag_and_drop(self) -> None:
        if not self.TkdndVersion:
            return

        widgets_destino = [
            self,
            self.top_card,
            self.dropzone_frame,
            self.doc_canvas,
            self.drop_center_frame,
            self.lbl_drop_title,
            self.lbl_drop_hint,
            self.status_banner,
        ]
        for w in widgets_destino:
            try:
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<DropEnter>>", self._on_drag_enter)
                w.dnd_bind("<<DropLeave>>", self._on_drag_leave)
                w.dnd_bind("<<Drop>>", self._on_file_drop)
            except Exception:
                continue

    def _on_drag_enter(self, _event) -> str:
        self.dropzone_frame.configure(
            fg_color=COLOR_DROPZONE_ACTIVE_BG,
            border_color=COLOR_DROPZONE_ACTIVE_BORDER,
            border_width=2,
        )
        self.doc_canvas.configure(bg=COLOR_DROPZONE_ACTIVE_BG)
        self._dibujar_icono_documento(
            adjuntado=self.filepath_actual is not None, hover=True
        )
        self.lbl_drop_title.configure(
            text="¡Suelta aquí tu archivo para adjuntarlo y procesarlo!",
            text_color=COLOR_BTN_BROWSE,
        )
        return "copy"

    def _on_drag_leave(self, _event) -> None:
        self._restaurar_aspecto_dropzone()

    def _restaurar_aspecto_dropzone(self) -> None:
        tiene_archivo = self.filepath_actual is not None
        bg_color = COLOR_CARD_WHITE
        border_col = COLOR_ATTACHED_BORDER if tiene_archivo else COLOR_DROPZONE_BORDER

        self.dropzone_frame.configure(
            fg_color=bg_color,
            border_color=border_col,
            border_width=2,
        )
        self.doc_canvas.configure(bg=bg_color)
        self._dibujar_icono_documento(adjuntado=tiene_archivo, hover=False)

        if tiene_archivo and self.filepath_actual:
            nombre = os.path.basename(self.filepath_actual)
            self.lbl_drop_title.configure(
                text=f"Archivo activo: {nombre} (Arrastra otro para reemplazarlo)",
                text_color=COLOR_ATTACHED_TEXT,
            )
        else:
            self.lbl_drop_title.configure(
                text="Arrastra aquí tu archivo de fichadas (.xls, .xlsx, .csv) o",
                text_color=COLOR_TEXT_PRIMARY,
            )

    def _on_file_drop(self, event) -> None:
        self._restaurar_aspecto_dropzone()
        if not event or not getattr(event, "data", None):
            return

        try:
            rutas = self.tk.splitlist(event.data)
        except Exception:
            rutas = [str(event.data).strip("{}")]

        if rutas:
            primera_ruta = str(rutas[0]).strip()
            if os.path.isfile(primera_ruta):
                self._cargar_archivo_seleccionado(primera_ruta)
            else:
                messagebox.showwarning(
                    "Archivo no válido",
                    f"No se pudo acceder al archivo arrastrado:\n{primera_ruta}",
                )

    # ==========================================================================
    # ICONOGRAFÍA VECTORIAL EN CANVAS
    # ==========================================================================
    def _dibujar_icono_documento(
        self, adjuntado: bool = False, hover: bool = False
    ) -> None:
        self.doc_canvas.delete("all")
        if hover:
            color_linea = COLOR_BTN_BROWSE
            color_relleno = "#DBEAFE"
        elif adjuntado:
            color_linea = COLOR_ATTACHED_BORDER
            color_relleno = COLOR_ATTACHED_BG
        else:
            color_linea = "#475569"
            color_relleno = "#FFFFFF"

        puntos_hoja = [12, 6, 38, 6, 50, 18, 50, 64, 12, 64]
        self.doc_canvas.create_polygon(
            puntos_hoja,
            outline=color_linea,
            fill=color_relleno,
            width=2,
        )
        self.doc_canvas.create_line(
            38, 6, 38, 18, 50, 18,
            fill=color_linea,
            width=2,
        )
        # Líneas simulando tabla de datos dentro del ícono
        self.doc_canvas.create_line(19, 28, 43, 28, fill=color_linea, width=2)
        self.doc_canvas.create_line(19, 38, 43, 38, fill=color_linea, width=2)
        self.doc_canvas.create_line(19, 48, 35, 48, fill=color_linea, width=2)

        if adjuntado:
            # Sello circular verde con check en la esquina inferior derecha del documento
            self.doc_canvas.create_oval(
                36, 44, 60, 68, fill=COLOR_SUCCESS, outline="#FFFFFF", width=2
            )
            self.doc_canvas.create_text(
                48,
                56,
                text="✓",
                fill="#FFFFFF",
                font=(FONT_FAMILY, 11, "bold"),
            )

    def _dibujar_badge_estado(self, cargado: bool, error: bool = False) -> None:
        self.status_badge_canvas.delete("all")
        if error:
            color_circulo = COLOR_ERROR
            simbolo = "!"
            bg_canvas = "#FEF2F2"
        elif cargado:
            color_circulo = COLOR_SUCCESS
            simbolo = "✓"
            bg_canvas = COLOR_ATTACHED_BG
        else:
            color_circulo = "#94A3B8"
            simbolo = "•"
            bg_canvas = COLOR_UNATTACHED_BG

        self.status_badge_canvas.configure(bg=bg_canvas)
        self.status_badge_canvas.create_oval(
            3, 3, 33, 33, fill=color_circulo, outline=color_circulo
        )
        self.status_badge_canvas.create_text(
            18,
            18,
            text=simbolo,
            fill="#FFFFFF",
            font=(FONT_FAMILY, 13, "bold"),
        )

    # ==========================================================================
    # FLUJO DE CARGA, PROCESAMIENTO Y EXPORTACIÓN (MULTITHREADING)
    # ==========================================================================
    def _accion_examinar_archivo(self) -> None:
        if self._procesando:
            return

        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo de fichadas - Leal Sistemas",
            filetypes=[
                ("Archivos de Fichadas", "*.xls *.xlsx *.csv *.tsv *.txt"),
                ("Archivos Excel", "*.xls *.xlsx"),
                ("Archivos CSV / Texto", "*.csv *.tsv *.txt"),
                ("Todos los archivos", "*.*"),
            ],
        )
        if ruta:
            self._cargar_archivo_seleccionado(ruta)

    def _cargar_archivo_seleccionado(self, filepath: str) -> None:
        if self._procesando:
            return

        self._procesando = True
        nombre_archivo = os.path.basename(filepath)

        self.status_banner.configure(
            fg_color="#EFF6FF",
            border_color=COLOR_BTN_BROWSE,
            border_width=2,
        )
        self.lbl_filename.configure(
            text=f"CARGANDO Y ANALIZANDO: {nombre_archivo}...",
            text_color=COLOR_BTN_BROWSE,
        )
        self.lbl_file_status.configure(
            text="Leyendo estructura biométrica y limpiando marcas de horario...",
            text_color=COLOR_TEXT_SECONDARY,
        )
        self.lbl_attached_pill.configure(
            text="PROCESANDO...",
            fg_color="#DBEAFE",
            text_color=COLOR_BTN_BROWSE,
        )
        self.btn_browse.configure(state="disabled")
        self.btn_process.configure(state="disabled")
        self.btn_save.configure(state="disabled")

        def _worker() -> None:
            try:
                df_raw = cargar_archivo_crudo(filepath)
                df_clean = procesar_fichadas_df(df_raw)
                stats = obtener_estadisticas_df(df_clean)
                self.after(
                    0,
                    lambda: self._on_archivo_cargado_y_procesado_ok(
                        filepath, df_raw, df_clean, stats
                    ),
                )
            except Exception as exc:
                msg = str(exc)
                self.after(
                    0, lambda: self._on_archivo_cargado_error(nombre_archivo, msg)
                )

        threading.Thread(target=_worker, daemon=True).start()

    def _on_archivo_cargado_y_procesado_ok(
        self,
        filepath: str,
        df_raw: pd.DataFrame,
        df_clean: pd.DataFrame,
        stats: Dict[str, int],
    ) -> None:
        self._procesando = False
        self.filepath_actual = filepath
        self.df_crudo = df_raw
        self.df_procesado_completo = df_clean
        self.estadisticas_actuales = stats

        nombre_archivo = os.path.basename(filepath)
        tamano_str = _formatear_tamano_archivo(filepath)

        # Actualizar Dropzone y Banner de Archivo Adjunto
        self._restaurar_aspecto_dropzone()

        self.status_banner.configure(
            fg_color=COLOR_ATTACHED_BG,
            border_color=COLOR_ATTACHED_BORDER,
            border_width=2,
        )
        self._dibujar_badge_estado(cargado=True, error=False)

        self.lbl_filename.configure(
            text=f"ARCHIVO ADJUNTO: {nombre_archivo} ({tamano_str})",
            text_color=COLOR_ATTACHED_TEXT,
        )
        self.lbl_file_status.configure(
            text=(
                f"✓ Cargado y procesado con éxito • {stats['dias_con_fichadas']} días con fichadas reales "
                f"({stats['empleados_activos']} empleados activos de {stats['empleados_totales']} legajos en archivo)"
            ),
            text_color=COLOR_ATTACHED_SUBTEXT,
        )
        self.lbl_attached_pill.configure(
            text="✓ ARCHIVO ADJUNTO",
            fg_color="#10B981",
            text_color="#FFFFFF",
        )

        # Actualizar textos de pestañas con conteos reales
        self.btn_browse.configure(state="normal", text="Cambiar Archivo")
        self.btn_process.configure(state="normal")
        self.btn_save.configure(
            state="normal",
            fg_color=COLOR_BTN_SAVE_READY,
            hover_color=COLOR_BTN_BROWSE_HOVER,
        )

        # Mostrar de inmediato los datos limpios en la tabla
        self._aplicar_filtro_y_busqueda()

    def _on_archivo_cargado_error(
        self, nombre_archivo: str, error_msg: str
    ) -> None:
        self._procesando = False
        self.filepath_actual = None
        self.df_crudo = None
        self.df_procesado_completo = None
        self.df_vista_actual = None

        self._restaurar_aspecto_dropzone()
        self.status_banner.configure(
            fg_color="#FEF2F2",
            border_color=COLOR_ERROR,
            border_width=2,
        )
        self._dibujar_badge_estado(cargado=False, error=True)
        self.lbl_filename.configure(
            text=f"ERROR AL ADJUNTAR: {nombre_archivo}",
            text_color="#991B1B",
        )
        self.lbl_file_status.configure(
            text="El archivo no pudo ser interpretado como un reporte válido.",
            text_color="#B91C1C",
        )
        self.lbl_attached_pill.configure(
            text="ERROR DE FORMATO",
            fg_color=COLOR_ERROR,
            text_color="#FFFFFF",
        )
        self.btn_browse.configure(state="normal")
        self.btn_process.configure(state="normal")
        self.btn_save.configure(state="normal", fg_color=COLOR_BTN_SAVE)
        messagebox.showerror("Error al cargar archivo", error_msg)

    def _accion_procesar_datos(self) -> None:
        if self._procesando:
            return

        if self.df_crudo is None or self.filepath_actual is None:
            messagebox.showwarning(
                "Archivo requerido",
                "Primero debes arrastrar o examinar un archivo de fichadas (.xls, .xlsx o .csv).",
            )
            return

        self._procesando = True
        self.btn_process.configure(
            state="disabled", text="Limpiando y organizando..."
        )
        self.btn_browse.configure(state="disabled")
        self.btn_save.configure(state="disabled")

        df_entrada = self.df_crudo

        def _worker() -> None:
            try:
                df_clean = procesar_fichadas_df(df_entrada)
                stats = obtener_estadisticas_df(df_clean)
                self.after(0, lambda: self._on_reprocesamiento_ok(df_clean, stats))
            except Exception as exc:
                msg = str(exc)
                self.after(0, lambda: self._on_reprocesamiento_error(msg))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_reprocesamiento_ok(
        self, df_clean: pd.DataFrame, stats: Dict[str, int]
    ) -> None:
        self._procesando = False
        self.df_procesado_completo = df_clean
        self.estadisticas_actuales = stats

        self.btn_process.configure(
            state="normal", text="Procesar y Limpiar Datos   ⚙"
        )
        self.btn_browse.configure(state="normal")
        self.btn_save.configure(
            state="normal",
            fg_color=COLOR_BTN_SAVE_READY,
            hover_color=COLOR_BTN_BROWSE_HOVER,
        )

        self._aplicar_filtro_y_busqueda()

    def _on_reprocesamiento_error(self, error_msg: str) -> None:
        self._procesando = False
        self.btn_process.configure(
            state="normal", text="Procesar y Limpiar Datos   ⚙"
        )
        self.btn_browse.configure(state="normal")
        self.btn_save.configure(state="normal")
        messagebox.showerror(
            "Error al procesar datos",
            f"Ocurrió un problema al limpiar los datos:\n{error_msg}",
        )

    def _actualizar_estilo_botones_filtro(self) -> None:
        for clave_modo, btn in self._botones_filtro.items():
            if clave_modo == self._modo_filtro:
                btn.configure(
                    fg_color=COLOR_BTN_BROWSE,
                    hover_color=COLOR_BTN_BROWSE_HOVER,
                    text_color="#FFFFFF",
                )
            else:
                btn.configure(
                    fg_color="#E2E8F0",
                    hover_color="#CBD5E1",
                    text_color=COLOR_TEXT_PRIMARY,
                )

    def _seleccionar_modo_filtro(self, modo: str) -> None:
        self._modo_filtro = modo
        self._actualizar_estilo_botones_filtro()
        self._aplicar_filtro_y_busqueda()

    def _aplicar_filtro_y_busqueda(self) -> None:
        if self.df_procesado_completo is None:
            return

        texto_busqueda = self.search_var.get()
        df_filtrado = filtrar_dataframe_según_modo(
            self.df_procesado_completo,
            modo=self._modo_filtro,
            texto_busqueda=texto_busqueda,
        )
        self.df_vista_actual = df_filtrado
        self._actualizar_vista_previa(df_filtrado)

    def _actualizar_vista_previa(self, df_mostrar: pd.DataFrame) -> None:
        items_actuales = self.tree.get_children()
        if items_actuales:
            self.tree.delete(*items_actuales)

        total_vista = len(df_mostrar)
        stats = self.estadisticas_actuales

        df_preview = df_mostrar.head(MAX_PREVIEW_ROWS)

        for idx, fila in enumerate(
            df_preview[OUTPUT_COLUMNS].itertuples(index=False)
        ):
            es_sin_salida = str(fila[4]) == "Sin salida"
            if es_sin_salida:
                tag = "warnrow"
            else:
                tag = "evenrow" if idx % 2 == 1 else "oddrow"
            self.tree.insert("", "end", values=tuple(fila), tags=(tag,))

        if total_vista > MAX_PREVIEW_ROWS:
            prefijo_vista = f"Mostrando primeras {MAX_PREVIEW_ROWS} de {total_vista} filas en esta vista"
        else:
            prefijo_vista = f"Mostrando {total_vista} filas en esta vista"

        resumen_kpi = (
            f"{prefijo_vista}   |   "
            f"Días con fichaje: {stats['dias_con_fichadas']}   •   "
            f"Jornadas completas: {stats['jornadas_completas']}   •   "
            f"Sin salida: {stats['sin_salida']}   •   "
            f"Empleados activos: {stats['empleados_activos']}"
        )
        self.lbl_summary.configure(text=resumen_kpi)

    def _accion_guardar_reporte(self) -> None:
        if self._procesando:
            return

        if (
            self.df_procesado_completo is None
            or self.df_procesado_completo.empty
            or self.df_vista_actual is None
        ):
            messagebox.showwarning(
                "Sin datos para exportar",
                "Primero debes adjuntar un archivo de fichadas válido.",
            )
            return

        nombre_base = "Fichadas"
        if self.filepath_actual:
            nombre_sin_ext = os.path.splitext(
                os.path.basename(self.filepath_actual)
            )[0]
            nombre_base = "_".join(nombre_sin_ext.split())

        nombre_sugerido = f"Reporte_Limpio_{nombre_base}.xlsx"

        ruta_destino = filedialog.asksaveasfilename(
            title="Guardar Reporte Organizado en Excel",
            defaultextension=".xlsx",
            initialfile=nombre_sugerido,
            filetypes=[("Libro de Excel (*.xlsx)", "*.xlsx")],
        )
        if not ruta_destino:
            return

        self._procesando = True
        self.btn_save.configure(state="disabled", text="Generando Excel...")

        df_hoja_principal = self.df_vista_actual.copy()
        df_base_completa = self.df_procesado_completo.copy()

        def _worker() -> None:
            try:
                conteos = exportar_excel_profesional(
                    df_limpio=df_hoja_principal,
                    ruta_destino=ruta_destino,
                    df_completo_original=df_base_completa,
                )
                self.after(
                    0,
                    lambda: self._on_guardado_ok(ruta_destino, conteos),
                )
            except Exception as exc:
                msg = str(exc)
                self.after(0, lambda: self._on_guardado_error(msg))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_guardado_ok(
        self, ruta_destino: str, conteos: Dict[str, int]
    ) -> None:
        self._procesando = False
        self.btn_save.configure(
            state="normal", text="↓   Guardar Reporte Organizado (.xlsx)"
        )
        c1 = conteos.get("hoja_1_fichadas", 0)
        c2 = conteos.get("hoja_2_resumen", 0)
        c3 = conteos.get("hoja_3_mensual", 0)
        messagebox.showinfo(
            "Reporte Excel Generado con Éxito",
            f"El reporte organizado se ha guardado correctamente con todos los registros:\n\n"
            f"• Archivo: {os.path.basename(ruta_destino)}\n"
            f"• Hoja 1 (Fichadas Procesadas): {c1} registros completos (100% del archivo original)\n"
            f"• Hoja 2 (Resumen por Empleado): {c2} legajos totales\n"
            f"• Hoja 3 (Solo Días con Fichaje): {c3} días con fichaje registrado\n\n"
            f"Ubicación:\n{ruta_destino}",
        )

    def _on_guardado_error(self, error_msg: str) -> None:
        self._procesando = False
        self.btn_save.configure(
            state="normal", text="↓   Guardar Reporte Organizado (.xlsx)"
        )
        messagebox.showerror(
            "Error al guardar Excel",
            f"No se pudo guardar el archivo Excel:\n{error_msg}\n\n"
            "Verifique que el archivo no esté abierto en Excel e intente nuevamente.",
        )
