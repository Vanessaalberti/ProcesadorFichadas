# ==============================================================================
# CONFIGURACIÓN DE VERSIÓN, ACTUALIZACIONES OTA Y PALETA CORPORATIVA
# ==============================================================================

APP_VERSION = "1.3.0"
GITHUB_REPO = "Vanessaalberti/ProcesadorFichadas"

# Paleta corporativa exacta (Modo Claro)
COLOR_BG_OUTER = "#112239"           # Azul marino profundo del marco exterior
COLOR_BG_CONTAINER = "#F4F6F8"       # Gris muy claro para contenedores internos
COLOR_CARD_WHITE = "#FFFFFF"         # Tarjetas blancas
COLOR_BORDER = "#CBD5E1"             # Bordes sutiles de tarjetas y tablas
COLOR_DROPZONE_BORDER = "#64748B"    # Borde sobrio del área de carga
COLOR_DROPZONE_ACTIVE_BG = "#EFF6FF" # Fondo al arrastrar archivo encima
COLOR_DROPZONE_ACTIVE_BORDER = "#1565C0"

# Banner de estado de archivo adjunto
COLOR_ATTACHED_BG = "#ECFDF5"        # Fondo verde esmeralda suave cuando hay archivo
COLOR_ATTACHED_BORDER = "#10B981"    # Borde verde esmeralda intenso
COLOR_ATTACHED_TEXT = "#065F46"      # Texto principal de archivo adjunto
COLOR_ATTACHED_SUBTEXT = "#047857"   # Subtexto de archivo adjunto

COLOR_UNATTACHED_BG = "#F8FAFC"      # Fondo neutro cuando no hay archivo
COLOR_UNATTACHED_BORDER = "#CBD5E1"

# Botones principales
COLOR_BTN_BROWSE = "#1565C0"         # Azul corporativo "Examinar Archivo"
COLOR_BTN_BROWSE_HOVER = "#0D47A1"
COLOR_BTN_PROCESS = "#81C784"        # Verde suave empresarial "Procesar y Limpiar Datos"
COLOR_BTN_PROCESS_HOVER = "#66BB6A"
COLOR_BTN_PROCESS_TEXT = "#1B5E20"
COLOR_BTN_SAVE = "#90A4AE"           # Gris azulado "Guardar Reporte Organizado"
COLOR_BTN_SAVE_HOVER = "#78909C"
COLOR_BTN_SAVE_READY = "#1565C0"     # Azul corporativo activo cuando hay datos listos

# Textos y Tabla
COLOR_TEXT_PRIMARY = "#1E293B"
COLOR_TEXT_SECONDARY = "#64748B"
COLOR_TABLE_HEADER_BG = "#F1F5F9"
COLOR_TABLE_HEADER_FG = "#334155"
COLOR_ROW_ODD = "#FFFFFF"
COLOR_ROW_EVEN = "#F8FAFC"
COLOR_ROW_WARN_BG = "#FFFBEB"        # Fondo suave ámbar para filas "Sin salida"
COLOR_ROW_WARN_FG = "#B45309"        # Texto ámbar oscuro para "Sin salida"
COLOR_SUCCESS = "#10B981"
COLOR_WARNING = "#F59E0B"
COLOR_ERROR = "#EF4444"

FONT_FAMILY = "Segoe UI"
MAX_PREVIEW_ROWS = 50000

OUTPUT_COLUMNS = [
    "Nro. ID",
    "Nombre",
    "Fecha",
    "Ingreso",
    "Salida",
    "Tiempo Transcurrido",
]

SUMMARY_COLUMNS = [
    "Nro. ID",
    "Nombre",
    "Días con Fichaje",
    "Jornadas Completas",
    "Días Sin Salida",
    "Total Horas (HH:MM)",
    "Promedio Diario (HH:MM)",
]
