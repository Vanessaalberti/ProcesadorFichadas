import importlib.util
import os
import re
import sys


def _aplicar_modulos_incrementales() -> None:
    """
    Si existen módulos actualizados incrementalmente en %LOCALAPPDATA%,
    reemplaza únicamente los módulos modificados en sys.modules sin tocar
    el resto del ejecutable compilado.
    """
    appdata = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    patch_src_dir = os.path.join(
        appdata, "ProcesadorFichadas_Leal", "live_updates", "src"
    )
    if not os.path.isdir(patch_src_dir):
        return

    # Verificar que el parche no sea más viejo que el ejecutable base
    patch_cfg = os.path.join(patch_src_dir, "config.py")
    if os.path.isfile(patch_cfg):
        try:
            import src.config as base_cfg

            with open(patch_cfg, "r", encoding="utf-8") as f_in:
                txt = f_in.read()
            m = re.search(r'APP_VERSION\s*=\s*["\']([^"\']+)["\']', txt)
            if m:
                ver_patch = tuple(
                    int(x) for x in re.findall(r"\d+", m.group(1))
                )
                ver_base = tuple(
                    int(x) for x in re.findall(r"\d+", base_cfg.APP_VERSION)
                )
                if ver_patch < ver_base:
                    return
        except Exception:
            return

    orden_carga = ["config", "processor", "exporter", "updater", "ui"]
    import src as paquete_src

    for nombre_mod in orden_carga:
        ruta_py = os.path.join(patch_src_dir, f"{nombre_mod}.py")
        if os.path.isfile(ruta_py):
            fullname = f"src.{nombre_mod}"
            try:
                spec = importlib.util.spec_from_file_location(fullname, ruta_py)
                if spec and spec.loader:
                    modulo = importlib.util.module_from_spec(spec)
                    sys.modules[fullname] = modulo
                    spec.loader.exec_module(modulo)
                    setattr(paquete_src, nombre_mod, modulo)
            except Exception:
                continue


_aplicar_modulos_incrementales()

from src.config import APP_VERSION, GITHUB_REPO, OUTPUT_COLUMNS
from src.exporter import exportar_excel_profesional
from src.processor import (
    cargar_archivo_crudo,
    filtrar_dataframe_según_modo,
    generar_resumen_por_empleado,
    obtener_estadisticas_df,
    procesar_fichadas_df,
)
from src.ui import ProcesadorFichadasApp
from src.updater import verificar_actualizaciones_async

__all__ = [
    "APP_VERSION",
    "GITHUB_REPO",
    "OUTPUT_COLUMNS",
    "ProcesadorFichadasApp",
    "cargar_archivo_crudo",
    "exportar_excel_profesional",
    "filtrar_dataframe_según_modo",
    "generar_resumen_por_empleado",
    "obtener_estadisticas_df",
    "procesar_fichadas_df",
    "verificar_actualizaciones_async",
]


def main() -> None:
    app = ProcesadorFichadasApp()
    app.mainloop()


if __name__ == "__main__":
    main()
