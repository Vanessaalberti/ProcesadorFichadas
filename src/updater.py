import hashlib
import importlib
import os
import re
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import messagebox
from typing import Dict, List, Optional, Tuple

import requests

from src.config import APP_VERSION, GITHUB_REPO

# Archivos modulares livianos (~35 KB en total) que pueden actualizarse
# incrementalmente sin necesidad de reinstalar el ejecutable de 50 MB.
MODULOS_ACTUALIZABLES = [
    "src/__init__.py",
    "src/config.py",
    "src/processor.py",
    "src/exporter.py",
    "src/updater.py",
    "src/ui.py",
]


def obtener_directorio_parches() -> str:
    """
    Devuelve la ruta donde se guardan únicamente los módulos actualizados (.py),
    permitiendo actualizar solo lo que cambió sin reinstalar todo el programa.
    """
    appdata = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    patch_dir = os.path.join(appdata, "ProcesadorFichadas_Leal", "live_updates")
    os.makedirs(patch_dir, exist_ok=True)
    return patch_dir


def activar_parches_locales() -> None:
    """
    Inyecta el directorio de actualizaciones incrementales al inicio de sys.path
    para que Python cargue los archivos modificados si existen.
    """
    patch_dir = obtener_directorio_parches()
    src_patch_dir = os.path.join(patch_dir, "src")
    if os.path.isdir(src_patch_dir) and patch_dir not in sys.path:
        sys.path.insert(0, patch_dir)


def parse_version_tuple(version_str: str) -> Tuple[int, int, int]:
    cleaned = version_str.strip().lstrip("vV")
    parts = [int(p) for p in re.findall(r"\d+", cleaned)]
    while len(parts) < 3:
        parts.append(0)
    return (parts[0], parts[1], parts[2])


def _calcular_sha256_texto(contenido: bytes) -> str:
    return hashlib.sha256(contenido).hexdigest()


def _obtener_version_local_efectiva() -> str:
    patch_config = os.path.join(obtener_directorio_parches(), "src", "config.py")
    if os.path.isfile(patch_config):
        try:
            with open(patch_config, "r", encoding="utf-8") as f_cfg:
                texto = f_cfg.read()
            m = re.search(r'APP_VERSION\s*=\s*["\']([^"\']+)["\']', texto)
            if m:
                return m.group(1).strip()
        except Exception:
            return APP_VERSION
    return APP_VERSION


def _aplicar_actualizacion_incremental(
    archivos_modificados: Dict[str, bytes],
    nueva_version: str,
    app_root: tk.Misc,
) -> None:
    """
    Escribe en disco ÚNICAMENTE los archivos que cambiaron (unos pocos KB)
    y reinicia la ventana en menos de 1 segundo sin reinstalar el .exe.
    """
    def _worker() -> None:
        try:
            patch_dir = obtener_directorio_parches()
            src_dir = os.path.join(patch_dir, "src")
            os.makedirs(src_dir, exist_ok=True)

            for rel_path, raw_bytes in archivos_modificados.items():
                destino = os.path.join(patch_dir, rel_path.replace("/", os.sep))
                os.makedirs(os.path.dirname(destino), exist_ok=True)
                with open(destino, "wb") as f_out:
                    f_out.write(raw_bytes)

            def _notificar_y_reiniciar() -> None:
                messagebox.showinfo(
                    "Actualización Completada",
                    f"Se actualizaron únicamente los {len(archivos_modificados)} módulo(s) modificados "
                    f"a la versión v{nueva_version} (sin reinstalar el programa).\n\n"
                    "La aplicación se reiniciará ahora para aplicar los cambios.",
                )
                if getattr(sys, "frozen", False):
                    subprocess.Popen([sys.executable])
                else:
                    subprocess.Popen([sys.executable] + sys.argv)
                app_root.destroy()

            app_root.after(0, _notificar_y_reiniciar)
        except Exception as exc:
            app_root.after(
                0,
                lambda: messagebox.showerror(
                    "Error al actualizar",
                    f"No se pudieron aplicar los cambios incrementales:\n{exc}",
                ),
            )

    threading.Thread(target=_worker, daemon=True).start()


def _verificar_parche_incremental_github(
    version_local: str,
) -> Optional[Tuple[str, Dict[str, bytes]]]:
    """
    Consulta la rama principal (main/master) en GitHub y compara cada archivo de src/
    contra la copia local. Devuelve solo los archivos que realmente cambiaron.
    """
    for branch in ("main", "master"):
        raw_base = f"https://raw.githubusercontent.com/{GITHUB_REPO}/{branch}"
        try:
            resp_cfg = requests.get(f"{raw_base}/src/config.py", timeout=5)
            if resp_cfg.status_code != 200:
                continue

            texto_cfg = resp_cfg.text
            m_ver = re.search(
                r'APP_VERSION\s*=\s*["\']([^"\']+)["\']', texto_cfg
            )
            if not m_ver:
                continue

            version_remota = m_ver.group(1).strip()
            if parse_version_tuple(version_remota) <= parse_version_tuple(
                version_local
            ):
                return None

            # Descargar y comparar hash SHA-256 de cada módulo para bajar solo lo que cambió
            patch_dir = obtener_directorio_parches()
            archivos_cambiados: Dict[str, bytes] = {}

            for rel_path in MODULOS_ACTUALIZABLES:
                resp_mod = requests.get(f"{raw_base}/{rel_path}", timeout=5)
                if resp_mod.status_code != 200:
                    continue

                bytes_remotos = resp_mod.content
                hash_remoto = _calcular_sha256_texto(bytes_remotos)

                # Comparar con archivo local actual
                ruta_parche_existente = os.path.join(
                    patch_dir, rel_path.replace("/", os.sep)
                )
                ruta_original = os.path.join(
                    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                    rel_path.replace("/", os.sep),
                )

                hash_local = ""
                for candidata in (ruta_parche_existente, ruta_original):
                    if os.path.isfile(candidata):
                        try:
                            with open(candidata, "rb") as f_in:
                                hash_local = _calcular_sha256_texto(f_in.read())
                            break
                        except Exception:
                            hash_local = ""

                if hash_remoto != hash_local:
                    archivos_cambiados[rel_path] = bytes_remotos

            if archivos_cambiados:
                return (version_remota, archivos_cambiados)
        except Exception:
            continue

    return None


def _download_and_apply_full_exe_update(
    download_url: str, app_root: tk.Misc
) -> None:
    def _worker() -> None:
        try:
            if getattr(sys, "frozen", False):
                current_exe = os.path.abspath(sys.executable)
            else:
                current_exe = os.path.abspath(sys.argv[0])

            exe_dir = os.path.dirname(current_exe)
            temp_exe = os.path.join(exe_dir, "update_temp.exe")
            updater_bat = os.path.join(exe_dir, "updater.bat")

            response = requests.get(download_url, stream=True, timeout=30)
            response.raise_for_status()

            with open(temp_exe, "wb") as f_out:
                for chunk in response.iter_content(chunk_size=65536):
                    if chunk:
                        f_out.write(chunk)

            bat_content = (
                "@echo off\r\n"
                "timeout /t 2 /nobreak > NUL\r\n"
                f'move /y "{temp_exe}" "{current_exe}" > NUL\r\n'
                f'start "" "{current_exe}"\r\n'
                'del "%~f0"\r\n'
            )
            with open(updater_bat, "w", encoding="utf-8") as f_bat:
                f_bat.write(bat_content)

            creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
            subprocess.Popen(
                ["cmd.exe", "/c", updater_bat],
                cwd=exe_dir,
                creationflags=creation_flags,
            )
            app_root.after(0, app_root.destroy)
        except Exception as exc:
            app_root.after(
                0,
                lambda: messagebox.showerror(
                    "Error de actualización",
                    f"No se pudo completar la descarga de la actualización:\n{exc}",
                ),
            )

    threading.Thread(target=_worker, daemon=True).start()


def verificar_actualizaciones_async(app_root: tk.Misc) -> None:
    """
    1. Primero intenta la Actualización Incremental Inteligente (compara SHA-256 de src/
       en GitHub y descarga solo los pocos KB de los archivos .py que cambiaron, sin reinstalar).
    2. Como respaldo, si existe un Release con .exe nuevo, también lo soporta.
    """
    def _check() -> None:
        if not GITHUB_REPO or GITHUB_REPO == "TU_USUARIO/TU_REPO":
            return

        version_actual = _obtener_version_local_efectiva()

        # Paso 1: Verificar actualización incremental ligera (solo archivos modificados)
        parche_info = _verificar_parche_incremental_github(version_actual)
        if parche_info is not None:
            nueva_ver, archivos_modificados = parche_info
            cant_archivos = len(archivos_modificados)
            tamano_kb = sum(len(b) for b in archivos_modificados.values()) / 1024.0
            nombres_archivos: List[str] = [
                os.path.basename(k) for k in archivos_modificados.keys()
            ]
            lista_str = ", ".join(nombres_archivos)

            def _preguntar_parche() -> None:
                acepta = messagebox.askyesno(
                    "Actualización Rápida Disponible",
                    f"Hay una mejora disponible (v{nueva_ver}).\n\n"
                    f"• Archivos modificados ({cant_archivos}): {lista_str}\n"
                    f"• Descarga ultra-rápida: {tamano_kb:.1f} KB (no requiere reinstalar el programa)\n\n"
                    "¿Desea aplicar los cambios ahora?",
                )
                if acepta:
                    _aplicar_actualizacion_incremental(
                        archivos_modificados, nueva_ver, app_root
                    )

            app_root.after(0, _preguntar_parche)
            return

        # Paso 2: Respaldo por GitHub Releases (.exe completo solo si fuera necesario)
        api_url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        try:
            response = requests.get(
                api_url,
                headers={"Accept": "application/vnd.github+json"},
                timeout=5,
            )
            if response.status_code != 200:
                return

            data = response.json()
            remote_tag = str(data.get("tag_name", "")).strip()
            if not remote_tag:
                return

            if parse_version_tuple(remote_tag) <= parse_version_tuple(
                version_actual
            ):
                return

            assets = data.get("assets", [])
            exe_asset_url: Optional[str] = None
            for asset in assets:
                name = str(asset.get("name", "")).lower()
                url = asset.get("browser_download_url")
                if name.endswith(".exe") and url:
                    exe_asset_url = str(url)
                    break

            if not exe_asset_url:
                return

            clean_ver = remote_tag.lstrip("vV")

            def _prompt_user() -> None:
                acepta = messagebox.askyesno(
                    "Actualización Disponible",
                    f"Se ha detectado una nueva versión disponible (v{clean_ver}).\n"
                    f"Versión actual: v{version_actual}\n\n"
                    "¿Desea descargar e instalar la actualización ahora?",
                )
                if acepta:
                    _download_and_apply_full_exe_update(exe_asset_url, app_root)

            app_root.after(0, _prompt_user)
        except Exception:
            return

    threading.Thread(target=_check, daemon=True).start()
