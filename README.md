# Procesador de Fichadas - Leal Sistemas

Aplicación de escritorio en Python (Windows) para limpiar, normalizar y organizar los reportes biométricos mensuales exportados por **Leal Sistemas**.

## Estructura del Proyecto (Qué subir a GitHub)

Solo estos archivos forman parte del código fuente del programa:

```text
ProcesadorFichadas/
├── main.py               # Punto de entrada y cargador de parches incrementales
├── requirements.txt      # Librerías base (customtkinter, pandas, openpyxl, xlrd, tkinterdnd2, requests)
├── .gitignore            # Evita que Git suba archivos pesados (.exe, .xlsx, carpetas temporales)
├── vercel.json           # Configuración automática para desplegar la web de docs/ en Vercel
├── README.md             # Documentación del proyecto
├── docs/
│   └── index.html        # Página web oficial de descarga (lista para Vercel o GitHub Pages)
└── src/
    ├── __init__.py
    ├── config.py         # Versión de la app (APP_VERSION), repositorio (GITHUB_REPO) y paleta de colores
    ├── processor.py      # Lógica de lectura (.xls, HTML, CSV), limpieza de horas y cálculo de resúmenes
    ├── exporter.py       # Generación del reporte Excel gerencial de 3 hojas con openpyxl
    ├── ui.py             # Interfaz gráfica (CustomTkinter + Drag & Drop + Filtros + Buscador)
    └── updater.py        # Motor de Actualización Incremental (Delta Update) y respaldo OTA
```

---

## Cómo funciona la Actualización Incremental (Sin Reinstalar Todo)

El ejecutable `ProcesadorFichadas.exe` (`50 MB`) ya lleva adentro todas las librerías pesadas (`pandas`, `openpyxl`, `customtkinter`, `xlrd`, `tkinterdnd2`) que **nunca cambian**.

Lo que cambia cuando haces una mejora (cambiar una columna, un cálculo o un botón) son únicamente los archivos livianos de la carpeta `src/` (que pesan apenas **10 a 30 KB**).

Por eso, [`src/updater.py`](src/updater.py) y [`main.py`](main.py) implementan **Actualización Incremental por Archivo (Delta Update)**:
1. Al abrir el programa, verifica `src/config.py` directamente en tu rama `main` de GitHub.
2. Si ve que subiste `APP_VERSION` (por ejemplo de `"1.1.0"` a `"1.2.0"`), **compara el hash SHA-256 de cada archivo de `src/`** y detecta exactamente cuál cambió (por ejemplo, solo `exporter.py`).
3. Deja intacto el `ProcesadorFichadas.exe` y todos los archivos que no cambiaron, y **descarga únicamente el archivo modificado (`~6 KB`, tarda 0.2 segundos)** guardándolo en `%LOCALAPPDATA%\ProcesadorFichadas_Leal\live_updates`.
4. Reinicia la ventana en 1 segundo con el cambio ya activo, **sin reinstalar ni volver a bajar los 50 MB**.

---

## Paso a Paso para usar las Actualizaciones Incrementales

### 1. Configuración inicial (solo la primera vez antes de mandarle el `.exe` a tu jefa):
1. Crea un repositorio **Público** en GitHub.
2. En [`src/config.py`](src/config.py), el repositorio oficial ya se encuentra configurado como:
   ```python
   APP_VERSION = "1.1.0"
   GITHUB_REPO = "Vanessaalberti/ProcesadorFichadas"
   ```
3. Compila el `.exe` una sola vez (`pyinstaller --noconfirm ProcesadorFichadas.spec`) y envíaselo a tu jefa.

### 2. Cada vez que quieras hacer un cambio en el futuro:
1. Modificas el archivo que necesites dentro de `src/` (por ejemplo `src/exporter.py` o `src/ui.py`).
2. Subes el número de `APP_VERSION` en [`src/config.py`](src/config.py) (por ejemplo a `"1.2.0"`).
3. Subes los cambios a tu repositorio de GitHub (`git push` o subiendo el archivo editado a la web de GitHub). **¡Ni siquiera necesitas volver a compilar el `.exe`!**
4. Cuando tu jefa abra su programa, le avisará: *"Hay una mejora disponible (v1.2.0). Archivos modificados (1): exporter.py (6.2 KB)"*, y en medio segundo tendrá el cambio funcionando.
