"""Preferencias del usuario que se conservan entre ejecuciones.

Por ahora solo la carpeta donde se guardan los remitos. Se guarda en un
config.json dentro de la carpeta de datos del usuario.
"""

import json
import logging
import os

from utilidades.rutas import carpeta_datos_usuario, carpeta_escritorio

log = logging.getLogger(__name__)


def _ruta_archivo():
    return os.path.join(carpeta_datos_usuario(), "config.json")


def _leer():
    try:
        with open(_ruta_archivo(), encoding="utf-8") as archivo:
            return json.load(archivo)
    except FileNotFoundError:
        return {}
    except (OSError, json.JSONDecodeError) as error:
        log.warning("No se pudo leer la configuración: %s", error)
        return {}


def cargar_carpeta_remitos():
    """Última carpeta elegida para los remitos; si no hay ninguna, el Escritorio."""
    carpeta = _leer().get("carpeta_remitos")
    return os.path.abspath(carpeta) if carpeta else carpeta_escritorio()


def guardar_carpeta_remitos(carpeta):
    """Guarda la carpeta elegida. Si falla, la app sigue usándola en esta sesión."""
    configuracion = _leer()
    configuracion["carpeta_remitos"] = carpeta
    try:
        with open(_ruta_archivo(), "w", encoding="utf-8") as archivo:
            json.dump(configuracion, archivo, indent=4, ensure_ascii=False)
    except OSError as error:
        log.warning("No se pudo guardar la configuración: %s", error)
