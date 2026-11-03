"""Configuración del log de la aplicación.

El .exe se compila sin consola, así que los print() no los ve nadie. Todo lo
que haga falta para diagnosticar un problema va a un archivo de log en la
carpeta de datos del usuario (%APPDATA%/INFOLOG REMITOS/infolog_remitos.log).
"""

import logging
import os
from logging.handlers import RotatingFileHandler

from utilidades.rutas import carpeta_datos_usuario


def configurar_registro():
    """Activa el log en archivo (máximo 3 archivos de 1 MB) y lo devuelve."""
    ruta_log = os.path.join(carpeta_datos_usuario(), "infolog_remitos.log")
    manejador = RotatingFileHandler(ruta_log, maxBytes=1_000_000, backupCount=2, encoding="utf-8")
    manejador.setFormatter(logging.Formatter("%(asctime)s %(levelname)s [%(name)s] %(message)s"))

    raiz = logging.getLogger()
    raiz.setLevel(logging.INFO)
    raiz.addHandler(manejador)
    return ruta_log
