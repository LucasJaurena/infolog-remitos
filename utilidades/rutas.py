"""Rutas de recursos y carpetas del usuario.

Cuando la app corre como .exe generado con PyInstaller (--onefile), los
archivos empaquetados se descomprimen en una carpeta temporal que indica
sys._MEIPASS. Por eso los recursos se buscan siempre con ruta_recurso() y
nunca a partir de __file__.
"""

import os
import sys

RAIZ_PROYECTO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
NOMBRE_CARPETA_DATOS = "INFOLOG REMITOS"


def ruta_recurso(*partes):
    """Ruta a un archivo dentro de recursos/, tanto en desarrollo como en el .exe."""
    base = getattr(sys, "_MEIPASS", RAIZ_PROYECTO)
    return os.path.join(base, "recursos", *partes)


def carpeta_datos_usuario():
    """Carpeta donde se guardan la configuración y el log (%APPDATA% en Windows)."""
    base = os.environ.get("APPDATA") if os.name == "nt" else None
    if base:
        carpeta = os.path.join(base, NOMBRE_CARPETA_DATOS)
    else:
        carpeta = os.path.join(os.path.expanduser("~"), ".infolog_remitos")
    os.makedirs(carpeta, exist_ok=True)
    return carpeta


def carpeta_escritorio():
    """Escritorio real del usuario.

    Se le pregunta a Windows en lugar de armar "~/Desktop" porque con OneDrive
    el Escritorio suele estar redirigido (por ejemplo a OneDrive/Escritorio).
    """
    if os.name == "nt":
        try:
            import ctypes
            from ctypes import wintypes

            CSIDL_DESKTOPDIRECTORY = 0x0010
            ruta = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
            if ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_DESKTOPDIRECTORY, None, 0, ruta) == 0:
                return ruta.value
        except (OSError, AttributeError):
            pass
    return os.path.join(os.path.expanduser("~"), "Desktop")
