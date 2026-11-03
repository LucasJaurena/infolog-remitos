"""Punto de entrada de INFOLOG REMITOS."""

import logging

from utilidades.registro import configurar_registro
from interfaz.ventana_principal import App


if __name__ == "__main__":
    configurar_registro()
    logging.getLogger(__name__).info("Inicio de la aplicación")
    App().mainloop()
