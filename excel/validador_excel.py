"""Validaciones de negocio sobre el pedido ya leído.

La existencia de las columnas la controla leer_excel(); acá se revisa que
los datos tengan sentido para poder armar un remito.
"""

from utilidades.texto import a_texto


def validar_pedido(datos):
    """Devuelve la lista de problemas del pedido. Lista vacía = pedido válido."""
    errores = []
    if not a_texto(datos.get("ID Prestamo")):
        errores.append("El ID del préstamo está vacío.")
    return errores
