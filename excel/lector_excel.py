"""Lectura del Excel del pedido exportado desde el sistema de YPF.

El Excel trae una sola fila de datos: la fila 1 son los encabezados y la
fila 2 es el pedido. Los encabezados se buscan sin importar tildes,
mayúsculas ni espacios de más.
"""

import os

import openpyxl

from utilidades.texto import normalizar

COLUMNAS_REQUERIDAS = (
    "ID Prestamo",
    "Fecha Solicitud",
    "Usuario Solicitante",
    "Nombre Destinatario",
    "Compañía Destinataria",
    "Dirección Destinatario",
    "Email Destinatario",
    "Teléfono Destinatario",
    "Nombre Solicitante",
    "Compañía Solicitante",
    "Dirección Solicitante",
    "Email Solicitante",
    "Teléfono Solicitante",
    "Notas Solicitante",
    "Código Estatus",
    "Detalles Estatus (ES)",
    "Tipo Entrega",
    "Fecha Visita Aprob.",
    "Fecha Aprob. Retorno",
)


def hoja_activa(libro):
    """Hoja activa de un libro abierto con read_only, lista para recorrer.

    Hay Excel exportados de YPF (los más viejos) que declaran mal su tamaño,
    como si tuvieran solo la celda A1, y en read_only openpyxl lee solo esa
    celda de cada fila. reset_dimensions() hace que lea las filas completas.
    """
    hoja = libro.active
    hoja.reset_dimensions()
    return hoja


def leer_excel(ruta_archivo):
    """Lee el pedido y lo devuelve como diccionario {columna: valor}.

    Las claves son siempre los nombres de COLUMNAS_REQUERIDAS, aunque en el
    Excel estén escritos con otras mayúsculas o sin tildes. Las celdas vacías
    quedan como None.
    """
    if not ruta_archivo:
        raise ValueError("No se indicó un archivo Excel.")
    if os.path.splitext(ruta_archivo)[1].lower() != ".xlsx":
        raise ValueError("El pedido debe ser un archivo .xlsx (Excel 2007 o posterior).")

    try:
        libro = openpyxl.load_workbook(ruta_archivo, read_only=True, data_only=True)
    except Exception as error:
        raise ValueError(f"No se pudo abrir el archivo Excel: {error}") from error

    try:
        filas = hoja_activa(libro).iter_rows(min_row=1, max_row=2, values_only=True)
        encabezados = next(filas, None)
        fila = next(filas, None)
    finally:
        libro.close()

    if not encabezados:
        raise ValueError("El archivo Excel no contiene encabezados.")

    posiciones = {normalizar(nombre): i for i, nombre in enumerate(encabezados) if nombre is not None}
    faltantes = [c for c in COLUMNAS_REQUERIDAS if normalizar(c) not in posiciones]
    if faltantes:
        raise ValueError(
            "El archivo Excel no contiene las columnas requeridas:\n\n- " + "\n- ".join(faltantes)
        )

    if fila is None or all(valor is None for valor in fila):
        raise ValueError("El archivo Excel no contiene datos del pedido.")

    datos = {}
    for columna in COLUMNAS_REQUERIDAS:
        indice = posiciones[normalizar(columna)]
        datos[columna] = fila[indice] if indice < len(fila) else None
    return datos
