"""Lectura de los Excel con los elementos del pedido.

Cada tipo (Cutting, Legajo, Corona...) se exporta en un Excel distinto, con
una fila por elemento, y esos archivos son lo único que define qué lleva el
remito. El tipo sale del prefijo de los códigos de barras y acá se comprueba
que el archivo tenga las columnas de config.CATEGORIAS. Las columnas solas no
alcanzan para reconocerlo porque hay tipos con las mismas columnas.
"""

import os
from collections import Counter

from openpyxl import load_workbook

from config import CATEGORIAS, TIPOS_ELEMENTOS
from excel.lector_excel import hoja_activa
from utilidades.texto import a_texto, normalizar


def _buscar_columnas(categoria, encabezados):
    """Devuelve ({campo: índice de columna}, [títulos de las columnas que faltan])."""
    posiciones = {}
    for indice, nombre in enumerate(encabezados):
        if nombre is not None:
            posiciones.setdefault(normalizar(nombre), indice)

    indices, faltantes = {}, []
    for columna in categoria.columnas:
        indice = next((posiciones[normalizar(n)] for n in columna.encabezados if normalizar(n) in posiciones), None)
        if indice is None:
            faltantes.append(columna.titulo)
        else:
            indices[columna.campo] = indice
    return indices, faltantes


def _error_sin_codigos(tipo, otros_prefijos):
    """Mensaje para un archivo sin ningún código del tipo elegido."""
    esperado = TIPOS_ELEMENTOS[tipo]
    conocidos = [prefijo for prefijo, _ in otros_prefijos.most_common() if prefijo in TIPOS_ELEMENTOS]
    if conocidos:
        return ValueError(f"El archivo seleccionado corresponde a {TIPOS_ELEMENTOS[conocidos[0]]}, "
                          f"pero se esperaba {esperado}.")
    return ValueError(f"El Excel seleccionado no contiene códigos de barras de {esperado} (prefijo {tipo}).")


def _clave_codigo(elemento):
    # Los códigos sin número quedan al final, en el orden en que llegaron.
    numero = elemento["codigo_barras"][2:]
    return int(numero) if numero.isdigit() else float("inf")


def detectar_tipo(ruta_archivo):
    """Tipo de un Excel de elementos según el prefijo de su primer código de barras."""
    encabezados_codigo = {normalizar(n) for c in CATEGORIAS.values() for n in c.columnas[0].encabezados}

    libro = load_workbook(ruta_archivo, read_only=True, data_only=True)
    try:
        filas = hoja_activa(libro).iter_rows(values_only=True)
        encabezados = [normalizar(nombre) for nombre in next(filas, ())]
        indice = next((i for i, nombre in enumerate(encabezados) if nombre in encabezados_codigo), None)
        if indice is None:
            raise ValueError("El Excel no tiene una columna de código de barras.")

        codigo = next((a_texto(fila[indice]) for fila in filas if indice < len(fila) and fila[indice]), "")
    finally:
        libro.close()

    tipo = codigo[:2].upper()
    if tipo not in CATEGORIAS:
        raise ValueError(f"No se reconoce el tipo de elemento del archivo (código de ejemplo: {codigo or 'ninguno'}).")
    return tipo


def leer_datos_elementos(ruta_archivo, tipo):
    """Lee el Excel de elementos de un tipo. Las filas de otros tipos se ignoran.

    Retorna:
        {
            "tipo": "CO",
            "ruta": ruta_archivo,
            "elementos": [{campo: valor}, ...] ordenados por número de código,
            "duplicados": códigos repetidos (se conserva la primera fila),
        }
    """
    if not os.path.isfile(ruta_archivo):
        raise FileNotFoundError(f"No se encontró el archivo: {ruta_archivo}")
    if os.path.splitext(ruta_archivo)[1].lower() != ".xlsx":
        raise ValueError("El archivo de datos de elementos debe ser un archivo .xlsx.")

    libro = load_workbook(ruta_archivo, read_only=True, data_only=True)
    try:
        filas = hoja_activa(libro).iter_rows(values_only=True)
        encabezados = next(filas, None)
        if not encabezados:
            raise ValueError("El Excel de elementos está vacío.")

        indices, faltantes = _buscar_columnas(CATEGORIAS[tipo], encabezados)
        if faltantes:
            raise ValueError(f"Al Excel de {TIPOS_ELEMENTOS[tipo]} le faltan las columnas: {', '.join(faltantes)}.")

        por_codigo, duplicados, otros_prefijos = {}, [], Counter()
        for fila in filas:
            if not fila or all(valor is None for valor in fila):
                continue

            registro = {campo: (fila[i] if i < len(fila) else None) for campo, i in indices.items()}
            codigo = a_texto(registro["codigo_barras"])
            # Sin código de barras la fila no se puede asociar al pedido.
            if not codigo:
                continue

            clave = codigo.upper()
            if not clave.startswith(tipo):
                otros_prefijos[clave[:2]] += 1
                continue

            registro["codigo_barras"] = codigo
            if clave in por_codigo:
                duplicados.append(codigo)
            else:
                por_codigo[clave] = registro
    finally:
        libro.close()

    if not por_codigo:
        raise _error_sin_codigos(tipo, otros_prefijos)

    elementos = sorted(por_codigo.values(), key=_clave_codigo)
    return {"tipo": tipo, "ruta": ruta_archivo, "elementos": elementos, "duplicados": duplicados}
