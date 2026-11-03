"""Conversión de valores leídos de Excel a texto."""

import unicodedata
from datetime import date, datetime


def normalizar(valor):
    """Pasa un encabezado a minúsculas, sin tildes y sin espacios repetidos.

    Sirve para comparar nombres de columnas sin depender de cómo los haya
    escrito quien exportó el Excel ("Código de Barras" == "codigo de barras").
    """
    if valor is None:
        return ""
    texto = unicodedata.normalize("NFD", str(valor).strip().lower())
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return " ".join(texto.split())


def a_texto(valor, con_hora=False):
    """Convierte una celda de Excel a texto legible.

    - None queda como cadena vacía (no como "None").
    - Los números enteros guardados como float (1234.0) pierden el ".0".
    - Las fechas se muestran como dd/mm/aaaa. Con con_hora=True se agrega la
      hora, pero solo si la celda realmente tiene una distinta de 00:00.
    """
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    if isinstance(valor, datetime):
        if con_hora and (valor.hour or valor.minute):
            return valor.strftime("%d/%m/%Y %H:%M")
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, date):
        return valor.strftime("%d/%m/%Y")
    return str(valor).strip()


def plural(cantidad, palabra):
    """"1 elemento", "3 elementos". Alcanza para las palabras que usa la app."""
    return f"{cantidad} {palabra}" + ("" if cantidad == 1 else "s")
