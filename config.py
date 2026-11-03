"""Configuración general de INFOLOG REMITOS.

Acá vive todo lo que cambia con el negocio y no con el código: nombres de
los tipos de elemento, columnas que se leen de cada Excel, cómo se muestran
en el PDF y los colores de la interfaz.

Para agregar una categoría nueva alcanza con sumarla a CATEGORIAS; el
lector de Excel, la interfaz y el generador del remito la toman de acá.
"""

from dataclasses import dataclass


APP_NOMBRE = "INFOLOG REMITOS"
APP_VERSION = "1.1.0"
APP_AUTOR = "Lucas Jaurena"
APP_CONTACTO = "lujaurena_2@yahoo.com.ar"

ANCHO_VENTANA = 1000
ALTO_VENTANA = 650

COLOR_FONDO = "#F2F4F7"
COLOR_MENU = "#17375E"
COLOR_BOTON = "#2E75B6"
COLOR_SEPARADOR = "#5B8AC4"
COLOR_TEXTO_TITULO = "#17375E"

# "Cargar pedido" va en verde para distinguir la etapa de inicio. Es oscuro
# a propósito: con letras blancas se sigue leyendo bien.
COLOR_BOTON_PEDIDO = "#2E7D32"
COLOR_BOTON_PEDIDO_HOVER = "#1B5E20"

# "Carpeta de remitos" en un dorado oscuro, de la familia del amarillo del
# logo, para diferenciarlo de los pasos. Más claro ya no se lee en blanco.
COLOR_BOTON_CARPETA = "#8F6A00"
COLOR_BOTON_CARPETA_HOVER = "#735500"
COLOR_BOTON_QUITAR = "#B3261E"        # contraste con letras blancas: 6.5
COLOR_BOTON_QUITAR_HOVER = "#8C1D18"

# El botón de la pantalla donde está el usuario se marca con un borde del
# amarillo del logo. No se usa de fondo porque con letras blancas no se lee.
COLOR_RESALTADO = "#FFE12F"

# Etiquetas "PASO N" del menú según el avance del remito.
COLOR_PASO = "#8FB0D8"
COLOR_PASO_SIGUIENTE = "#FFFFFF"
COLOR_PASO_LISTO = "#7BD88F"


# Prefijos de los IDs que usa YPF. El prefijo son las dos primeras letras
# del ID (por ejemplo LE00001076 es un Legajo de pozo).
TIPOS_ELEMENTOS = {
    "LE": "Legajo de pozo",
    "PE": "Perfil de pozo",
    "CG": "Perfil de control geológico",
    "CO": "Corona",
    "CU": "Cutting",
    "CP": "Cutting pet headspace",
    "CL": "Cutting sin lavar",
    "TE": "Testigo lateral",
    "CI": "Cinta sísmica",
    "FM": "Línea sísmica",
    "IN": "Estudio e informe",
    "FL": "Fluido",
    "MS": "Muestra suelta",
    "AF": "Afloramiento",
    "CD": "Corte delgado",
    "PG": "Plug",
    "PA": "Partes de observador",
    "TO": "Taco de petrología orgánica",
}


@dataclass(frozen=True)
class Columna:
    """Una columna del Excel de elementos y su lugar en el PDF.

    encabezados: nombres aceptados en el Excel. Se comparan sin tildes ni
        mayúsculas, así que "Código de barras" también encuentra
        "CODIGO DE BARRAS".
    titulo: nombre de la columna en el PDF. Las columnas de ubicación del
        Excel traen la posición entre paréntesis ("Caja (8)"); en el remito
        van sin ese número.
    ancho: ancho relativo en el PDF; después se escala al ancho de la hoja.
    """

    campo: str
    encabezados: tuple
    titulo: str
    ancho: float


@dataclass(frozen=True)
class Categoria:
    """Tipo de elemento que ya se puede incluir en un remito.

    La clave en CATEGORIAS es el prefijo de sus códigos de barras: varios
    tipos tienen exactamente las mismas columnas (Perfiles, Cuttings...), así
    que lo único que los distingue son esas dos letras.
    """

    titulo_pdf: str
    columnas: tuple


CODIGO_BARRAS = Columna("codigo_barras", ("Código de barras",), "Código de barras", 30)
CAMPO = ("Campo",)
ZONA = ("Zona",)
OBSERVACIONES = ("Observaciones",)
NOMBRE_POZO = ("Nombre Pozo", "Nombre de Pozo")
CAJA = ("Caja (8)", "Cajón (8)")
CAJON = ("Cajón (8)", "Caja (8)")
LINEA = ("Línea",)
DE_PUNTO = ("De punto",)
A_PUNTO = ("A punto",)
TOPE_CARRERA = ("Tope Carrera (M)", "Tope de Carrera (M)")
BASE_CARRERA = ("Base Carrera (M)", "Base de Carrera (M)")


def _cuttings(titulo_pdf):
    """Cutting, Cutting sin lavar y Pet Headspace usan las mismas columnas."""
    return Categoria(
        titulo_pdf=titulo_pdf,
        columnas=(
            CODIGO_BARRAS,
            Columna("nombre_pozo", NOMBRE_POZO, "Nombre Pozo", 34),
            Columna("desde", ("Desde",), "Desde", 18),
            Columna("hasta", ("Hasta",), "Hasta", 18),
            Columna("campo", CAMPO, "Campo", 25),
            Columna("caja", CAJA, "Caja", 20),
            Columna("observaciones", OBSERVACIONES, "Observaciones", 45),
        ),
    )


def _perfiles(titulo_pdf):
    """Perfiles de pozo y de control geológico usan las mismas columnas."""
    return Categoria(
        titulo_pdf=titulo_pdf,
        columnas=(
            CODIGO_BARRAS,
            Columna("nombre_pozo", NOMBRE_POZO, "Nombre Pozo", 40),
            Columna("tope", ("Tope",), "Tope", 18),
            Columna("base", ("Base",), "Base", 18),
            Columna("campo", CAMPO, "Campo", 30),
            Columna("observaciones", OBSERVACIONES, "Observaciones", 50),
        ),
    )


# El orden de este diccionario es el orden de las secciones en el PDF.
CATEGORIAS = {
    "CU": _cuttings("CUTTINGS"),
    "LE": Categoria(
        titulo_pdf="LEGAJOS",
        columnas=(
            CODIGO_BARRAS,
            Columna("nombre_pozo", NOMBRE_POZO, "Nombre de Pozo", 45),
            Columna("campo", CAMPO, "Campo", 35),
            Columna("observaciones", OBSERVACIONES, "Observaciones", 70),
        ),
    ),
    "CO": Categoria(
        titulo_pdf="CORONAS",
        columnas=(
            CODIGO_BARRAS,
            Columna("nombre_pozo", NOMBRE_POZO, "Nombre Pozo", 32),
            Columna("campo", CAMPO, "Campo", 20),
            Columna("tope_carrera", TOPE_CARRERA, "Tope Carrera (M)", 20),
            Columna("base_carrera", BASE_CARRERA, "Base Carrera (M)", 20),
            Columna("carrera", ("Carrera",), "Carrera", 18),
            Columna("cajon", CAJON, "Cajón", 18),
            Columna("observaciones", OBSERVACIONES, "Observaciones", 32),
        ),
    ),
    "CI": Categoria(
        titulo_pdf="CINTAS SÍSMICAS",
        columnas=(
            CODIGO_BARRAS,
            Columna("linea", LINEA, "Línea", 30),
            Columna("de_punto", DE_PUNTO, "De punto", 18),
            Columna("a_punto", A_PUNTO, "A punto", 18),
            Columna("zona", ZONA, "Zona", 25),
            Columna("categoria", ("Categoría",), "Categoría", 25),
            Columna("observaciones", OBSERVACIONES, "Observaciones", 45),
        ),
    ),
    "FM": Categoria(
        titulo_pdf="LÍNEAS SÍSMICAS",
        columnas=(
            CODIGO_BARRAS,
            Columna("linea", LINEA, "Linea", 35),
            Columna("de_punto", DE_PUNTO, "De Punto", 20),
            Columna("a_punto", A_PUNTO, "A Punto", 20),
            Columna("zona", ZONA, "Zona", 30),
            Columna("observaciones", OBSERVACIONES, "Observaciones", 50),
        ),
    ),
    "PA": Categoria(
        titulo_pdf="PARTES DE OBSERVADOR",
        columnas=(
            CODIGO_BARRAS,
            Columna("linea", LINEA, "Linea", 28),
            Columna("de_punto", DE_PUNTO, "De Punto", 17),
            Columna("a_punto", A_PUNTO, "A Punto", 17),
            Columna("descripcion", ("Descripción",), "Descripción", 40),
            Columna("zona", ZONA, "Zona", 25),
            Columna("observaciones", OBSERVACIONES, "Observaciones", 40),
        ),
    ),
    "IN": Categoria(
        titulo_pdf="ESTUDIOS E INFORMES",
        columnas=(
            CODIGO_BARRAS,
            Columna("nombre_pozo", NOMBRE_POZO, "Nombre Pozo", 35),
            Columna("titulo", ("Título",), "Titulo", 55),
            Columna("anio", ("Año",), "Año", 14),
            Columna("zona", ZONA, "Zona", 25),
            Columna("observaciones", OBSERVACIONES, "Observaciones", 40),
        ),
    ),
    "CG": _perfiles("PERFILES DE CONTROL GEOLÓGICO"),
    "PE": _perfiles("PERFILES DE POZOS"),
    "FL": Categoria(
        titulo_pdf="FLUIDOS",
        columnas=(
            CODIGO_BARRAS,
            Columna("nombre_pozo", NOMBRE_POZO, "Nombre Pozo", 40),
            Columna("tipo_muestra", ("Tipo De Muestra",), "Tipo De Muestra", 30),
            Columna("campo", CAMPO, "Campo", 30),
            Columna("observaciones", OBSERVACIONES, "Observaciones", 55),
        ),
    ),
    "MS": Categoria(
        titulo_pdf="MUESTRAS SUELTAS",
        columnas=(
            CODIGO_BARRAS,
            Columna("nombre_pozo", NOMBRE_POZO, "Nombre Pozo", 34),
            Columna("base_carrera", BASE_CARRERA, "Base Carrera (M)", 20),
            Columna("tope_carrera", TOPE_CARRERA, "Tope Carrera (M)", 20),
            Columna("campo", CAMPO, "Campo", 25),
            Columna("cajon", CAJON, "Cajón", 18),
            Columna("observaciones", OBSERVACIONES, "Observaciones", 42),
        ),
    ),
    "AF": Categoria(
        titulo_pdf="AFLORAMIENTOS",
        columnas=(
            CODIGO_BARRAS,
            Columna("nombre", ("Nombre",), "Nombre", 35),
            Columna("formacion", ("Formación",), "Formacion", 28),
            Columna("cuenca", ("Cuenca",), "Cuenca", 25),
            Columna("contenedor", ("Contenedor (8)",), "Contenedor", 22),
            Columna("observaciones", OBSERVACIONES, "Observaciones", 45),
        ),
    ),
    "CD": Categoria(
        titulo_pdf="CORTES DELGADOS",
        columnas=(
            CODIGO_BARRAS,
            Columna("nombre_pozo", NOMBRE_POZO, "Nombre Pozo", 34),
            Columna("corte", ("Corte",), "Corte", 18),
            Columna("profundidad", ("Profundidad",), "Profundidad", 20),
            Columna("campo", CAMPO, "Campo", 25),
            Columna("caja", CAJA, "Caja", 18),
            Columna("observaciones", OBSERVACIONES, "Observaciones", 43),
        ),
    ),
    "CP": _cuttings("CUTTING PET HEADSPACE"),
    "CL": _cuttings("CUTTING SIN LAVAR"),
    "PG": Categoria(
        titulo_pdf="PLUGS",
        columnas=(
            CODIGO_BARRAS,
            Columna("nombre_pozo", NOMBRE_POZO, "Nombre Pozo", 36),
            Columna("numero_identificador", ("Numero Identificador",), "Numero Identificador", 24),
            Columna("zona", ZONA, "Zona", 25),
            Columna("caja_bolsa", ("Caja_Bolsa (8)",), "Caja_Bolsa", 20),
            Columna("observaciones", OBSERVACIONES, "Observaciones", 45),
        ),
    ),
    "TE": Categoria(
        titulo_pdf="TESTIGOS LATERALES",
        columnas=(
            CODIGO_BARRAS,
            Columna("nombre_pozo", NOMBRE_POZO, "Nombre Pozo", 40),
            Columna("zona", ZONA, "Zona", 30),
            Columna("caja", CAJA, "Caja", 20),
            Columna("observaciones", OBSERVACIONES, "Observaciones", 55),
        ),
    ),
    "TO": Categoria(
        titulo_pdf="TACOS DE PETROLOGÍA ORGÁNICA",
        columnas=(
            # Es el único Excel que trae este encabezado en singular y mayúsculas.
            Columna("codigo_barras", ("CODIGO DE BARRA", "Código de barras"), "CODIGO DE BARRA", 30),
            Columna("nombre_pozo", NOMBRE_POZO, "Nombre Pozo", 40),
            Columna("campo", CAMPO, "Campo", 30),
            Columna("cajon", CAJON, "Cajón", 20),
            Columna("observaciones", OBSERVACIONES, "Observaciones", 55),
        ),
    ),
}
