"""Generación del remito en PDF (A4) con ReportLab.

El remito se arma con los datos que ya están en memoria: la fila del pedido
(leer_excel) y los elementos de cada tipo (leer_datos_elementos). Este módulo
no lee archivos Excel.

Estructura del PDF:
    1. Encabezado azul con logos, número de remito y fechas.
    2. Tipos de elementos incluidos.
    3. Datos de contacto y fechas de retiro/devolución.
    4. Notas del solicitante.
    5. Una tabla por tipo de elemento (en el orden de config.CATEGORIAS), cada
       una en hoja nueva.
    6. Espacio para los datos y la firma de quien recibe, sin partirse.
"""

import os
import re
from datetime import datetime
from html import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    BaseDocTemplate, CondPageBreak, Frame, Image, KeepTogether, PageBreak, PageTemplate, Paragraph, Spacer, Table,
    TableStyle,
)

from config import APP_NOMBRE, CATEGORIAS
from utilidades.rutas import ruta_recurso
from utilidades.texto import a_texto, plural

MARGEN = 10 * mm
ANCHO_BLOQUES = A4[0] - 2 * MARGEN  # 190 mm: todos los bloques ocupan el mismo ancho

COLOR_AZUL = colors.HexColor("#17375E")
COLOR_AZUL_CLARO = colors.HexColor("#D9E7F5")
COLOR_AZUL_MUY_CLARO = colors.HexColor("#EEF4FA")
COLOR_VERDE_CLARO = colors.HexColor("#E5F0E0")
COLOR_NARANJA_CLARO = colors.HexColor("#F8D8B5")
COLOR_NARANJA_MUY_CLARO = colors.HexColor("#FCEBD8")
COLOR_BORDE = colors.HexColor("#8A8A8A")


class NumeradoCanvas(canvas.Canvas):
    """Canvas que escribe "Página N de M" en el pie.

    El total de páginas recién se conoce al final, así que se guarda el
    estado de cada página y se dibujan todas juntas en save().
    """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._paginas = []

    def showPage(self):
        self._paginas.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._paginas)
        for estado in self._paginas:
            self.__dict__.update(estado)
            self.setFont("Helvetica-Bold", 7.5)
            self.setFillColor(COLOR_AZUL)
            self.drawRightString(A4[0] - MARGEN, 5 * mm, f"Página {self._pageNumber} de {total}")
            super().showPage()
        super().save()


# ============================================================
# Utilidades
# ============================================================

def nombre_archivo_remito(id_pedido):
    """Nombre del PDF a partir del ID, sin caracteres que Windows no acepta."""
    nombre = re.sub(r'[\\/:*?"<>|]', "_", a_texto(id_pedido)) or "PEDIDO_SIN_ID"
    return f"{nombre}_Remito.pdf"


def _parrafo(valor, estilo, vacio="-"):
    """Paragraph a partir de un dato del Excel.

    ReportLab interpreta el texto como marcado tipo HTML: sin escapar, un
    "<urgente>" en las notas desaparece del PDF y los saltos de línea se
    pierden. Por eso todo dato externo pasa por acá.
    """
    texto = a_texto(valor) or vacio
    return Paragraph(escape(texto).replace("\n", "<br/>"), estilo)


def _estilo_tabla(*comandos, padding=2):
    return TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), padding),
        ("RIGHTPADDING", (0, 0), (-1, -1), padding),
        ("TOPPADDING", (0, 0), (-1, -1), padding),
        ("BOTTOMPADDING", (0, 0), (-1, -1), padding),
        *comandos,
    ])


def _franja(contenido, fondo, alineacion_vertical="MIDDLE", **opciones):
    """Bloque de una sola celda del ancho de la hoja (títulos y notas)."""
    tabla = Table([[contenido]], colWidths=[ANCHO_BLOQUES], **opciones)
    tabla.setStyle(_estilo_tabla(
        ("BACKGROUND", (0, 0), (-1, -1), fondo),
        ("BOX", (0, 0), (-1, -1), 0.5, COLOR_BORDE),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("VALIGN", (0, 0), (-1, -1), alineacion_vertical),
        padding=3,
    ))
    return tabla


def _logo(nombre_archivo, ancho, alto, texto_alternativo, estilo):
    ruta = ruta_recurso("imagenes", nombre_archivo)
    if os.path.isfile(ruta):
        return Image(ruta, width=ancho, height=alto, kind="proportional")
    return Paragraph(texto_alternativo, estilo)


# ============================================================
# Validación
# ============================================================

def validar_datos_remito(datos_pedido, grupos):
    """Lista de motivos por los que todavía no se puede generar el remito.

    Lista vacía = se puede generar. La interfaz usa esta misma función para
    habilitar el botón, así las reglas están en un solo lugar.
    """
    if not isinstance(datos_pedido, dict) or not isinstance(grupos, dict):
        return ["No se recibieron correctamente los datos del pedido."]

    errores = []
    if not a_texto(datos_pedido.get("ID Prestamo")):
        errores.append("El pedido no tiene ID Prestamo.")
    if not any(grupos.values()):
        errores.append("Todavía no se cargaron elementos.")
    return errores


# ============================================================
# Estilos
# ============================================================

def _crear_estilos():
    def estilo(nombre, fuente="Helvetica", tamano=9, interlineado=10.5, alineacion=TA_LEFT, color=colors.black):
        return ParagraphStyle(nombre, fontName=fuente, fontSize=tamano, leading=interlineado,
                              alignment=alineacion, textColor=color)

    return {
        "titulo": estilo("Titulo", "Helvetica-Bold", 21, 24, TA_CENTER, colors.white),
        "seccion": estilo("Seccion", "Helvetica-Bold", 11.5, 13, TA_CENTER, COLOR_AZUL),
        "cabecera_tabla": estilo("CabeceraTabla", "Helvetica-Bold", 7.8, 9, TA_CENTER),
        "celda": estilo("Celda", tamano=7.6, interlineado=8.8),
        "dato": estilo("Dato"),
        "dato_negrita": estilo("DatoNegrita", "Helvetica-Bold", 9.5, 11),
        "titulo_bloque": estilo("TituloBloque", "Helvetica-Bold", 12.5, 14, TA_CENTER, COLOR_AZUL),
        "fecha_impresion": estilo("FechaImpresion", "Helvetica-Bold", 9.3, 12.5, TA_CENTER, colors.white),
        "fecha_pedido": estilo("FechaPedido", "Helvetica-Bold", 8.5, 10.5, TA_CENTER, colors.white),
    }


# ============================================================
# Bloques del remito
# ============================================================

def _crear_encabezado(datos, estilos):
    id_pedido = a_texto(datos.get("ID Prestamo"))
    logo_infolog = _logo("logo_infolog.png", 32 * mm, 28 * mm, "INFOLOG", estilos["dato_negrita"])
    # El logo del cliente es opcional: el repositorio público no lo incluye por ser
    # una marca registrada, y sin el archivo ese lugar queda en blanco.
    logo_ypf = _logo("logo_ypf.png", 28 * mm, 13 * mm, "", estilos["dato_negrita"])
    fecha_pedido = a_texto(datos.get("Fecha Solicitud")) or "-"

    bloque_derecho = Table(
        [
            [logo_ypf],
            [Paragraph(f"Fecha impresión remito: {datetime.now():%d/%m/%Y}", estilos["fecha_impresion"])],
            [Paragraph(f"Fecha creación del pedido: {escape(fecha_pedido)}", estilos["fecha_pedido"])],
        ],
        colWidths=[42 * mm],
        rowHeights=[13 * mm, 9 * mm, 10 * mm],
    )
    bloque_derecho.setStyle(_estilo_tabla(
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        padding=1,
    ))

    # Las columnas de los costados miden lo mismo para que el título quede
    # centrado en el rectángulo y no solo en el espacio entre los logos.
    encabezado = Table(
        [[logo_infolog, _parrafo(f"REMITO {id_pedido}", estilos["titulo"]), bloque_derecho]],
        colWidths=[42 * mm, 106 * mm, 42 * mm],
        rowHeights=[32 * mm],
    )
    encabezado.setStyle(_estilo_tabla(
        ("BACKGROUND", (0, 0), (-1, -1), COLOR_AZUL),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("BOX", (0, 0), (-1, -1), 0.5, COLOR_AZUL),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ))
    return encabezado


def _crear_tabla_contactos(datos, estilos):
    """Tabla de contactos y fechas.

    Por diseño del remito, la columna "Datos del contacto" muestra al
    destinatario y la columna "Datos del destinatario" muestra al
    solicitante. No es un error.
    """
    cabecera = estilos["cabecera_tabla"]

    def fila_datos(*columnas):
        return [_parrafo(datos.get(columna), estilos["dato"]) for columna in columnas]

    filas = [
        [Paragraph("Datos del contacto", estilos["titulo_bloque"]), "",
         Paragraph("Datos del destinatario", estilos["titulo_bloque"]), ""],
        [Paragraph(t, cabecera) for t in ("Nombre", "Email", "Nombre", "Email")],
        fila_datos("Nombre Destinatario", "Email Destinatario", "Nombre Solicitante", "Email Solicitante"),
        [Paragraph(t, cabecera) for t in ("Número de teléfono", "Empresa", "Número de teléfono", "Empresa")],
        fila_datos("Teléfono Destinatario", "Compañía Destinataria", "Teléfono Solicitante", "Compañía Solicitante"),
        [Paragraph("Fecha retiro/entrega", cabecera), "", Paragraph("Fecha devolución", cabecera), ""],
        [_parrafo(datos.get("Fecha Visita Aprob."), estilos["dato"]), "",
         _parrafo(datos.get("Fecha Aprob. Retorno"), estilos["dato"]), ""],
    ]

    tabla = Table(filas, colWidths=[45 * mm, 50 * mm, 45 * mm, 50 * mm], hAlign="CENTER")
    tabla.setStyle(_estilo_tabla(
        *[("SPAN", (0, fila), (1, fila)) for fila in (0, 5, 6)],
        *[("SPAN", (2, fila), (3, fila)) for fila in (0, 5, 6)],
        ("BACKGROUND", (0, 0), (-1, 0), COLOR_AZUL_CLARO),
        *[("BACKGROUND", (0, fila), (-1, fila), COLOR_AZUL_MUY_CLARO) for fila in (1, 3, 5)],
        ("GRID", (0, 0), (-1, -1), 0.35, COLOR_BORDE),
        ("LEFTPADDING", (0, 0), (-1, -1), 2.5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2.5),
        padding=4,
    ))
    return tabla


def _crear_tabla_elementos(tipo, elementos, estilos):
    """Tabla de un tipo de elemento y la altura mínima para empezarla en la página.

    La altura mínima es la del encabezado más la primera fila: se usa para
    no dejar el título de la sección solo al pie de una hoja.
    """
    columnas = CATEGORIAS[tipo].columnas
    escala = ANCHO_BLOQUES / sum(c.ancho for c in columnas)
    anchos = [c.ancho * escala for c in columnas]

    encabezado = [Paragraph(c.titulo, estilos["cabecera_tabla"]) for c in columnas]
    filas = [encabezado] + [[_parrafo(e.get(c.campo), estilos["celda"]) for c in columnas] for e in elementos]

    estilo = _estilo_tabla(
        ("BACKGROUND", (0, 0), (-1, 0), COLOR_VERDE_CLARO),
        ("GRID", (0, 0), (-1, -1), 0.35, COLOR_BORDE),
    )

    tabla = Table(filas, colWidths=anchos, repeatRows=1, hAlign="CENTER")
    tabla.setStyle(estilo)

    inicio = Table(filas[:2], colWidths=anchos)
    inicio.setStyle(estilo)
    _, alto_inicio = inicio.wrap(ANCHO_BLOQUES, A4[1])

    return tabla, alto_inicio


def _crear_seccion(tipo, elementos, estilos):
    categoria = CATEGORIAS[tipo]
    titulo = _franja(
        Paragraph(f"{categoria.titulo_pdf}: {plural(len(elementos), 'elemento')}", estilos["seccion"]),
        COLOR_AZUL_CLARO,
    )
    tabla, alto_inicio = _crear_tabla_elementos(tipo, elementos, estilos)
    _, alto_titulo = titulo.wrap(ANCHO_BLOQUES, A4[1])

    # CondPageBreak solo importa en la primera sección, que comparte hoja con
    # el encabezado: pasa a la siguiente si no entran el título, el encabezado
    # de la tabla y al menos un elemento. La tabla no va en un KeepTogether
    # porque una tabla larga tiene que poder partirse entre páginas.
    return [CondPageBreak(alto_titulo + alto_inicio + 2 * mm), titulo, tabla]


def _crear_bloque_recepcion(estilos):
    negrita = estilos["dato_negrita"]
    campos = ("Nombre:", "Número de documento:", "Empresa:", "Fecha:", "Firma:")
    tabla = Table(
        [[Paragraph(campo, negrita), ""] for campo in campos],
        colWidths=[50 * mm, 140 * mm],
        rowHeights=[9 * mm] * 4 + [14 * mm],
    )
    tabla.setStyle(_estilo_tabla(
        ("BOX", (0, 0), (-1, -1), 0.5, COLOR_BORDE),
        ("INNERGRID", (0, 0), (-1, -1), 0.35, COLOR_BORDE),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ))
    # KeepTogether para que la firma no quede partida entre dos hojas.
    return KeepTogether([Spacer(1, 8 * mm), Paragraph("DATOS DE LA PERSONA QUE LO RECIBE", negrita),
                         Spacer(1, 2 * mm), tabla])


# ============================================================
# Generador principal
# ============================================================

def generar_remito(datos_pedido, grupos, ruta_salida):
    """Genera el remito y devuelve la ruta absoluta del PDF.

    Parámetros:
        datos_pedido: diccionario devuelto por leer_excel().
        grupos: {tipo: elementos devueltos por leer_datos_elementos()}.
        ruta_salida: ruta completa del PDF a crear (se sobrescribe si existe).

    Lanza ValueError si faltan datos y PermissionError si el archivo no se
    puede escribir (por ejemplo, porque está abierto en otro programa).
    """
    errores = validar_datos_remito(datos_pedido, grupos)
    if errores:
        raise ValueError("No se puede generar el remito:\n\n- " + "\n- ".join(errores))

    ruta_salida = os.path.abspath(ruta_salida)
    os.makedirs(os.path.dirname(ruta_salida), exist_ok=True)

    estilos = _crear_estilos()
    tipos =[tipo for tipo in CATEGORIAS if grupos.get(tipo)]
    id_pedido = a_texto(datos_pedido.get("ID Prestamo"))

    contenido = [
        _crear_encabezado(datos_pedido, estilos),
        _franja(
            Paragraph(
                "Elementos incluidos en el pedido: " + " - ".join(CATEGORIAS[t].titulo_pdf for t in tipos),
                estilos["dato_negrita"],
            ),
            COLOR_AZUL_MUY_CLARO,
        ),
        _crear_tabla_contactos(datos_pedido, estilos),
        _franja(Paragraph("Notas:", estilos["dato_negrita"]), COLOR_NARANJA_CLARO),
        _franja(_parrafo(datos_pedido.get("Notas Solicitante"), estilos["dato"], vacio=""),
                COLOR_NARANJA_MUY_CLARO, alineacion_vertical="TOP", minRowHeights=[18 * mm]),
        Spacer(1, 2 * mm),
    ]
    # Cada tipo empieza en hoja nueva para que su título no quede perdido en
    # medio de una página.
    for indice, tipo in enumerate(tipos):
        if indice:
            contenido.append(PageBreak())
        contenido += _crear_seccion(tipo, grupos[tipo], estilos)
    contenido.append(_crear_bloque_recepcion(estilos))

    documento = BaseDocTemplate(
        ruta_salida, pagesize=A4,
        leftMargin=MARGEN, rightMargin=MARGEN, topMargin=MARGEN, bottomMargin=MARGEN,
        title=f"Remito {id_pedido}", author=APP_NOMBRE,
    )
    marco = Frame(MARGEN, MARGEN, ANCHO_BLOQUES, A4[1] - 2 * MARGEN, id="contenido")
    documento.addPageTemplates([PageTemplate(id="A4", frames=[marco])])

    try:
        documento.build(contenido, canvasmaker=NumeradoCanvas)
    except PermissionError as error:
        raise PermissionError(
            f"No se pudo guardar el remito en:\n{ruta_salida}\n\n"
            "Si el PDF está abierto en otro programa, ciérrelo e intente nuevamente."
        ) from error

    return ruta_salida
