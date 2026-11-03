"""Widgets reutilizables de la interfaz.

Los datos del pedido se muestran con widgets comunes de Tkinter y no con los
de CustomTkinter: cada widget de CustomTkinter es un lienzo que se redibuja
por separado, y con las decenas de campos de un pedido la pantalla tardaba
en aparecer y se veía cómo se iba armando de a poco.
"""

import tkinter as tk


class _MenuCopiar:
    """Agrega un menú de clic derecho con la opción "Copiar"."""

    def _crear_menu_copiar(self):
        self.menu_contextual = tk.Menu(self, tearoff=0)
        self.menu_contextual.add_command(label="Copiar", command=self.copiar)
        self.bind("<Button-3>", self._mostrar_menu)

    def copiar(self):
        self.clipboard_clear()
        self.clipboard_append(self.texto())

    def _mostrar_menu(self, evento):
        try:
            self.focus_set()
            self.menu_contextual.tk_popup(evento.x_root, evento.y_root)
        finally:
            self.menu_contextual.grab_release()


class DatoCopiable(_MenuCopiar, tk.Entry):
    """Texto de una línea, de solo lectura, que se puede seleccionar y copiar."""

    def __init__(self, master, texto, fuente, fondo):
        super().__init__(master, font=fuente, relief="flat", bd=0, highlightthickness=0,
                         readonlybackground=fondo, width=1)
        self.insert(0, texto)
        self.configure(state="readonly")
        self._crear_menu_copiar()

    def texto(self):
        return self.get()


class DatoCopiableMultilinea(_MenuCopiar, tk.Text):
    """Igual que DatoCopiable, pero para textos largos o con saltos de línea."""

    def __init__(self, master, texto, fuente, fondo, lineas):
        super().__init__(master, font=fuente, relief="flat", bd=0, highlightthickness=0, bg=fondo,
                         wrap="word", width=1, height=lineas)
        self.insert("1.0", texto)
        self.configure(state="disabled")
        self._crear_menu_copiar()

    def texto(self):
        return self.get("1.0", "end-1c")


def crear_dato(master, texto, fuente, fondo, max_caracteres_por_linea=45):
    """Devuelve un DatoCopiable o, si el texto no entra en una línea, uno multilínea.

    El ancho lo da quien lo ubica (con fill="x"). La altura del multilínea se
    estima por cantidad de caracteres y se limita a 4 líneas; si el texto es
    más largo, el cuadro tiene su propio scroll con la rueda del mouse.
    """
    texto = texto or ""
    if len(texto) <= max_caracteres_por_linea and "\n" not in texto:
        return DatoCopiable(master, texto, fuente, fondo)

    lineas = sum(
        max(1, -(-len(renglon) // max_caracteres_por_linea))  # división redondeando hacia arriba
        for renglon in texto.splitlines() or [""]
    )
    return DatoCopiableMultilinea(master, texto, fuente, fondo, lineas=min(max(lineas, 2), 4))
