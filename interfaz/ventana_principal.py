"""Ventana principal de INFOLOG REMITOS.

Flujo de uso:
    1. "Cargar pedido": se lee el Excel del pedido y se muestra la pantalla
       con los datos del pedido. Con eso se habilitan "Detalles del pedido"
       (vuelve a esa pantalla cuando se quiera) y "Cargar elementos".
    2. "Cargar elementos": muestra la pantalla "Elementos del pedido". Cada
       Excel cargado con "Cargar Excel de elementos" crea el recuadro de su
       tipo; desde el recuadro se puede reemplazar por otro archivo.
    3. Con al menos un tipo cargado se habilita "Generar Remito", que crea el
       PDF en la carpeta elegida por el usuario.

Las lecturas de Excel y la generación del PDF corren en un hilo aparte para
no congelar la ventana (ver ejecutar_en_segundo_plano).
"""

import logging
import os
import queue
import threading
from tkinter import filedialog, messagebox
import tkinter as tk

import customtkinter as ctk
from PIL import Image

from config import (
    ALTO_VENTANA, ANCHO_VENTANA, APP_AUTOR, APP_CONTACTO, APP_NOMBRE, APP_VERSION,
    COLOR_BOTON, COLOR_BOTON_CARPETA, COLOR_BOTON_CARPETA_HOVER, COLOR_BOTON_PEDIDO, COLOR_BOTON_PEDIDO_HOVER,
    COLOR_BOTON_QUITAR, COLOR_BOTON_QUITAR_HOVER,
    COLOR_FONDO, COLOR_MENU, COLOR_PASO,
    COLOR_PASO_LISTO, COLOR_PASO_SIGUIENTE, COLOR_RESALTADO, COLOR_SEPARADOR, COLOR_TEXTO_TITULO,
    TIPOS_ELEMENTOS,
)
from excel.lector_elementos import detectar_tipo, leer_datos_elementos
from excel.lector_excel import leer_excel
from excel.validador_excel import validar_pedido
from interfaz.componentes import crear_dato
from remitos.generador_remito import generar_remito, nombre_archivo_remito, validar_datos_remito
from utilidades import preferencias
from utilidades.rutas import carpeta_escritorio, ruta_recurso
from utilidades.texto import a_texto, plural

log = logging.getLogger(__name__)

FUENTE = "Segoe UI"
TIPOS_ARCHIVO_EXCEL = [("Archivos Excel", "*.xlsx")]

COLOR_DESHABILITADO = "#9E9E9E"
COLOR_TEXTO_DESHABILITADO = "#E8E8E8"
COLOR_PANEL_ELEMENTOS = "#E8EDF3"
COLOR_RECUADRO = "#FFFFFF"
COLOR_OK = "#2E7D32"

# Cada cuánto se revisa si terminó una tarea en segundo plano (ms).
INTERVALO_REVISION = 50

# Recuadros por fila en la pantalla de elementos. Con 3 entran justo en el
# ancho de la ventana sin scroll horizontal.
TIPOS_POR_FILA = 3


class App(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.datos_pedido = None   # fila del pedido tal como la devuelve leer_excel()
        self.elementos = {}        # tipo -> filas de su Excel: lo que va en el remito
        self.remito_generado = False  # se generó el PDF con los datos que hay cargados ahora
        self.carpeta_remitos = preferencias.cargar_carpeta_remitos()

        # Cada tarea en segundo plano recibe un número. Si termina una que ya
        # no es la última (por ejemplo, se cargó otro pedido), se descarta.
        self.tarea_actual = 0

        # Las pantallas del pedido se arman una sola vez por pedido y después
        # solo se alternan; armarlas de nuevo en cada clic era lo que demoraba.
        self.pantallas = {}           # "pedido" / "elementos" -> marco de la pantalla
        self.frame_cargando = None
        self.animacion_id = None

        self._configurar_ventana()
        self.escala = ctk.ScalingTracker.get_widget_scaling(self)
        self._cargar_imagenes()
        self._crear_menu()
        self._crear_area_principal()
        self._crear_pantalla_inicio()
        self._crear_barra_estado()
        self._actualizar_botones()

        # Se verifica después de mostrar la ventana para que el aviso no
        # aparezca antes que la aplicación.
        self.after(150, self._verificar_carpeta_remitos)

    def report_callback_exception(self, tipo, valor, traza):
        """Cualquier error no controlado en un evento de la interfaz termina acá."""
        log.error("Error no controlado en la interfaz", exc_info=(tipo, valor, traza))
        messagebox.showerror("Error inesperado", f"{valor}\n\nEl detalle quedó registrado en el archivo de log.")

    # ============================================================
    # Construcción de la ventana
    # ============================================================

    def _configurar_ventana(self):
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title(APP_NOMBRE)
        self.iconbitmap(ruta_recurso("imagenes", "icono_infolog.ico"))
        self.resizable(False, False)
        self.configure(fg_color=COLOR_FONDO)

        x = (self.winfo_screenwidth() - ANCHO_VENTANA) // 2
        y = (self.winfo_screenheight() - ALTO_VENTANA) // 2
        self.geometry(f"{ANCHO_VENTANA}x{ALTO_VENTANA}+{x}+{y}")

    def _cargar_imagenes(self):
        def imagen(nombre, tamano):
            return ctk.CTkImage(Image.open(ruta_recurso("imagenes", nombre)), size=tamano)

        self.iconos_estado = {
            "ok": imagen("estado_ok.png", (16, 16)),
            "espera": imagen("estado_espera.png", (16, 16)),
            "advertencia": imagen("estado_advertencia.png", (16, 16)),
            "error": imagen("estado_error.png", (16, 16)),
        }
        self.logo = imagen("logo_infolog.png", (100, 90))

    def _crear_menu(self):
        menu = ctk.CTkFrame(self, width=240, corner_radius=0, fg_color=COLOR_MENU)
        menu.pack(side="left", fill="y")
        # Sin esto el menú toma el ancho de los botones y no el suyo, y
        # CustomTkinter deja una franja clara de 1 px en el borde derecho.
        menu.pack_propagate(False)

        # Tres bloques: arriba el logo, "Cargar pedido" y "Detalles del pedido";
        # abajo la carpeta y "Acerca de"; en el medio los pasos 2 y 3, que se
        # habilitan a medida que se cargan los datos. El del medio va con expand=True
        # para que quede centrado en el espacio libre.
        superior = ctk.CTkFrame(menu, corner_radius=0, fg_color=COLOR_MENU)
        superior.pack(fill="x")
        inferior = ctk.CTkFrame(menu, corner_radius=0, fg_color=COLOR_MENU)
        inferior.pack(side="bottom", fill="x", pady=10)
        medio = ctk.CTkFrame(menu, corner_radius=0, fg_color=COLOR_MENU)
        medio.pack(fill="x", expand=True)

        ctk.CTkLabel(superior, image=self.logo, text="").pack(pady=(20, 5))
        self._separador(superior, pady=(10, 20))

        def boton(padre, texto, comando, habilitado=True, fg_color=COLOR_BOTON, hover_color=None):
            b = ctk.CTkButton(padre, text=texto, width=180, height=40, fg_color=fg_color, hover_color=hover_color,
                              command=comando, border_color=COLOR_RESALTADO)
            b.pack(pady=5, padx=20)
            if not habilitado:
                self._habilitar_boton(b, False)
            return b

        def paso(padre, numero, pady=(0, 0)):
            etiqueta = tk.Label(padre, text=f"PASO {numero}", font=self._fuente(11, "bold"), fg=COLOR_PASO,
                                bg=COLOR_MENU)
            etiqueta.pack(anchor="w", padx=32, pady=pady)
            return etiqueta

        # Hover de los botones azules, para restaurarlo al volver a habilitarlos.
        self.color_hover = ctk.ThemeManager.theme["CTkButton"]["hover_color"]

        etiqueta_1 = paso(superior, 1)
        self.btn_cargar_pedido = boton(superior, "Cargar pedido", self.cargar_pedido,
                                       fg_color=COLOR_BOTON_PEDIDO, hover_color=COLOR_BOTON_PEDIDO_HOVER)

        # "Detalles del pedido" va pegado a "Cargar pedido" porque muestra lo
        # que se cargó en ese paso.
        self.btn_detalles = boton(superior, "Detalles del pedido", self.mostrar_detalles_pedido, habilitado=False)

        etiqueta_2 = paso(medio, 2)
        self.btn_elementos = boton(medio, "Cargar elementos", self.mostrar_elementos, habilitado=False)
        etiqueta_3 = paso(medio, 3, pady=(14, 0))
        self.btn_pdf = boton(medio, "Generar Remito", self.generar_remito_pdf, habilitado=False)

        # Etiqueta de cada paso del flujo; _actualizar_botones() les pone
        # LISTO / SIGUIENTE según lo que ya se cargó.
        self.etiquetas_paso = {1: etiqueta_1, 2: etiqueta_2, 3: etiqueta_3}
        # Botones que se marcan en amarillo cuando el usuario está ahí (ver _marcar_boton).
        self.botones_marcables = (self.btn_cargar_pedido, self.btn_detalles, self.btn_elementos, self.btn_pdf)
        self.boton_marcado = self.btn_cargar_pedido

        boton(inferior, "Carpeta de remitos", self.seleccionar_carpeta_remitos,
              fg_color=COLOR_BOTON_CARPETA, hover_color=COLOR_BOTON_CARPETA_HOVER)
        self._separador(inferior, pady=(10, 15))
        ctk.CTkButton(
            inferior, text="Acerca de", width=190, height=50, fg_color=COLOR_BOTON, command=self.acerca_de,
        ).pack(padx=20, pady=(5, 20))

    def _separador(self, padre, pady):
        # Va con un Frame común de Tkinter: CustomTkinter no dibuja los
        # CTkFrame de 2 px de alto y la línea quedaba invisible.
        tk.Frame(padre, height=2, bg=COLOR_SEPARADOR, bd=0, highlightthickness=0).pack(fill="x", padx=15, pady=pady)

    def _habilitar_boton(self, boton, habilitado):
        if habilitado:
            boton.configure(state="normal", fg_color=COLOR_BOTON, hover_color=self.color_hover, text_color="#FFFFFF")
        else:
            boton.configure(state="disabled", fg_color=COLOR_DESHABILITADO, hover_color=COLOR_DESHABILITADO,
                            text_color=COLOR_TEXTO_DESHABILITADO)

    def _crear_area_principal(self):
        self.contenido = ctk.CTkFrame(self, fg_color=COLOR_FONDO, corner_radius=0)
        self.contenido.pack(side="right", fill="both", expand=True)

    def _crear_pantalla_inicio(self):
        self.frame_inicio = ctk.CTkFrame(self.contenido, fg_color="transparent")
        self.frame_inicio.pack(fill="both", expand=True)

        def texto(contenido, pady, tamano=16, peso="normal", color=None):
            ctk.CTkLabel(self.frame_inicio, text=contenido, justify="center",
                         font=(FUENTE, tamano, peso), text_color=color).pack(pady=pady)

        texto(APP_NOMBRE, (50, 5), 28, "bold", COLOR_TEXTO_TITULO)
        texto(f"Versión {APP_VERSION}", 0)
        texto("Sistema para la confección de remitos\nCreada para INFOLOG SRL", (20, 10))
        texto(f"Desarrollado por {APP_AUTOR}  -  {APP_CONTACTO}", 0, tamano=13)

        # Guía de los tres pasos, con los mismos colores que sus botones del menú.
        guia = ctk.CTkFrame(self.frame_inicio, fg_color=COLOR_RECUADRO, corner_radius=12)
        guia.pack(pady=(35, 0), ipadx=10, ipady=10)
        ctk.CTkLabel(guia, text="Cómo generar un remito", font=(FUENTE, 17, "bold"),
                     text_color=COLOR_TEXTO_TITULO).pack(pady=(4, 10))
        pasos = (
            ("Cargar pedido", "Seleccione el Excel del pedido exportado de la base de datos.", COLOR_BOTON_PEDIDO),
            ("Cargar elementos", "Cargue el/los Excel con el detalle de cada tipo de elemento.", COLOR_BOTON),
            ("Generar Remito", "Se habilita cuando están todos los datos y crea el PDF.", COLOR_BOTON),
        )
        for numero, (nombre, descripcion, color) in enumerate(pasos, start=1):
            fila = ctk.CTkFrame(guia, fg_color="transparent")
            fila.pack(fill="x", padx=20, pady=5)
            ctk.CTkLabel(fila, text=str(numero), width=32, height=32, corner_radius=16, fg_color=color,
                         text_color="#FFFFFF", font=(FUENTE, 15, "bold")).pack(side="left")
            ctk.CTkLabel(fila, text=f"{nombre}:", font=(FUENTE, 15, "bold"),
                         text_color=COLOR_TEXTO_TITULO).pack(side="left", padx=(12, 6))
            ctk.CTkLabel(fila, text=descripcion, font=(FUENTE, 15)).pack(side="left")

    def _crear_barra_estado(self):
        self.barra_estado = ctk.CTkLabel(
            self.contenido, text="Estado: Esperando pedido...", image=self.iconos_estado["espera"],
            compound="left", anchor="w", height=30, font=(FUENTE, 15),
        )
        self.barra_estado.pack(side="bottom", fill="x", padx=15, pady=10)

    def actualizar_estado(self, tipo, mensaje):
        """Muestra un mensaje en la barra inferior. tipo: ok, espera, advertencia o error."""
        self.barra_estado.configure(image=self.iconos_estado.get(tipo, self.iconos_estado["espera"]),
                                    text=f"Estado: {mensaje}")

    def _actualizar_botones(self):
        """Habilita los botones del menú según lo que haya cargado.

        "Generar Remito" depende solo de validar_datos_remito(), así la regla
        está en un único lugar.
        """
        hay_pedido = self.datos_pedido is not None
        remito_listo = hay_pedido and not validar_datos_remito(self.datos_pedido, self.elementos)
        self._habilitar_boton(self.btn_detalles, hay_pedido)
        self._habilitar_boton(self.btn_elementos, hay_pedido)
        self._habilitar_boton(self.btn_pdf, remito_listo)

        listos = {1: hay_pedido, 2: remito_listo, 3: self.remito_generado}
        siguiente = next((numero for numero, listo in listos.items() if not listo), None)
        for numero, etiqueta in self.etiquetas_paso.items():
            if listos[numero]:
                texto, color = f"PASO {numero}  -  LISTO", COLOR_PASO_LISTO
            elif numero == siguiente:
                texto, color = f"PASO {numero}  -  SIGUIENTE", COLOR_PASO_SIGUIENTE
            else:
                texto, color = f"PASO {numero}", COLOR_PASO
            etiqueta.configure(text=texto, fg=color)
        self._marcar_boton(self.boton_marcado)

    def _marcar_boton(self, boton):
        """Marca en amarillo el botón de la pantalla o acción donde está el usuario."""
        self.boton_marcado = boton
        for b in self.botones_marcables:
            b.configure(border_width=2 if b is boton else 0)

    # ============================================================
    # Tareas en segundo plano
    # ============================================================

    def ejecutar_en_segundo_plano(self, mensaje, tarea, al_terminar, titulo_error="Error"):
        """Ejecuta tarea() en otro hilo y después al_terminar(resultado) en el principal.

        Tkinter no se puede usar desde otro hilo, así que el hilo solo deja el
        resultado en una cola y la ventana la revisa cada pocos milisegundos.
        Mientras tanto se muestra la pantalla de carga, que tapa toda la
        ventana y evita que el usuario dispare otra acción.
        """
        self.tarea_actual += 1
        numero = self.tarea_actual
        cola = queue.Queue()

        def trabajar():
            try:
                cola.put((True, tarea()))
            except Exception as error:  # se informa en el hilo principal
                cola.put((False, error))

        self._mostrar_cargando(mensaje)
        threading.Thread(target=trabajar, daemon=True).start()
        self.after(INTERVALO_REVISION, self._esperar_resultado, cola, numero, al_terminar, titulo_error)

    def _esperar_resultado(self, cola, numero, al_terminar, titulo_error):
        try:
            exito, resultado = cola.get_nowait()
        except queue.Empty:
            self.after(INTERVALO_REVISION, self._esperar_resultado, cola, numero, al_terminar, titulo_error)
            return

        if numero != self.tarea_actual:
            return

        # La pantalla de carga se saca recién después de al_terminar: así lo
        # que se arma queda tapado y aparece completo de una vez, en lugar de
        # verse cómo se dibuja. Los avisos la sacan antes (ver _avisar).
        try:
            if not exito:
                raise resultado
            al_terminar(resultado)
            self.update_idletasks()
        except Exception as error:
            self._informar_error(titulo_error, error)
        finally:
            self._actualizar_botones()
            self._ocultar_cargando()

    def _avisar(self, titulo, mensaje):
        """Muestra una advertencia, sacando antes la pantalla de carga si está."""
        self._ocultar_cargando()
        messagebox.showwarning(titulo, mensaje)

    def _informar_error(self, titulo, error):
        self._ocultar_cargando()
        self.actualizar_estado("error", titulo)
        if isinstance(error, (ValueError, PermissionError, FileNotFoundError)):
            # Errores esperables (archivo equivocado, PDF abierto...): el
            # mensaje ya está pensado para el usuario.
            log.warning("%s: %s", titulo, error)
            messagebox.showwarning(titulo, str(error))
        else:
            log.exception(titulo, exc_info=error)
            messagebox.showerror(titulo, f"{error}\n\nEl detalle quedó registrado en el archivo de log.")

    def _mostrar_cargando(self, mensaje):
        self._ocultar_cargando()

        self.frame_cargando = ctk.CTkFrame(self, fg_color=COLOR_FONDO)
        self.frame_cargando.place(relx=0, rely=0, relwidth=1, relheight=1)

        ctk.CTkLabel(self.frame_cargando, text=APP_NOMBRE, font=(FUENTE, 24, "bold"),
                     text_color=COLOR_TEXTO_TITULO).pack(pady=(180, 15))
        ctk.CTkLabel(self.frame_cargando, text=mensaje, font=(FUENTE, 16),
                     text_color=COLOR_TEXTO_TITULO).pack()
        self.lbl_puntos = ctk.CTkLabel(self.frame_cargando, text="", font=(FUENTE, 18, "bold"),
                                       text_color=COLOR_TEXTO_TITULO)
        self.lbl_puntos.pack(pady=(10, 0))

        self.frame_cargando.lift()
        self._animar_puntos(0)

    def _animar_puntos(self, paso):
        self.lbl_puntos.configure(text="." * (paso + 1))
        self.animacion_id = self.after(400, self._animar_puntos, (paso + 1) % 3)

    def _ocultar_cargando(self):
        if self.animacion_id is not None:
            self.after_cancel(self.animacion_id)
            self.animacion_id = None
        if self.frame_cargando is not None:
            self.frame_cargando.destroy()
            self.frame_cargando = None

    # ============================================================
    # Pantallas
    # ============================================================

    def _fuente(self, tamano, peso="normal"):
        """Fuente para widgets comunes de Tkinter, del mismo tamaño que en CustomTkinter.

        CustomTkinter toma el tamaño en píxeles y lo escala según el zoom de
        Windows; Tkinter lo toma en puntos. El tamaño negativo es en píxeles.
        """
        return (FUENTE, -round(tamano * self.escala), peso)

    def _etiqueta(self, padre, texto, tamano, fondo=COLOR_FONDO, color=COLOR_TEXTO_TITULO, peso="bold"):
        return tk.Label(padre, text=texto, font=self._fuente(tamano, peso), fg=color, bg=fondo)

    def _armar_pantallas(self):
        """Arma las pantallas del pedido recién cargado, sin mostrarlas."""
        for pantalla in self.pantallas.values():
            pantalla.destroy()
        self.pantallas = {"pedido": self._crear_pantalla_pedido(), "elementos": self._crear_pantalla_elementos()}

    def _refrescar_elementos(self):
        """Vuelve a armar la pantalla de elementos con lo que hay cargado y la muestra."""
        self.pantallas["elementos"].destroy()
        self.pantallas["elementos"] = self._crear_pantalla_elementos()
        self.mostrar_elementos()

    def _mostrar_pantalla(self, nombre):
        self.frame_inicio.pack_forget()
        for clave, pantalla in self.pantallas.items():
            if clave != nombre:
                pantalla.pack_forget()
        self.pantallas[nombre].pack(fill="both", expand=True, padx=18, pady=20)
        self._marcar_boton(self.btn_detalles if nombre == "pedido" else self.btn_elementos)

    def mostrar_detalles_pedido(self):
        """Muestra la pantalla con los datos del pedido."""
        if "pedido" in self.pantallas:
            self._mostrar_pantalla("pedido")

    def mostrar_elementos(self):
        """Muestra la pantalla "Elementos del pedido"."""
        if "elementos" in self.pantallas:
            self._mostrar_pantalla("elementos")

    # ============================================================
    # Pantalla: datos del pedido
    # ============================================================

    def _crear_pantalla_pedido(self):
        marco = ctk.CTkScrollableFrame(self.contenido, fg_color="transparent")
        datos = self.datos_pedido

        self._etiqueta(marco, "DETALLES DEL PEDIDO", 24).pack(pady=(0, 20))

        fila = self._fila(marco, pady=(0, 15))
        self._campo(fila, "ID del pedido", datos["ID Prestamo"], tamano=18)
        self._campo(fila, "Fecha Solicitud", datos["Fecha Solicitud"], tamano=18, peso="bold",
                    color_titulo=COLOR_BOTON)

        fila = self._fila(marco, pady=(0, 15))
        self._campo(fila, "Usuario Solicitante", datos["Usuario Solicitante"], tamano=16)

        for dato in ("Nombre", "Compañía", "Dirección", "Email", "Teléfono"):
            # "Compañía Destinataria" es la única columna en femenino.
            destinatario = "Compañía Destinataria" if dato == "Compañía" else f"{dato} Destinatario"
            fila = self._fila(marco, pady=4)
            self._campo(fila, destinatario, datos[destinatario])
            self._campo(fila, f"{dato} Solicitante", datos[f"{dato} Solicitante"])

        self._campo(marco, "Notas:", datos["Notas Solicitante"], max_caracteres=85, en_columna=False,
                    pady_titulo=(12, 0), pady_valor=(0, 10))
        for columna in ("Código Estatus", "Detalles Estatus (ES)", "Tipo Entrega"):
            self._campo(marco, columna, datos[columna], max_caracteres=85, en_columna=False, pady_valor=(0, 5))

        self._etiqueta(marco, "Seleccione 'Cargar elementos' para cargar el Excel de cada tipo.", 14).pack(
            pady=(20, 10))
        return marco

    def _fila(self, padre, pady):
        fila = tk.Frame(padre, bg=COLOR_FONDO)
        fila.pack(fill="x", pady=pady)
        # uniform hace que las dos columnas midan lo mismo aunque el texto
        # de una sea más largo, así quedan alineadas entre filas.
        fila.grid_columnconfigure((0, 1), weight=1, uniform="columnas")
        return fila

    def _campo(self, padre, titulo, valor, tamano=15, peso="normal", max_caracteres=38,
               color_titulo=COLOR_TEXTO_TITULO, en_columna=True, pady_titulo=0, pady_valor=0):
        """Título en negrita con el valor copiable debajo.

        Con en_columna=True ocupa una de las dos columnas de una fila; si no,
        se apila directamente en el padre y usa todo el ancho.
        """
        bloque = padre
        if en_columna:
            bloque = tk.Frame(padre, bg=COLOR_FONDO)
            bloque.grid(row=0, column=len(padre.grid_slaves()), sticky="ew")

        self._etiqueta(bloque, titulo, 14, color=color_titulo).pack(anchor="w", pady=pady_titulo)
        dato = crear_dato(bloque, a_texto(valor, con_hora=True), self._fuente(tamano, peso), COLOR_FONDO,
                          max_caracteres)
        dato.pack(fill="x", padx=(8, 20), pady=pady_valor)

    # ============================================================
    # Pantalla: elementos del pedido
    # ============================================================

    def _crear_pantalla_elementos(self):
        marco = tk.Frame(self.contenido, bg=COLOR_FONDO)
        total = sum(len(elementos) for elementos in self.elementos.values())
        self._etiqueta(marco, "ELEMENTOS DEL PEDIDO", 24).pack(pady=(0, 5))
        self._etiqueta(marco, f"Pedido {a_texto(self.datos_pedido['ID Prestamo'])}  -  "
                              f"Total de elementos: {total}", 16).pack(pady=(0, 15))

        ctk.CTkButton(marco, text="Cargar Excel de elementos", width=240, height=40, fg_color=COLOR_BOTON,
                      command=self.cargar_excel_elementos).pack(pady=(0, 12))

        grilla = ctk.CTkScrollableFrame(marco, fg_color=COLOR_PANEL_ELEMENTOS, corner_radius=12)
        grilla.pack(fill="both", expand=True)
        for columna in range(TIPOS_POR_FILA):
            grilla.grid_columnconfigure(columna, weight=1, uniform="tipos")

        tipos = [tipo for tipo in TIPOS_ELEMENTOS if tipo in self.elementos]
        if not tipos:
            self._etiqueta(grilla, "Todavía no se cargaron elementos.\nSeleccione el Excel de cada tipo con "
                                   "'Cargar Excel de elementos'.", 15, fondo=COLOR_PANEL_ELEMENTOS, peso="normal",
                           ).grid(row=0, column=0, columnspan=TIPOS_POR_FILA, pady=60)

        for indice, tipo in enumerate(tipos):
            fila, columna = divmod(indice, TIPOS_POR_FILA)
            self._crear_recuadro_tipo(grilla, tipo).grid(row=fila, column=columna, padx=8, pady=8, sticky="nsew")
        return marco

    def _crear_recuadro_tipo(self, padre, tipo):
        recuadro = ctk.CTkFrame(padre, fg_color=COLOR_RECUADRO, corner_radius=10)

        # wraplength porque los nombres más largos no entran en el ancho del recuadro.
        titulo = self._etiqueta(recuadro, f"{TIPOS_ELEMENTOS[tipo]} ({len(self.elementos[tipo])})", 14,
                                fondo=COLOR_RECUADRO)
        titulo.configure(wraplength=round(185 * self.escala), justify="left")
        titulo.pack(anchor="w", padx=12, pady=(10, 2))

        self._etiqueta(recuadro, "Datos cargados", 12, fondo=COLOR_RECUADRO, color=COLOR_OK).pack(
            anchor="w", padx=12, pady=(0, 8))

        ctk.CTkButton(
            recuadro, text="Eliminar", font=(FUENTE, 11, "bold"), height=34, fg_color=COLOR_BOTON_QUITAR,
            hover_color=COLOR_BOTON_QUITAR_HOVER, command=lambda: self.quitar_tipo(tipo),
        ).pack(fill="x", padx=12, pady=(0, 12))
        return recuadro

    # ============================================================
    # Carga del pedido
    # ============================================================

    def cargar_pedido(self):
        ruta = filedialog.askopenfilename(title="Seleccionar pedido", filetypes=TIPOS_ARCHIVO_EXCEL)
        if not ruta:
            return

        def leer():
            datos = leer_excel(ruta)
            errores = validar_pedido(datos)
            if errores:
                raise ValueError("\n".join(errores))
            return datos

        def mostrar(datos):
            # El pedido anterior se reemplaza recién ahora: si el archivo
            # nuevo tenía un error, el usuario no pierde lo que tenía cargado.
            self.datos_pedido = datos
            self.elementos = {}
            self.remito_generado = False
            self._armar_pantallas()
            self.mostrar_detalles_pedido()

            nombre_archivo = os.path.basename(ruta)
            log.info("Pedido cargado: %s", nombre_archivo)
            self.actualizar_estado("ok", f"Pedido cargado correctamente: {nombre_archivo}")

        self.ejecutar_en_segundo_plano("Cargando pedido...", leer, mostrar, "Error al cargar el pedido")

    # ============================================================
    # Elementos del pedido
    # ============================================================

    def quitar_tipo(self, tipo):
        """Saca del remito un tipo cargado, por ejemplo si se eligió un Excel equivocado."""
        nombre = TIPOS_ELEMENTOS[tipo]
        cantidad = plural(len(self.elementos[tipo]), "elemento")
        if not messagebox.askyesno("Eliminar elementos", f"¿Desea eliminar {nombre} ({cantidad}) del remito?"):
            return

        del self.elementos[tipo]
        self.remito_generado = False
        self._refrescar_elementos()
        self._actualizar_botones()
        log.info("Se eliminó %s (%s)", nombre, cantidad)
        self.actualizar_estado("ok", f"Se eliminó {nombre} del remito ({cantidad})")

    def cargar_excel_elementos(self):
        """Carga un Excel: el tipo sale de sus códigos de barras y crea su recuadro."""
        ruta = filedialog.askopenfilename(title="Seleccionar Excel de elementos", filetypes=TIPOS_ARCHIVO_EXCEL)
        if not ruta:
            return

        self.ejecutar_en_segundo_plano(
            "Cargando datos de los elementos...",
            lambda: leer_datos_elementos(ruta, detectar_tipo(ruta)),
            self._asociar_datos_elementos,
            "Error al cargar los datos de los elementos",
        )

    def _asociar_datos_elementos(self, resultado):
        """Guarda los elementos del Excel; si el tipo ya estaba, reemplaza los anteriores."""
        tipo = resultado["tipo"]
        elementos = resultado["elementos"]
        self.elementos[tipo] = elementos
        self._refrescar_elementos()
        # Con datos nuevos, el remito que se haya generado antes ya no está al día.
        self.remito_generado = False

        nombre = TIPOS_ELEMENTOS[tipo]
        nombre_archivo = os.path.basename(resultado["ruta"])
        cantidad = plural(len(elementos), "elemento")
        log.info("Datos de %s cargados desde %s: %s", nombre, nombre_archivo, cantidad)

        if resultado["duplicados"]:
            self.actualizar_estado("advertencia", f"Datos de {nombre} con códigos repetidos ({nombre_archivo})")
            self._avisar("Códigos repetidos", f"El archivo tiene {len(resultado['duplicados'])} código(s) de barras "
                                              "repetidos; se usó la primera fila de cada uno.")
        else:
            self.actualizar_estado("ok", f"Datos de {nombre} cargados: {nombre_archivo} ({cantidad})")

    # ============================================================
    # Remito
    # ============================================================

    def generar_remito_pdf(self):
        if self.datos_pedido is None:
            return
        self._marcar_boton(self.btn_pdf)

        nombre_pdf = nombre_archivo_remito(self.datos_pedido["ID Prestamo"])
        ruta_pdf = os.path.join(self.carpeta_remitos, nombre_pdf)
        if os.path.exists(ruta_pdf) and not messagebox.askyesno(
            "Remito existente", f"Ya existe el archivo:\n\n{nombre_pdf}\n\n¿Desea reemplazarlo?"
        ):
            return

        self.ejecutar_en_segundo_plano(
            "Generando remito...",
            lambda: generar_remito(self.datos_pedido, self.elementos, ruta_pdf),
            self._remito_generado,
            "Error al generar el remito",
        )

    def _remito_generado(self, ruta_pdf):
        nombre_pdf = os.path.basename(ruta_pdf)
        log.info("Remito generado: %s", ruta_pdf)
        self.remito_generado = True
        self.actualizar_estado("ok", f"Remito generado correctamente: {nombre_pdf}")
        self._ocultar_cargando()

        if messagebox.askyesno(
            "Remito generado",
            f"El remito se generó correctamente.\n\nArchivo:\n{nombre_pdf}\n\n"
            f"Carpeta:\n{os.path.dirname(ruta_pdf)}\n\n¿Desea abrir el PDF ahora?",
        ):
            try:
                os.startfile(ruta_pdf)
            except OSError as error:
                messagebox.showwarning(
                    "PDF generado",
                    f"El PDF se generó, pero no se pudo abrir automáticamente.\n\nRuta:\n{ruta_pdf}\n\nDetalle: {error}",
                )

    # ============================================================
    # Carpeta de remitos
    # ============================================================

    def _verificar_carpeta_remitos(self):
        """Al iniciar, avisa si la carpeta guardada ya no existe y pide otra."""
        if os.path.isdir(self.carpeta_remitos):
            return

        messagebox.showwarning(
            "Carpeta de remitos no encontrada",
            "La carpeta de remitos fue movida o eliminada.\n\nPor favor, seleccione una nueva ubicación.",
        )
        self.seleccionar_carpeta_remitos(mostrar_confirmacion=False)

        if not os.path.isdir(self.carpeta_remitos):
            self.carpeta_remitos = carpeta_escritorio()
            preferencias.guardar_carpeta_remitos(self.carpeta_remitos)
            self.actualizar_estado("advertencia", "No se seleccionó una carpeta. Se utilizará el Escritorio.")

    def seleccionar_carpeta_remitos(self, mostrar_confirmacion=True):
        inicial = self.carpeta_remitos if os.path.isdir(self.carpeta_remitos) else carpeta_escritorio()
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta para guardar los remitos",
                                          initialdir=inicial, mustexist=True)
        if not carpeta:
            return

        self.carpeta_remitos = os.path.abspath(carpeta)
        preferencias.guardar_carpeta_remitos(self.carpeta_remitos)
        self.actualizar_estado("ok", f"Carpeta de remitos seleccionada: {self.carpeta_remitos}")

        if mostrar_confirmacion:
            messagebox.showinfo("Carpeta seleccionada", f"Los próximos remitos se guardarán en:\n\n{self.carpeta_remitos}")

    # ============================================================
    # Acerca de
    # ============================================================

    def acerca_de(self):
        messagebox.showinfo(
            f"Acerca de {APP_NOMBRE}",
            f"{APP_NOMBRE}\nVersión {APP_VERSION}\n\nSistema para la confección de remitos.\n\n"
            f"Desarrollado por:\n{APP_AUTOR}\n\nContacto:\n{APP_CONTACTO}",
        )
