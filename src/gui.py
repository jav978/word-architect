"""Interfaz gráfica de Word-Architect.

Permite redactar contenido, aplicar formato mediante marcas y generar
archivos DOCX profesionales con opciones de página, tipografía, encabezado,
pie de página, numeración, portada y logotipo.
"""

from __future__ import annotations

import re
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import List, Optional

if __package__:
    from .documento import (
        OUTPUT_DIR,
        PROJECT_ROOT,
        ConfiguracionDocumento,
        crear_documento as _crear_documento,
        crear_documento_profesional,
    )
else:
    from documento import (  # type: ignore[no-redef]
        OUTPUT_DIR,
        PROJECT_ROOT,
        ConfiguracionDocumento,
        crear_documento as _crear_documento,
        crear_documento_profesional,
    )

__all__ = [
    "crear_documento",
    "main",
    "PROJECT_ROOT",
    "OUTPUT_DIR",
    "abrir_configuracion",
    "generar_documento",
]

# La configuración seleccionada se mantiene durante la sesión.
configuracion_actual = ConfiguracionDocumento()

# Widgets principales: solo se crean cuando se inicia main().
ventana = None
entrada_titulo = None
cuadro_contenido = None
etiqueta_estado = None

# Plantilla que inserta el botón «▦ Tabla»: cabecera y dos filas de datos.
PLANTILLA_TABLA = "| Concepto | Detalle |\n|  |  |\n|  |  |"


class ToolTip:
    """Mostrar una descripción cuando el ratón pasa sobre un widget."""

    def __init__(self, widget, texto: str, retraso: int = 450):
        self.widget = widget
        self.texto = texto
        self.retraso = retraso
        self._temporizador = None
        self._ventana = None
        widget.bind("<Enter>", self._programar, add="+")
        widget.bind("<Leave>", self._ocultar, add="+")
        widget.bind("<ButtonPress>", self._ocultar, add="+")

    def _programar(self, event=None):
        self._cancelar_temporizador()
        self._temporizador = self.widget.after(self.retraso, self._mostrar)

    def _cancelar_temporizador(self):
        if self._temporizador is not None:
            try:
                self.widget.after_cancel(self._temporizador)
            except tk.TclError:
                pass
            self._temporizador = None

    def _mostrar(self):
        self._temporizador = None
        if self._ventana is not None or not self.texto:
            return
        try:
            if not self.widget.winfo_exists():
                return
            x = self.widget.winfo_rootx() + 8
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
            self._ventana = tk.Toplevel(self.widget)
            self._ventana.wm_overrideredirect(True)
            try:
                self._ventana.attributes("-topmost", True)
            except tk.TclError:
                pass
            self._ventana.wm_geometry(f"+{x}+{y}")
            ttk.Label(
                self._ventana,
                text=self.texto,
                padding=(9, 6),
                relief="solid",
                wraplength=320,
            ).pack()
        except tk.TclError:
            self._ventana = None

    def _ocultar(self, event=None):
        self._cancelar_temporizador()
        if self._ventana is not None:
            try:
                self._ventana.destroy()
            except tk.TclError:
                pass
            self._ventana = None


def crear_documento(
    titulo: str,
    parrafos: List[str],
    ruta_destino: Optional[Path] = None,
) -> Path:
    """API anterior: conserva la generación básica con marcas de formato."""
    return _crear_documento(
        titulo,
        parrafos,
        ruta_destino,
        procesar_marcas=True,
    )


def aplicar_marcador(marcador: str) -> None:
    """Aplicar una marca de formato a las líneas seleccionadas o actuales."""
    try:
        inicio = cuadro_contenido.index("sel.first linestart")
        fin = cuadro_contenido.index("sel.last lineend")
    except tk.TclError:
        inicio = cuadro_contenido.index("insert linestart")
        fin = cuadro_contenido.index("insert lineend")

    contenido = cuadro_contenido.get(inicio, fin)
    resultado = []
    for linea in contenido.split("\n"):
        texto = re.sub(r"^(#{1,3}\s+|[-*]\s+|\d+[.)]\s+)", "", linea.strip())
        resultado.append(marcador + texto if texto else marcador)

    cuadro_contenido.delete(inicio, fin)
    cuadro_contenido.insert(inicio, "\n".join(resultado))
    cuadro_contenido.focus_set()


def aplicar_marcador_atajo(event, marcador: str):
    """Aplicar formato con atajo y consumir el evento de teclado."""
    aplicar_marcador(marcador)
    return "break"


def manejar_enter_editor(event=None):
    """Continuar listas numeradas o viñetas al pulsar Enter al final de línea."""
    widget = cuadro_contenido
    try:
        widget.index("sel.first")
        return None
    except tk.TclError:
        pass

    if widget.compare("insert", "!=", "insert lineend"):
        return None

    linea = widget.get("insert linestart", "insert")
    numeracion = re.match(r"^([ \t]*)(\d+)([.)])\s+(.*)$", linea)
    if numeracion:
        sangria, numero, separador, texto = numeracion.groups()
        if not texto.strip():
            widget.delete("insert linestart", "insert")
            widget.insert("insert", "\n")
            return "break"
        widget.insert("insert", f"\n{sangria}{int(numero) + 1}{separador} ")
        return "break"

    viñeta = re.match(r"^([ \t]*)([-*])\s+(.*)$", linea)
    if viñeta:
        sangria, simbolo, texto = viñeta.groups()
        if not texto.strip():
            widget.delete("insert linestart", "insert")
            widget.insert("insert", "\n")
            return "break"
        widget.insert("insert", f"\n{sangria}{simbolo} ")
        return "break"
    return None


def limpiar_campos() -> None:
    """Limpiar el editor, solicitando confirmación si hay contenido."""
    tiene_contenido = bool(
        entrada_titulo.get().strip()
        or cuadro_contenido.get("1.0", "end-1c").strip()
    )
    if tiene_contenido:
        confirmar = messagebox.askyesno(
            "Confirmar limpieza",
            "¿Deseas borrar el título y todo el contenido escrito?",
            parent=ventana,
        )
        if not confirmar:
            return

    entrada_titulo.delete(0, tk.END)
    cuadro_contenido.delete("1.0", tk.END)
    etiqueta_estado.config(text="Listo para crear un documento.")
    entrada_titulo.focus_set()


def _es_linea_tabla(texto: str) -> bool:
    """Indicar si una línea tiene estructura de fila de tabla.

    Se exigen al menos dos barras verticales (por ejemplo ``| A | B |`` o
    ``A | B | C``). Así una barra suelta dentro de una frase no convierte el
    párrafo en tabla. Un marcador de imagen nunca se trata como fila de tabla.
    """
    return texto.count("|") >= 2 and not texto.startswith("{{imagen:")


def _celdas_tabla(linea: str) -> list[str]:
    """Separar una fila de tabla por ``|`` y limpiar solo el espacio exterior."""
    texto = linea.strip()
    if texto.startswith("|"):
        texto = texto[1:]
    if texto.endswith("|"):
        texto = texto[:-1]
    return [celda.strip() for celda in texto.split("|")]


def _bloque_tabla(filas: list[list[str]]) -> dict[str, object]:
    """Validar las filas y construir el bloque ``table`` del generador."""
    if not filas:
        raise ValueError("La tabla está vacía.")

    columnas = len(filas[0])
    if columnas < 2:
        raise ValueError("La tabla necesita al menos dos columnas separadas por «|».")

    if not any(celda for celda in filas[0]):
        raise ValueError("La cabecera de la tabla no puede estar vacía.")

    for numero, fila in enumerate(filas[1:], start=2):
        if len(fila) != columnas:
            raise ValueError(
                f"La fila {numero} de la tabla tiene {len(fila)} celdas; "
                f"la cabecera tiene {columnas}."
            )

    return {"tipo": "table", "encabezados": filas[0], "filas": filas[1:]}

# Extensiones de imagen que python-docx puede incrustar en un DOCX.
EXTENSIONES_IMAGEN = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif", ".webp"}


def _es_linea_imagen(texto: str) -> bool:
    """Indicar si una línea es una referencia de imagen insertada por el botón.

    El marcador empieza por ``{{imagen:`` y termina por ``}}``. Nunca se
    confunde con un párrafo normal y admite espacios, acentos y caracteres
    especiales en la ruta del archivo.
    """
    return texto.startswith("{{imagen:") and texto.endswith("}}")


def _partes_imagen(texto: str) -> tuple[str, str, str]:
    """Extraer ruta, ancho y leyenda de una referencia de imagen.

    Formato: ``{{imagen: ruta | ancho | leyenda}}``. La ruta es obligatoria;
    el ancho y la leyenda son opcionales.
    """
    contenido = texto[len("{{imagen:"):-len("}}")].strip()
    partes = [parte.strip() for parte in contenido.split(" | ")]
    if not partes or not partes[0]:
        raise ValueError("La referencia de imagen necesita una ruta de archivo.")
    ruta = partes[0]
    ancho_texto = partes[1] if len(partes) > 1 else "5.5"
    leyenda = " | ".join(partes[2:]).strip() if len(partes) > 2 else ""
    return ruta, ancho_texto, leyenda


def _validar_archivo_imagen(archivo: Path) -> None:
    """Comprobar que el archivo exista, sea legible y tenga formato de imagen."""
    if not archivo.is_file():
        raise ValueError(f"No se encontró la imagen: {archivo}")
    if archivo.suffix.lower() not in EXTENSIONES_IMAGEN:
        raise ValueError(
            f"Formato de imagen no compatible: {archivo.suffix}. "
            f"Use: {', '.join(sorted(EXTENSIONES_IMAGEN))}."
        )
    try:
        with archivo.open("rb") as fichero:
            fichero.read(16)
    except OSError as error:
        raise ValueError(f"No se puede leer la imagen: {error}")


def _bloque_imagen(texto: str) -> dict[str, object]:
    """Construir y validar un bloque ``image`` a partir de una línea del editor."""
    ruta, ancho_texto, leyenda = _partes_imagen(texto)

    archivo = Path(ruta)
    _validar_archivo_imagen(archivo)

    try:
        ancho = float(ancho_texto)
    except ValueError:
        raise ValueError(f"El ancho de la imagen debe ser un número: {ancho_texto}")
    if ancho <= 0:
        raise ValueError("El ancho de la imagen debe ser mayor que cero.")

    bloque: dict[str, object] = {"tipo": "image", "ruta": str(archivo), "ancho": ancho}
    if leyenda:
        bloque["leyenda"] = leyenda
    return bloque


def _es_linea_ecuacion_display(texto: str) -> bool:
    """Indicar si una línea es una ecuación independiente $$...$$."""
    return texto.startswith("$$") and texto.endswith("$$") and len(texto) > 4


def _contenido_ecuacion_display(texto: str) -> str:
    """Extraer la expresión LaTeX de una línea $$...$$."""
    return texto[2:-2].strip()


def _dividir_ecuaciones_en_linea(texto: str) -> list[dict[str, str]] | None:
    """Dividir una línea en partes de texto y ecuaciones en línea $...$.

    Devuelve ``None`` cuando no hay delimitadores ``$...$`` válidos.
    """
    partes_raw = re.split(r"(\$[^$]+\$)", texto)
    # Si re.split no encontró ningún patrón válido, devuelve la cadena completa
    # como un solo elemento; no hay ecuación que dividir.
    if len(partes_raw) == 1:
        return None
    partes: list[dict[str, str]] = []
    for parte in partes_raw:
        if not parte:
            continue
        if parte.startswith("$") and parte.endswith("$") and len(parte) > 2:
            partes.append(
                {"tipo": "ecuacion", "modo": "inline", "latex": parte[1:-1]}
            )
        else:
            partes.append({"tipo": "texto", "contenido": parte})
    return partes


def _bloque_ecuacion(texto: str) -> dict[str, object]:
    """Construir un bloque ``equation`` a partir de una línea $$...$$."""
    latex = _contenido_ecuacion_display(texto)
    return {"tipo": "equation", "modo": "display", "latex": latex}


def _convertir_a_bloques(contenido: str) -> list[dict[str, object]]:
    """Convertir marcas de texto del editor en bloques profesionales."""
    bloques: list[dict[str, object]] = []
    lineas = contenido.splitlines()
    indice = 0

    while indice < len(lineas):
        texto = lineas[indice].strip()
        if not texto:
            indice += 1
            continue

        if texto.startswith("### "):
            bloques.append({"tipo": "heading", "nivel": 3, "texto": texto[4:].strip()})
        elif texto.startswith("## "):
            bloques.append({"tipo": "heading", "nivel": 2, "texto": texto[3:].strip()})
        elif texto.startswith("# "):
            bloques.append({"tipo": "heading", "nivel": 1, "texto": texto[2:].strip()})
        elif texto.startswith(("- ", "* ")):
            bloques.append({"tipo": "bullet", "texto": texto[2:].strip()})
        elif re.match(r"^\d+[.)]\s+", texto):
            elemento = re.sub(r"^\d+[.)]\s+", "", texto, count=1)
            bloques.append({"tipo": "numbered", "texto": elemento})
        elif texto.startswith("> "):
            bloques.append({"tipo": "quote", "texto": texto[2:].strip()})
        elif _es_linea_imagen(texto):
            bloques.append(_bloque_imagen(texto))
        elif _es_linea_ecuacion_display(texto):
            bloques.append(_bloque_ecuacion(texto))
        elif "$" in texto:
            # Línea con posibles ecuaciones en línea $...$: dividir en partes.
            partes = _dividir_ecuaciones_en_linea(texto)
            if partes is None:
                # No hay delimitadores $...$ válidos: párrafo normal.
                bloques.append({"tipo": "paragraph", "texto": texto})
            else:
                bloques.append({"tipo": "paragraph", "partes": partes})
        elif _es_linea_tabla(texto):
            filas: list[list[str]] = []
            while indice < len(lineas):
                candidata = lineas[indice].strip()
                if not _es_linea_tabla(candidata):
                    break
                filas.append(_celdas_tabla(candidata))
                indice += 1
            bloques.append(_bloque_tabla(filas))
            continue
        else:
            bloques.append({"tipo": "paragraph", "texto": texto})

        indice += 1

    return bloques


def insertar_tabla() -> None:
    """Insertar en el cursor una tabla de 2 columnas con cabecera y 2 filas."""
    if cuadro_contenido is None:
        return

    posicion = cuadro_contenido.index("insert")
    texto = PLANTILLA_TABLA
    if cuadro_contenido.compare(posicion, "!=", "insert linestart"):
        texto = "\n" + texto
    if cuadro_contenido.compare(posicion, "!=", "insert lineend"):
        texto = texto + "\n"

    cuadro_contenido.insert(posicion, texto)
    cuadro_contenido.focus_set()
    if etiqueta_estado is not None:
        etiqueta_estado.config(
            text="Tabla insertada: sustituye las celdas y añade filas con «|»."
        )


def insertar_imagen() -> None:
    """Seleccionar una imagen, configurar ancho y leyenda, e insertar su referencia en el cursor."""
    if cuadro_contenido is None:
        return

    # 1. Seleccionar el archivo de imagen.
    ruta = filedialog.askopenfilename(
        parent=ventana,
        title="Seleccionar imagen",
        filetypes=[
            ("Imágenes", "*.png *.jpg *.jpeg *.gif *.bmp *.tiff *.tif *.webp"),
            ("Todos los archivos", "*.*"),
        ],
    )
    if not ruta:
        return  # El usuario canceló el selector.

    archivo = Path(ruta)
    try:
        _validar_archivo_imagen(archivo)
    except ValueError as error:
        messagebox.showerror("Imagen no válida", str(error), parent=ventana)
        return

    # 2. Diálogo para configurar ancho y leyenda.
    resultado: dict[str, object] = {}

    dialogo = tk.Toplevel(ventana)
    dialogo.title("Insertar imagen")
    dialogo.transient(ventana)
    dialogo.grab_set()

    panel = ttk.Frame(dialogo, padding=12)
    panel.pack(fill="both", expand=True)

    ttk.Label(panel, text=f"Archivo: {ruta}").pack(anchor="w", pady=(0, 10))

    ttk.Label(panel, text="Ancho (pulgadas):").pack(anchor="w")
    ancho_var = tk.StringVar(value="5.5")
    entrada_ancho = ttk.Entry(panel, textvariable=ancho_var, width=15)
    entrada_ancho.pack(anchor="w", pady=(0, 10))

    ttk.Label(panel, text="Leyenda (opcional):").pack(anchor="w")
    leyenda_var = tk.StringVar(value="")
    ttk.Entry(panel, textvariable=leyenda_var, width=40).pack(anchor="w", pady=(0, 10))

    def aceptar() -> None:
        try:
            ancho = float(ancho_var.get().strip())
            if ancho <= 0:
                raise ValueError()
        except ValueError:
            messagebox.showerror(
                "Ancho inválido",
                "El ancho debe ser un número mayor que cero.",
                parent=dialogo,
            )
            return
        resultado["ancho"] = ancho
        resultado["leyenda"] = leyenda_var.get().strip()
        dialogo.destroy()

    def cancelar() -> None:
        dialogo.destroy()

    botones = ttk.Frame(panel)
    botones.pack(anchor="e", pady=(10, 0))
    ttk.Button(botones, text="Cancelar", command=cancelar).pack(side="right", padx=(8, 0))
    ttk.Button(botones, text="Insertar", command=aceptar).pack(side="right")

    entrada_ancho.focus_set()
    ventana.wait_window(dialogo)

    if "ancho" not in resultado:
        return  # El usuario canceló la configuración.

    ancho = float(resultado["ancho"])  # type: ignore[arg-type]
    leyenda = str(resultado["leyenda"])  # type: ignore[arg-type]

    # 3. Insertar la referencia interna en la posición del cursor.
    posicion = cuadro_contenido.index("insert")
    marcador = f"{{{{imagen: {ruta} | {ancho} | {leyenda}}}}}"
    if cuadro_contenido.compare(posicion, "!=", "insert linestart"):
        marcador = "\n" + marcador
    if cuadro_contenido.compare(posicion, "!=", "insert lineend"):
        marcador = marcador + "\n"

    cuadro_contenido.insert(posicion, marcador)
    cuadro_contenido.focus_set()
    if etiqueta_estado is not None:
        etiqueta_estado.config(text="Imagen insertada: referencia lista para generar el documento.")


def mostrar_ayuda_formulas() -> None:
    """Mostrar la ayuda de sintaxis para fórmulas LaTeX."""
    ayuda = r"""Fórmulas matemáticas LaTeX → ecuaciones nativas de Word (OMML)

Sintaxis admitida:
  • En línea:  $f(x)=x^2+3x+2$
    Se inserta dentro del párrafo actual.

  • Independiente: $$\int_0^1 x^2\,dx=\frac{1}{3}$$
    Ocupa un párrafo propio, centrado.

Ejemplos de expresiones:
  • Fracción:    \frac{a+b}{c}
  • Potencia:    x^2  o  x^{2}
  • Subíndice:   a_i  o  a_{i}
  • Raíz:        \sqrt{x}
  • Integral:    \int_0^1 x^2\,dx
  • Sumatoria:   \sum_{i=1}^{n} i
  • Límite:      \lim_{x \to \infty} f(x)
  • Funciones:   \sin, \cos, \log, \ln, \exp
  • Espacio fino:\,

Las ecuaciones se guardan como objetos OMML editables en Word, no como imágenes."""
    if ventana is not None:
        messagebox.showinfo("Fórmulas LaTeX", ayuda, parent=ventana)
    else:
        messagebox.showinfo("Fórmulas LaTeX", ayuda)


def abrir_configuracion() -> None:
    """Mostrar un cuadro de diálogo para configurar un documento profesional."""
    global configuracion_actual

    dialogo = tk.Toplevel(ventana)
    dialogo.title("Configuración profesional del documento")
    dialogo.geometry("680x540")
    dialogo.minsize(610, 500)
    dialogo.transient(ventana)
    dialogo.grab_set()

    panel = ttk.Frame(dialogo, padding=12)
    panel.pack(fill="both", expand=True)
    panel.columnconfigure(0, weight=1)

    pestañas = ttk.Notebook(panel)
    pestañas.pack(fill="both", expand=True)

    pagina = ttk.Frame(pestañas, padding=12)
    encabezado_pie = ttk.Frame(pestañas, padding=12)
    portada = ttk.Frame(pestañas, padding=12)
    pestañas.add(pagina, text="Página y texto")
    pestañas.add(encabezado_pie, text="Encabezado y pie")
    pestañas.add(portada, text="Portada y autor")

    # Variables de formato
    tamano_var = tk.StringVar(value={"Letter": "Carta", "letter": "Carta", "carta": "Carta"}.get(configuracion_actual.tamano_pagina, configuracion_actual.tamano_pagina))
    orientacion_var = tk.StringVar(value="Horizontal" if configuracion_actual.orientacion.lower() in {"horizontal", "landscape"} else "Vertical")
    fuente_var = tk.StringVar(value=configuracion_actual.fuente)
    fuente_tamano_var = tk.StringVar(value=str(configuracion_actual.tamano_fuente))
    interlineado_var = tk.StringVar(value=str(configuracion_actual.interlineado))

    def campo_etiquetado(parent, etiqueta, variable, fila, columna, opciones=None, ancho=18):
        ttk.Label(parent, text=etiqueta).grid(row=fila, column=columna, sticky="w", padx=5, pady=6)
        if opciones is not None:
            control = ttk.Combobox(parent, textvariable=variable, values=opciones, state="readonly", width=ancho)
        else:
            control = ttk.Entry(parent, textvariable=variable, width=ancho)
        control.grid(row=fila, column=columna + 1, sticky="ew", padx=5, pady=6)
        return control

    pagina.columnconfigure(1, weight=1)
    pagina.columnconfigure(3, weight=1)
    campo_etiquetado(pagina, "Tamaño de página", tamano_var, 0, 0, ["A4", "Carta", "Legal"])
    campo_etiquetado(pagina, "Orientación", orientacion_var, 0, 2, ["Vertical", "Horizontal"])
    campo_etiquetado(pagina, "Fuente", fuente_var, 1, 0, ["Arial", "Calibri", "Aptos", "Times New Roman", "Cambria", "Georgia"])
    campo_etiquetado(pagina, "Tamaño de letra (pt)", fuente_tamano_var, 1, 2)
    campo_etiquetado(pagina, "Interlineado", interlineado_var, 2, 0, ["1.0", "1.15", "1.5", "2.0"])

    ttk.Separator(pagina).grid(row=3, column=0, columnspan=4, sticky="ew", pady=12)
    ttk.Label(pagina, text="Márgenes en centímetros").grid(row=4, column=0, columnspan=4, sticky="w", padx=5, pady=(0, 6))
    margenes_cm = {
        "superior": tk.StringVar(value=f"{configuracion_actual.margen_superior * 2.54:.2f}"),
        "inferior": tk.StringVar(value=f"{configuracion_actual.margen_inferior * 2.54:.2f}"),
        "izquierdo": tk.StringVar(value=f"{configuracion_actual.margen_izquierdo * 2.54:.2f}"),
        "derecho": tk.StringVar(value=f"{configuracion_actual.margen_derecho * 2.54:.2f}"),
    }
    campo_etiquetado(pagina, "Superior", margenes_cm["superior"], 5, 0)
    campo_etiquetado(pagina, "Inferior", margenes_cm["inferior"], 5, 2)
    campo_etiquetado(pagina, "Izquierdo", margenes_cm["izquierdo"], 6, 0)
    campo_etiquetado(pagina, "Derecho", margenes_cm["derecho"], 6, 2)

    # Encabezado y pie
    encabezado_pie.columnconfigure(1, weight=1)
    texto_encabezado_var = tk.StringVar(value=configuracion_actual.texto_encabezado or "")
    texto_pie_var = tk.StringVar(value=configuracion_actual.texto_pie or "")
    numeracion_var = tk.BooleanVar(value=configuracion_actual.numeracion_paginas)
    logo_var = tk.StringVar(value=str(configuracion_actual.ruta_logotipo or ""))

    ttk.Label(encabezado_pie, text="Texto del encabezado").grid(row=0, column=0, sticky="w", padx=5, pady=8)
    ttk.Entry(encabezado_pie, textvariable=texto_encabezado_var).grid(row=0, column=1, columnspan=2, sticky="ew", padx=5, pady=8)
    ttk.Label(encabezado_pie, text="Texto del pie de página").grid(row=1, column=0, sticky="w", padx=5, pady=8)
    ttk.Entry(encabezado_pie, textvariable=texto_pie_var).grid(row=1, column=1, columnspan=2, sticky="ew", padx=5, pady=8)
    ttk.Checkbutton(encabezado_pie, text="Añadir numeración de páginas", variable=numeracion_var).grid(row=2, column=0, columnspan=3, sticky="w", padx=5, pady=8)

    ttk.Label(encabezado_pie, text="Logotipo institucional").grid(row=3, column=0, sticky="w", padx=5, pady=8)
    ttk.Entry(encabezado_pie, textvariable=logo_var).grid(row=3, column=1, sticky="ew", padx=5, pady=8)

    def seleccionar_logo():
        ruta = filedialog.askopenfilename(
            parent=dialogo,
            title="Seleccionar logotipo",
            filetypes=[("Imágenes", "*.png *.jpg *.jpeg *.bmp *.gif"), ("Todos los archivos", "*.*")],
        )
        if ruta:
            logo_var.set(ruta)

    ttk.Button(encabezado_pie, text="Examinar…", command=seleccionar_logo).grid(row=3, column=2, padx=5, pady=8)
    ttk.Label(encabezado_pie, text="El logotipo se insertará en el encabezado.", foreground="#666666").grid(row=4, column=0, columnspan=3, sticky="w", padx=5, pady=4)

    # Portada y metadatos
    portada.columnconfigure(1, weight=1)
    portada_var = tk.BooleanVar(value=configuracion_actual.crear_portada)
    subtitulo_portada_var = tk.StringVar(value=configuracion_actual.subtitulo_portada or "")
    institucion_var = tk.StringVar(value=configuracion_actual.institucion or "")
    autor_var = tk.StringVar(value=configuracion_actual.autor or "")

    ttk.Checkbutton(portada, text="Crear portada independiente", variable=portada_var).grid(row=0, column=0, columnspan=2, sticky="w", padx=5, pady=8)
    ttk.Label(portada, text="Subtítulo de portada").grid(row=1, column=0, sticky="w", padx=5, pady=8)
    ttk.Entry(portada, textvariable=subtitulo_portada_var).grid(row=1, column=1, sticky="ew", padx=5, pady=8)
    ttk.Label(portada, text="Institución").grid(row=2, column=0, sticky="w", padx=5, pady=8)
    ttk.Entry(portada, textvariable=institucion_var).grid(row=2, column=1, sticky="ew", padx=5, pady=8)
    ttk.Label(portada, text="Autor").grid(row=3, column=0, sticky="w", padx=5, pady=8)
    ttk.Entry(portada, textvariable=autor_var).grid(row=3, column=1, sticky="ew", padx=5, pady=8)

    botones = ttk.Frame(panel)
    botones.pack(fill="x", pady=(10, 0))

    def guardar_configuracion():
        global configuracion_actual
        try:
            tamano = {"A4": "A4", "Carta": "Letter", "Legal": "Legal"}[tamano_var.get()]
            orientacion = "horizontal" if orientacion_var.get() == "Horizontal" else "vertical"
            tamano_fuente = float(fuente_tamano_var.get())
            interlineado = float(interlineado_var.get())
            cm = {nombre: float(variable.get()) for nombre, variable in margenes_cm.items()}
            if any(valor < 0 for valor in cm.values()):
                raise ValueError("Los márgenes no pueden ser negativos.")
            if tamano_fuente <= 0 or interlineado <= 0:
                raise ValueError("El tamaño de letra y el interlineado deben ser mayores que cero.")

            ruta_logo = logo_var.get().strip() or None
            if ruta_logo and not Path(ruta_logo).is_file():
                raise FileNotFoundError(f"No se encontró el logotipo: {ruta_logo}")

            configuracion_actual = ConfiguracionDocumento(
                tamano_pagina=tamano,
                orientacion=orientacion,
                fuente=fuente_var.get().strip(),
                tamano_fuente=tamano_fuente,
                interlineado=interlineado,
                margen_superior=cm["superior"] / 2.54,
                margen_inferior=cm["inferior"] / 2.54,
                margen_izquierdo=cm["izquierdo"] / 2.54,
                margen_derecho=cm["derecho"] / 2.54,
                texto_encabezado=texto_encabezado_var.get().strip() or None,
                texto_pie=texto_pie_var.get().strip() or None,
                numeracion_paginas=numeracion_var.get(),
                ruta_logotipo=ruta_logo,
                autor=autor_var.get().strip() or None,
                crear_portada=portada_var.get(),
                subtitulo_portada=subtitulo_portada_var.get().strip() or None,
                institucion=institucion_var.get().strip() or None,
            )
            if etiqueta_estado is not None:
                etiqueta_estado.config(text="Configuración profesional actualizada.")
            dialogo.destroy()
        except (ValueError, TypeError, FileNotFoundError) as error:
            messagebox.showerror("Configuración inválida", str(error), parent=dialogo)

    ttk.Button(botones, text="Cancelar", command=dialogo.destroy).pack(side="right", padx=(8, 0))
    ttk.Button(botones, text="Guardar configuración", command=guardar_configuracion).pack(side="right")


def generar_documento() -> None:
    """Validar el contenido y crear un documento DOCX profesional."""
    titulo = entrada_titulo.get().strip()
    contenido = cuadro_contenido.get("1.0", "end-1c").strip()

    if not titulo:
        messagebox.showwarning("Datos obligatorios", "Escribe el título del documento.", parent=ventana)
        entrada_titulo.focus_set()
        return
    if not contenido:
        messagebox.showwarning("Datos obligatorios", "Escribe el contenido del documento.", parent=ventana)
        cuadro_contenido.focus_set()
        return

    try:
        bloques = _convertir_a_bloques(contenido)
    except ValueError as error:
        messagebox.showerror("Contenido no válido", str(error), parent=ventana)
        cuadro_contenido.focus_set()
        return

    if not bloques:
        messagebox.showwarning("Contenido vacío", "Escribe al menos un párrafo o elemento.", parent=ventana)
        return

    nombre = re.sub(r"[^\w\s-]", "", titulo.lower())
    nombre = re.sub(r"[\s-]+", "_", nombre).strip("_") or "documento"

    try:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        ruta_seleccionada = filedialog.asksaveasfilename(
            parent=ventana,
            title="Guardar documento de Word",
            initialdir=str(OUTPUT_DIR),
            initialfile=f"{nombre}.docx",
            defaultextension=".docx",
            filetypes=[("Documentos de Word", "*.docx")],
        )
        if not ruta_seleccionada:
            etiqueta_estado.config(text="Guardado cancelado.")
            return

        ruta = crear_documento_profesional(
            titulo,
            bloques,
            Path(ruta_seleccionada),
            configuracion=configuracion_actual,
        )
        etiqueta_estado.config(text=f"Documento guardado: {ruta.name}")
        messagebox.showinfo(
            "Word-Architect",
            "¡Documento profesional creado correctamente!\n\n"
            f"Nombre: {ruta.name}\n\nUbicación:\n{ruta}",
            parent=ventana,
        )
    except Exception as error:
        messagebox.showerror(
            "Error al crear el documento",
            f"No se pudo generar el documento.\n\n{error}",
            parent=ventana,
        )


def main() -> None:
    """Iniciar la interfaz gráfica adaptable."""
    global ventana, entrada_titulo, cuadro_contenido, etiqueta_estado

    ventana = tk.Tk()
    ventana.title("Word-Architect | Generador de documentos")
    ventana.geometry("780x630")
    ventana.minsize(590, 480)
    ventana.columnconfigure(0, weight=1)
    ventana.rowconfigure(0, weight=1)

    marco = ttk.Frame(ventana, padding=12)
    marco.grid(row=0, column=0, sticky="nsew")
    marco.columnconfigure(0, weight=1)
    marco.rowconfigure(6, weight=1, minsize=120)

    ttk.Label(marco, text="WORD-ARCHITECT", font=("Arial", 20, "bold")).grid(
        row=0, column=0, sticky="w", pady=(0, 3)
    )
    ttk.Label(marco, text="Generador de documentos Microsoft Word").grid(
        row=1, column=0, sticky="w", pady=(0, 12)
    )

    ttk.Label(marco, text="Título del documento:").grid(row=2, column=0, sticky="w", pady=(0, 4))
    entrada_titulo = ttk.Entry(marco, font=("Arial", 11))
    entrada_titulo.grid(row=3, column=0, sticky="ew", pady=(0, 10))
    ttk.Label(marco, text="Contenido del documento:").grid(row=4, column=0, sticky="w")

    herramientas = ttk.Frame(marco)
    herramientas.grid(row=5, column=0, sticky="ew", pady=6)
    opciones_formato = [
        ("H1  Título", "# ", "Aplicar un título principal de nivel 1."),
        ("H2  Subtítulo", "## ", "Aplicar un subtítulo de nivel 2."),
        ("H3  Sección", "### ", "Aplicar un encabezado de nivel 3."),
        ("•  Viñeta", "- ", "Crear una lista con viñetas; Enter continúa la lista."),
        ("1.  Numeración", "1. ", "Crear una lista numerada; Enter incrementa el número."),
    ]
    botones_formato = []
    tooltips = []
    for texto, marcador, descripcion in opciones_formato:
        boton = ttk.Button(herramientas, text=texto, command=lambda valor=marcador: aplicar_marcador(valor))
        botones_formato.append(boton)
        tooltips.append(ToolTip(boton, descripcion))

    boton_tabla = ttk.Button(herramientas, text="▦  Tabla", command=insertar_tabla)
    botones_formato.append(boton_tabla)
    tooltips.append(
        ToolTip(
            boton_tabla,
            "Insertar una tabla editable de 2 columnas.\n"
            "Cada línea es una fila y las celdas se separan con «|».\n"
            "La primera fila se usa como cabecera.",
        )
    )

    boton_imagen = ttk.Button(herramientas, text="🖼 Imagen", command=insertar_imagen)
    botones_formato.append(boton_imagen)
    tooltips.append(
        ToolTip(
            boton_imagen,
            "Insertar una imagen en el documento.\n"
            "Selecciona el archivo, el ancho y una leyenda opcional.\n"
            "La referencia se guarda como {{imagen: ruta | ancho | leyenda}}.",
        )
    )

    boton_formulas = ttk.Button(
        herramientas, text="📐 Fórmulas LaTeX", command=mostrar_ayuda_formulas
    )
    botones_formato.append(boton_formulas)
    tooltips.append(
        ToolTip(
            boton_formulas,
            r"""Ayuda para fórmulas matemáticas LaTeX.
En línea:  $f(x)=x^2+3x+2$
Independiente: $$\int_0^1 x^2\,dx=\frac{1}{3}$$
Las ecuaciones se insertan como objetos nativos de Word (OMML), editables.""",
        )
    )

    distribucion_actual = {"columnas": 0}

    def ajustar_herramientas(event=None):
        ancho = herramientas.winfo_width()
        columnas = 2 if ancho < 620 else len(botones_formato)
        if distribucion_actual["columnas"] == columnas:
            return
        for boton in botones_formato:
            boton.grid_forget()
        for columna in range(len(botones_formato)):
            herramientas.columnconfigure(columna, weight=0)
        for columna in range(columnas):
            herramientas.columnconfigure(columna, weight=1)
        for indice, boton in enumerate(botones_formato):
            boton.grid(row=indice // columnas, column=indice % columnas, sticky="ew", padx=3, pady=2)
        distribucion_actual["columnas"] = columnas

    herramientas.bind("<Configure>", ajustar_herramientas)

    editor = ttk.Frame(marco)
    editor.grid(row=6, column=0, sticky="nsew", pady=(0, 10))
    editor.columnconfigure(0, weight=1)
    editor.rowconfigure(0, weight=1)
    cuadro_contenido = tk.Text(editor, wrap="word", font=("Arial", 11), undo=True, padx=8, pady=8)
    cuadro_contenido.grid(row=0, column=0, sticky="nsew")
    cuadro_contenido.bind("<Return>", manejar_enter_editor)
    for secuencia, marcador in [
        ("<Control-Key-1>", "# "),
        ("<Control-Key-2>", "## "),
        ("<Control-Key-3>", "### "),
        ("<Control-Key-4>", "- "),
        ("<Control-Key-5>", "1. "),
    ]:
        cuadro_contenido.bind(
            secuencia,
            lambda event, valor=marcador: aplicar_marcador_atajo(event, valor),
        )

    barra = ttk.Scrollbar(editor, orient="vertical", command=cuadro_contenido.yview)
    barra.grid(row=0, column=1, sticky="ns")
    cuadro_contenido.configure(yscrollcommand=barra.set)

    acciones = ttk.Frame(marco)
    acciones.grid(row=7, column=0, sticky="ew", pady=(0, 8))
    boton_configuracion = ttk.Button(
        acciones,
        text="⚙ Configuración",
        command=abrir_configuracion,
    )
    boton_configuracion.pack(side="left", padx=(0, 8))
    tooltips.append(ToolTip(boton_configuracion, "Configurar página, fuente, márgenes, portada, logotipo y numeración."))

    boton_guardar = ttk.Button(acciones, text="▣ Guardar Word", command=generar_documento)
    boton_guardar.pack(side="left", padx=(0, 8))
    tooltips.append(ToolTip(boton_guardar, "Generar el documento profesional y elegir dónde guardarlo."))

    boton_limpiar = ttk.Button(acciones, text="⌫ Limpiar", command=limpiar_campos)
    boton_limpiar.pack(side="left")
    tooltips.append(ToolTip(boton_limpiar, "Borrar título y contenido, con confirmación."))

    etiqueta_estado = ttk.Label(marco, text="Listo para crear un documento.")
    etiqueta_estado.grid(row=8, column=0, sticky="ew")
    entrada_titulo.focus_set()
    ventana.mainloop()


if __name__ == "__main__":
    main()
