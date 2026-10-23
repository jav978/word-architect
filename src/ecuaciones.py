r"""Conversión de expresiones LaTeX a Office Math Markup Language (OMML).

Permite incrustar ecuaciones nativas y editables de Microsoft Word en los
documentos generados por Word-Architect, sin recurrir a imágenes ni capturas
de pantalla.

Este módulo NO ejecuta comandos LaTeX arbitrarios ni requiere Pandoc ni
herramientas externas. Convierte un subconjunto controlado de LaTeX a OMML
mediante un analizador sintáctico puro en Python y lxml, que ya es una
dependencia del proyecto.

Elementos soportados (primer incremento):
  * Fracciones: \frac{a+b}{c}
  * Potencias y subíndices: x^2, a_i, x^{2} y a_{i}
  * Raíces cuadradas: \sqrt{x} (y enésimas con \sqrt[n]{x} no soportado aún)
  * Integrales: \int_0^1 x^2 \,dx
  * Sumatorias: \sum_{i=1}^{n} i
  * Límites: \lim_{x \to \infty}
  * Derivadas: \frac{d}{dx} f(x), f'(x)
  * Vectores y matrices: \begin{matrix} ... \end{matrix}
  * Operadores y letras griegas comunes
  * Espacio fino \, y espacio normal \

Sintaxis reconocida por el editor (gui.py):
  * En línea:  $f(x)=x^2+3x+2$
  * Independiente: $$\int_0^1 x^2\,dx=\frac{1}{3}$$
"""

from __future__ import annotations

from lxml import etree

# ---------------------------------------------------------------------------
# Espacios de nombres OMML y WordprocessingML
# ---------------------------------------------------------------------------

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _m(tag: str) -> str:
    """Etiqueta con el espacio de nombres OMML."""
    return f"{{{M_NS}}}{tag}"


def _w(tag: str) -> str:
    """Etiqueta con el espacio de nombres WordprocessingML."""
    return f"{{{W_NS}}}{tag}"


# ---------------------------------------------------------------------------
# Excepciones
# ---------------------------------------------------------------------------


class LatexConversionError(ValueError):
    """Error al convertir una expresión LaTeX a OMML."""


# ---------------------------------------------------------------------------
# Analizador sintáctico LaTeX (modo matemático)
# ---------------------------------------------------------------------------


class _Parser:
    """Analizador recursivo descendente para expresiones LaTeX matemáticas.

    Produce un árbol de sintaxis abstracta (AST) representado como tuplas
    anidadas. No ejecuta código del usuario; solo reconoce una gramática
    acotada y segura.
    """

    def __init__(self, texto: str):
        self.texto = texto
        self.pos = 0
        self.longitud = len(texto)

    # --- Acceso a caracteres ----------------------------------------------

    def _actual(self) -> str | None:
        if self.pos < self.longitud:
            return self.texto[self.pos]
        return None

    def _avanzar(self) -> str:
        c = self.texto[self.pos]
        self.pos += 1
        return c

    def _esperar(self, esperado: str) -> None:
        actual = self._actual()
        if actual != esperado:
            raise LatexConversionError(
                f"Se esperaba '{esperado}' pero se encontró {actual!r} "
                f"en la posición {self.pos}."
            )
        self._avanzar()

    def _skip_espacios(self) -> None:
        while self._actual() == " ":
            self._avanzar()

    # --- Entrada principal ------------------------------------------------

    def parse(self) -> object:
        self._skip_espacios()
        nodo = self._expresion()
        self._skip_espacios()
        if self.pos < self.longitud:
            sobrante = self.texto[self.pos:]
            raise LatexConversionError(
                f"Caracteres no esperados al final: {sobrante!r}"
            )
        return nodo

    # --- Gramática --------------------------------------------------------
    # expresion  := termino (('+' | '-' | '=') termino)*
    # termino    := factor (factor)*          # yuxtaposición = multiplicación
    # factor     | atomo ('^' grupo)? ('_' grupo)?
    # atomo      | '{' expresion '}' | '(' expresion ')' | comando
    #            | numero | variable | operador | grupo
    # comando    | \frac{a}{b} | \sqrt{x} | \int_^ | \sum_^ | \lim_^
    #            | \sin, \cos, ... | \, | \  | símbolos griegos

    def _expresion(self) -> object | None:
        self._skip_espacios()
        nodo = self._termino()
        if nodo is None:
            return None
        while self._actual() in ("+", "-", "="):
            op = self._avanzar()
            self._skip_espacios()
            der = self._termino()
            if der is None:
                break
            nodo = ("infix", op, nodo, der)
            self._skip_espacios()
        return nodo

    def _termino(self) -> object | None:
        self._skip_espacios()
        factores: list = []
        while True:
            if self._actual() == "=":
                break
            factor = self._factor()
            if factor is None:
                break
            factores.append(factor)
            self._skip_espacios()
            if self._actual() == "=":
                break
        if not factores:
            return None
        if len(factores) == 1:
            return factores[0]
        return ("mul", factores)

    def _factor(self) -> object | None:
        self._skip_espacios()
        base = self._atomo()
        if base is None:
            return None
        while self._actual() in ("^", "_"):
            tipo = self._avanzar()  # "^" o "_"
            self._skip_espacios()
            grupo = self._grupo_o_atomo()
            if grupo is None:
                raise LatexConversionError(
                    f"Se esperaba un grupo o atomo después de '{tipo}'."
                )
            if tipo == "^":
                base = ("sup", base, grupo)
            else:
                base = ("sub", base, grupo)
        return base

    def _grupo_o_atomo(self) -> object | None:
        c = self._actual()
        if c == "{":
            self._avanzar()
            nodo = self._expresion()
            self._skip_espacios()
            self._esperar("}")
            return ("group", nodo)
        return self._atomo()

    def _atomo(self) -> object | None:
        self._skip_espacios()
        c = self._actual()
        if c is None:
            return None

        # Agrupaciones
        if c == "{":
            self._avanzar()
            nodo = self._expresion()
            self._skip_espacios()
            self._esperar("}")
            return ("group", nodo)
        if c == "(":
            self._avanzar()
            nodo = self._expresion()
            self._skip_espacios()
            self._esperar(")")
            return ("group", nodo)
        if c == "[":
            self._avanzar()
            nodo = self._expresion()
            self._skip_espacios()
            self._esperar("]")
            return ("group", nodo)

        # Comandos de barra invertida
        if c == "\\":
            return self._comando()

        # Números
        if c.isdigit():
            return self._numero()

        # Variables (letras, dígitos y prima)
        if c.isalpha():
            return self._variable()

        # Operadores simples
        if c in "+-=<>:;!,.?*/":
            self._avanzar()
            return ("operador", c)

        # Delimitadores de cierre: devolver None para que el padre gestione
        if c in "})]":
            return None

        # Cualquier otro caracter se trata como texto literal
        self._avanzar()
        return ("texto", c)

    def _numero(self) -> object:
        inicio = self.pos
        while self._actual() is not None and self._actual().isdigit():
            self._avanzar()
        if self._actual() == ".":
            self._avanzar()
            while self._actual() is not None and self._actual().isdigit():
                self._avanzar()
        return ("numero", self.texto[inicio:self.pos])

    def _variable(self) -> object:
        inicio = self.pos
        while (
            self._actual() is not None
            and (self._actual().isalpha() or self._actual().isdigit() or self._actual() == "'")
        ):
            self._avanzar()
        nombre = self.texto[inicio:self.pos]
        return ("variable", nombre)

    def _comando(self) -> object:
        r"""Parsear un comando LaTeX que empieza por \."""
        self._avanzar()  # consumir "\"
        siguiente = self._actual()
        if siguiente is None:
            raise LatexConversionError("Comando LaTeX vacío (barra invertida al final).")

        # Comandos de un solo símbolo no alfabético: \, (espacio fino),
        # \  (espacio normal), etc.
        if not (siguiente.isalpha() or siguiente == ":"):
            nombre = self._avanzar()  # un solo caracter
        else:
            inicio = self.pos
            while (
                self._actual() is not None
                and (self._actual().isalpha() or self._actual() == ":")
            ):
                self._avanzar()
            nombre = self.texto[inicio:self.pos]
            if not nombre:
                raise LatexConversionError("Comando LaTeX vacío (barra invertida sin nombre).")

        # Comandos con argumento entre llaves obligatorio
        if nombre == "sqrt":
            grado = None
            if self._actual() == "[":
                self._avanzar()
                grado = self._expresion()
                self._skip_espacios()
                self._esperar("]")
            operando = self._argumento_obligatorio()
            return ("raiz", operando, grado)

        if nombre == "frac":
            numerador = self._argumento_obligatorio()
            denominador = self._argumento_obligatorio()
            return ("fraccion", numerador, denominador)

        # Comandos con límites opcionales: \int, \sum, \lim, \prod
        if nombre in ("int", "sum", "lim", "prod"):
            inferior = None
            superior = None
            if self._actual() == "_":
                self._avanzar()
                inferior = self._grupo_o_atomo()
            if self._actual() == "^":
                self._avanzar()
                superior = self._grupo_o_atomo()
            # Capturar el integrando / cuerpo como el siguiente término
            integrando = self._termino()
            if nombre == "int":
                return ("integral", inferior, superior, integrando)
            if nombre == "sum":
                return ("suma", inferior, superior, integrando)
            if nombre == "lim":
                return ("limite", inferior, superior, integrando)
            return ("producto", inferior, superior, integrando)

        # Espacio fino \,
        if nombre == ",":
            return ("espacio_fino",)

        # Espacio normal \  (barra invertida seguida de espacio real)
        if nombre == " ":
            return ("espacio",)

        # Funciones comunes: \sin, \cos, \tan, \log, \ln, \exp
        if nombre in ("sin", "cos", "tan", "log", "ln", "exp"):
            arg = self._grupo_o_atomo()
            return ("funcion", nombre, arg)

        # Derivada parcial \partial
        if nombre == "partial":
            return ("operador", "∂")

        # Símbolos comunes
        simbolos = {
            "infty": "∞",
            "alpha": "α",
            "beta": "β",
            "gamma": "γ",
            "delta": "δ",
            "epsilon": "ε",
            "zeta": "ζ",
            "eta": "η",
            "theta": "θ",
            "lambda": "λ",
            "mu": "μ",
            "pi": "π",
            "sigma": "σ",
            "tau": "τ",
            "phi": "φ",
            "chi": "χ",
            "psi": "ψ",
            "omega": "ω",
            "cdot": "·",
            "times": "×",
            "div": "÷",
            "le": "≤",
            "ge": "≥",
            "ne": "≠",
            "in": "∈",
            "subset": "⊂",
            "nabla": "∇",
            "to": "→",
        }
        if nombre in simbolos:
            return ("texto", simbolos[nombre])

        # Matriz: \begin{matrix} ... \end{matrix}
        if nombre == "begin":
            return self._entorno_matriz()

        # Comandos desconocidos -> error informativo (no se omite silenciosamente)
        raise LatexConversionError(f"Comando LaTeX no soportado: {nombre}")

    def _entorno_matriz(self) -> object:
        r"""Parsear \begin{matrix} ... \end{matrix}."""
        self._skip_espacios()
        if self._actual() != "{":
            raise LatexConversionError(r"Se esperaba '{' después de \begin.")
        self._avanzar()
        inicio = self.pos
        while self._actual() is not None and self._actual().isalpha():
            self._avanzar()
        entorno = self.texto[inicio:self.pos]
        self._skip_espacios()
        self._esperar("}")

        filas: list = []
        while True:
            self._skip_espacios()
            if self._actual() is None:
                raise LatexConversionError(
                    r"Se esperaba \end{matrix} pero se llegó al final de la expresión."
                )
            # Detectar \end{matrix}
            if self._actual() == "\\":
                self._avanzar()
                inicio = self.pos
                while self._actual() is not None and self._actual().isalpha():
                    self._avanzar()
                cmd = self.texto[inicio:self.pos]
                if cmd == "end":
                    self._skip_espacios()
                    self._esperar("{")
                    self._skip_espacios()
                    inicio2 = self.pos
                    while self._actual() is not None and self._actual().isalpha():
                        self._avanzar()
                    nombre_end = self.texto[inicio2:self.pos]
                    self._skip_espacios()
                    self._esperar("}")
                    if nombre_end != entorno:
                        raise LatexConversionError(
                            rf"\begin{{{entorno}}} no coincide con \end{{{nombre_end}}}."
                        )
                    break
                else:
                    # No es \end, volver y tratar como comando normal
                    self.pos = inicio - 1  # retroceder a la barra invertida
                    cmd_nodo = self._comando()
                    filas.append(cmd_nodo)
                    continue

            celda = self._expresion()
            if celda is None:
                raise LatexConversionError("Celda de matriz vacía.")
            filas.append(celda)
            self._skip_espacios()
            if self._actual() == "\\":
                self._avanzar()
                self._skip_espacios()
                continue
            if self._actual() is None:
                break

        if not filas:
            raise LatexConversionError("La matriz no tiene filas.")

        # Determinar columnas por la primera fila (soportar \\ para separar filas,
        # pero dentro de una fila las celdas van separadas por & no soportado aún;
        # aquí cada línea \\ es una fila con una celda, salvo que se acepte &).
        if entorno == "matrix":
            return ("matrix", filas)
        if entorno == "pmatrix":
            return ("pmatrix", filas)
        return ("matrix", filas)

    def _argumento_obligatorio(self) -> object:
        c = self._actual()
        if c is None:
            raise LatexConversionError("Se esperaba un argumento y se llegó al final.")
        if c == "{":
            self._avanzar()
            nodo = self._expresion()
            self._skip_espacios()
            self._esperar("}")
            return ("group", nodo)
        # Sin llaves: tomar el siguiente atomo
        return self._atomo()


def latex_a_ast(texto: str) -> object:
    """Convertir una expresión LaTeX en modo matemático a un AST."""
    parser = _Parser(texto)
    return parser.parse()


# ---------------------------------------------------------------------------
# Conversión de AST a OMML (Office Math Markup Language)
# ---------------------------------------------------------------------------


def _texto_run(contenido: str):
    """Crear <m:r><m:t>contenido</m:t></m:r>."""
    r = etree.Element(_m("r"))
    t = etree.SubElement(r, _m("t"))
    t.text = contenido
    return r


def _espacio_omml():
    """Run con un espacio."""
    return _texto_run(" ")


def _envuelve_omml(hijos):
    """Envolver varios nodos OMML dentro de un <m:oMath>."""
    omath = etree.Element(_m("oMath"))
    for h in hijos:
        if h is not None:
            omath.append(h)
    return omath


def _encierra_grupo(omml_hijo):
    """Envolver un nodo OMML en <m:r><m:t>...</m:t></m:r> cuando es texto simple
    o en <m:r> si ya es un elemento. Para grupos usamos un oMath anidado."""
    if omml_hijo is None:
        return _texto_run("")
    if omml_hijo.tag == _m("oMath"):
        # Ya es un bloque completo: insertar directamente
        return omml_hijo
    return omml_hijo


def _base_sup_sub(tipo: str, base, extra):
    """Construir <m:sSup> o <m:sSub>."""
    base_omml = _ast_a_omml(base)
    extra_omml = _ast_a_omml(extra)
    elem = etree.Element(_m(tipo))
    e = etree.SubElement(elem, _m("e"))
    e.append(_encierra_grupo(base_omml))
    sup_sub = etree.SubElement(elem, _m("sup") if tipo == "sSup" else _m("sub"))
    sup_sub.append(_encierra_grupo(extra_omml))
    return elem


def _ast_a_omml_lista(nodo) -> list:
    """Convertir un nodo AST a una lista de elementos OMML (sin envolver en oMath).

    Cada elemento de la lista es un elemento OMML listo para ser insertado como
    hijo directo de un <m:oMath>.
    """
    if nodo is None:
        return [_texto_run("")]

    tipo = nodo[0]

    # Átomos simples
    if tipo == "texto":
        return [_texto_run(nodo[1])]
    if tipo == "espacio":
        return [_texto_run(" ")]
    if tipo == "espacio_fino":
        return [_texto_run("\u2009")]  # thin space U+2009
    if tipo == "numero":
        return [_texto_run(nodo[1])]
    if tipo == "variable":
        return [_texto_run(nodo[1])]
    if tipo == "operador":
        return [_texto_run(nodo[1])]
    if tipo == "group":
        return _ast_a_omml_lista(nodo[1])

    # Operación binaria (infix): izquierdo, operador, derecho como elementos separados
    if tipo == "infix":
        izq = _ast_a_omml_lista(nodo[2])
        der = _ast_a_omml_lista(nodo[3])
        op = [_texto_run(nodo[1])]
        return izq + op + der

    # Multiplicación por yuxtaposición: cada factor como elemento separado
    if tipo == "mul":
        hijos: list = []
        for f in nodo[1]:
            hijos.extend(_ast_a_omml_lista(f))
        return hijos

    # Potencia: un solo elemento sSup
    if tipo == "sup":
        return [_base_sup_sub("sSup", nodo[1], nodo[2])]

    # Subíndice: un solo elemento sSub
    if tipo == "sub":
        return [_base_sup_sub("sSub", nodo[1], nodo[2])]

    # Fracción: un solo elemento f
    if tipo == "fraccion":
        num = _ast_a_omml_lista(nodo[1])
        den = _ast_a_omml_lista(nodo[2])
        f = etree.Element(_m("f"))
        numel = etree.SubElement(f, _m("num"))
        for n in num:
            numel.append(_encierra_grupo(n))
        denel = etree.SubElement(f, _m("den"))
        for d in den:
            denel.append(_encierra_grupo(d))
        return [f]

    # Raíz cuadrada (o enésima con grado): un solo elemento rad
    if tipo == "raiz":
        operando = _ast_a_omml_lista(nodo[1])
        grado = nodo[2]
        rad = etree.Element(_m("rad"))
        if grado is not None:
            deg = etree.SubElement(rad, _m("deg"))
            for g in _ast_a_omml_lista(grado):
                deg.append(_encierra_grupo(g))
        e = etree.SubElement(rad, _m("e"))
        for op_el in operando:
            e.append(_encierra_grupo(op_el))
        return [rad]

    # Integral: oMath con func, límites e integrando como hijos directos
    if tipo == "integral":
        omath = etree.Element(_m("oMath"))
        func = etree.Element(_m("func"))
        func_name = etree.SubElement(func, _m("funcName"))
        func_name.append(_texto_run("∫"))
        omath.append(func)
        inferior = nodo[1]
        superior = nodo[2]
        if inferior is not None:
            lim = etree.Element(_m("lim"))
            e = etree.SubElement(lim, _m("e"))
            for el in _ast_a_omml_lista(inferior):
                e.append(_encierra_grupo(el))
            omath.append(lim)
        if superior is not None:
            lim = etree.Element(_m("lim"))
            e = etree.SubElement(lim, _m("e"))
            for el in _ast_a_omml_lista(superior):
                e.append(_encierra_grupo(el))
            omath.append(lim)
        integrando = nodo[3] if len(nodo) > 3 else None
        if integrando is not None:
            for el in _ast_a_omml_lista(integrando):
                omath.append(el)
        return [omath]

    # Sumatoria: oMath con func, límites e integrando como hijos directos
    if tipo == "suma":
        omath = etree.Element(_m("oMath"))
        func = etree.Element(_m("func"))
        func_name = etree.SubElement(func, _m("funcName"))
        func_name.append(_texto_run("∑"))
        omath.append(func)
        inferior = nodo[1]
        superior = nodo[2]
        if inferior is not None:
            lim = etree.Element(_m("lim"))
            e = etree.SubElement(lim, _m("e"))
            for el in _ast_a_omml_lista(inferior):
                e.append(_encierra_grupo(el))
            omath.append(lim)
        if superior is not None:
            lim = etree.Element(_m("lim"))
            e = etree.SubElement(lim, _m("e"))
            for el in _ast_a_omml_lista(superior):
                e.append(_encierra_grupo(el))
            omath.append(lim)
        integrando = nodo[3] if len(nodo) > 3 else None
        if integrando is not None:
            for el in _ast_a_omml_lista(integrando):
                omath.append(el)
        return [omath]

    # Límite: oMath con func, límites e cuerpo como hijos directos
    if tipo == "limite":
        omath = etree.Element(_m("oMath"))
        func = etree.Element(_m("func"))
        func_name = etree.SubElement(func, _m("funcName"))
        func_name.append(_texto_run("lim"))
        omath.append(func)
        inferior = nodo[1]
        superior = nodo[2]
        if inferior is not None:
            lim = etree.Element(_m("lim"))
            e = etree.SubElement(lim, _m("e"))
            for el in _ast_a_omml_lista(inferior):
                e.append(_encierra_grupo(el))
            omath.append(lim)
        if superior is not None:
            lim = etree.Element(_m("lim"))
            e = etree.SubElement(lim, _m("e"))
            for el in _ast_a_omml_lista(superior):
                e.append(_encierra_grupo(el))
            omath.append(lim)
        cuerpo = nodo[3] if len(nodo) > 3 else None
        if cuerpo is not None:
            for el in _ast_a_omml_lista(cuerpo):
                omath.append(el)
        return [omath]

    # Producto: oMath con func, límites e integrando como hijos directos
    if tipo == "producto":
        omath = etree.Element(_m("oMath"))
        func = etree.Element(_m("func"))
        func_name = etree.SubElement(func, _m("funcName"))
        func_name.append(_texto_run("∏"))
        omath.append(func)
        inferior = nodo[1]
        superior = nodo[2]
        if inferior is not None:
            lim = etree.Element(_m("lim"))
            e = etree.SubElement(lim, _m("e"))
            for el in _ast_a_omml_lista(inferior):
                e.append(_encierra_grupo(el))
            omath.append(lim)
        if superior is not None:
            lim = etree.Element(_m("lim"))
            e = etree.SubElement(lim, _m("e"))
            for el in _ast_a_omml_lista(superior):
                e.append(_encierra_grupo(el))
            omath.append(lim)
        integrando = nodo[3] if len(nodo) > 3 else None
        if integrando is not None:
            for el in _ast_a_omml_lista(integrando):
                omath.append(el)
        return [omath]

    # Función (sin, cos, etc.): oMath con func y argumento como hijos directos
    if tipo == "funcion":
        nombre = nodo[1]
        arg = nodo[2]
        omath = etree.Element(_m("oMath"))
        func = etree.Element(_m("func"))
        func_name = etree.SubElement(func, _m("funcName"))
        func_name.append(_texto_run(nombre))
        omath.append(func)
        if arg is not None:
            for el in _ast_a_omml_lista(arg):
                omath.append(_encierra_grupo(el))
        return [omath]

    # Matriz: un solo elemento matrix (o groupChr para pmatrix)
    if tipo in ("matrix", "pmatrix"):
        filas = nodo[1]
        matriz = etree.Element(_m("matrix"))
        for fila in filas:
            fila_el = etree.SubElement(matriz, _m("matrixRow"))
            for celda_ast in _ast_a_omml_lista(fila):
                celda = etree.SubElement(fila_el, _m("cell"))
                celda.append(_encierra_grupo(celda_ast))
        if tipo == "pmatrix":
            group = etree.Element(_m("groupChr"))
            group.append(matriz)
            return [group]
        return [matriz]

    raise LatexConversionError(f"Nodo AST no reconocido: {tipo!r}")


def _ast_a_omml(nodo):
    """Convertir un nodo AST a un elemento OMML (oMath o run)."""
    hijos = _ast_a_omml_lista(nodo)
    if len(hijos) == 1:
        return hijos[0]
    return _envuelve_omml(hijos)


def latex_a_omml(texto: str):
    """Convertir una expresión LaTeX a un elemento OMML <m:oMath>."""
    ast = latex_a_ast(texto)
    return _ast_a_omml(ast)


def latex_a_omml_parrafo(texto: str):
    """Convertir una expresión LaTeX a un <m:oMathPara> (modo display)."""
    omath = latex_a_omml(texto)
    omath_para = etree.Element(_m("oMathPara"))
    omath_para.append(omath)
    return omath_para
