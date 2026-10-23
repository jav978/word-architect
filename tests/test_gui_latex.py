"""Pruebas de soporte LaTeX → OMML (ecuaciones nativas de Word).

Verifica:
  * Fórmulas en línea $...$
  * Ecuaciones independientes $$...$$
  * Conservación del texto circundante
  * Generación de DOCX real con estructuras OMML nativas
  * Que las ecuaciones NO se convierten en imágenes
  * Manejo de expresiones inválidas
  * Compatibilidad con las 57 pruebas anteriores
"""

import tempfile
import unittest
import zipfile
from pathlib import Path

from docx import Document
from lxml import etree

from src.documento import crear_documento_profesional
from src.gui import _convertir_a_bloques
from src.ecuaciones import latex_a_omml, latex_a_omml_parrafo, LatexConversionError

OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _tiene_omml(ruta: Path) -> bool:
    """Detectar si el DOCX contiene elementos OMML (ecuaciones nativas)."""
    with zipfile.ZipFile(ruta) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    return "oMath" in xml


def _tiene_imagen(ruta: Path) -> bool:
    with zipfile.ZipFile(ruta) as z:
        nombres = z.namelist()
    return any(n.startswith("word/media/") for n in nombres)


def _texto_func_name(func) -> str | None:
    """Obtener el texto del <m:funcName> (está dentro de <m:r><m:t>)."""
    func_name = func.find(f"{{{OMML_NS}}}funcName")
    if func_name is None:
        return None
    t = func_name.find(f"{{{OMML_NS}}}r").find(f"{{{OMML_NS}}}t")
    return t.text if t is not None else None


class TestConversionLatexOmml(unittest.TestCase):
    """Conversión de expresiones LaTeX a elementos OMML."""

    def test_en_linea_simple(self):
        omml = latex_a_omml("f(x)=x^2+3x+2")
        self.assertIsNotNone(omml)
        self.assertEqual(omml.tag, f"{{{OMML_NS}}}oMath")

    def test_fraccion(self):
        omml = latex_a_omml(r"\frac{a+b}{c}")
        self.assertEqual(omml.tag, f"{{{OMML_NS}}}oMath")
        hijo = omml[0]
        self.assertEqual(etree.QName(hijo).localname, "f")

    def test_potencia(self):
        omml = latex_a_omml("x^2")
        self.assertEqual(omml.tag, f"{{{OMML_NS}}}oMath")
        hijo = omml[0]
        self.assertEqual(etree.QName(hijo).localname, "sSup")

    def test_subindice(self):
        omml = latex_a_omml("a_i")
        self.assertEqual(omml.tag, f"{{{OMML_NS}}}oMath")
        hijo = omml[0]
        self.assertEqual(etree.QName(hijo).localname, "sSub")

    def test_raiz(self):
        omml = latex_a_omml(r"\sqrt{x}")
        self.assertEqual(omml.tag, f"{{{OMML_NS}}}oMath")
        hijo = omml[0]
        self.assertEqual(etree.QName(hijo).localname, "rad")

    def test_integral(self):
        omml = latex_a_omml(r"\int_0^1 x^2\,dx")
        self.assertEqual(omml.tag, f"{{{OMML_NS}}}oMath")
        func = omml[0]
        self.assertEqual(etree.QName(func).localname, "func")
        self.assertEqual(_texto_func_name(func), "∫")
        limites = func.findall(f"{{{OMML_NS}}}lim")
        self.assertEqual(len(limites), 2)
        self.assertEqual(limites[0].get(f"{{{OMML_NS}}}val"), "inf")
        self.assertEqual(limites[1].get(f"{{{OMML_NS}}}val"), "sup")
        arg = func.find(f"{{{OMML_NS}}}arg")
        self.assertIsNotNone(arg)

    def test_sumatoria(self):
        omml = latex_a_omml(r"\sum_{i=1}^{n} i")
        self.assertEqual(omml.tag, f"{{{OMML_NS}}}oMath")
        func = omml[0]
        self.assertEqual(etree.QName(func).localname, "func")
        self.assertEqual(_texto_func_name(func), "∑")
        limites = func.findall(f"{{{OMML_NS}}}lim")
        self.assertEqual(len(limites), 2)
        self.assertEqual(limites[0].get(f"{{{OMML_NS}}}val"), "inf")
        self.assertEqual(limites[1].get(f"{{{OMML_NS}}}val"), "sup")
        self.assertIsNotNone(func.find(f"{{{OMML_NS}}}arg"))

    def test_limite(self):
        omml = latex_a_omml(r"\lim_{x \to \infty} f(x)")
        self.assertEqual(omml.tag, f"{{{OMML_NS}}}oMath")
        func = omml[0]
        self.assertEqual(etree.QName(func).localname, "func")
        self.assertEqual(_texto_func_name(func), "lim")
        self.assertIsNotNone(func.find(f"{{{OMML_NS}}}arg"))

    def test_omml_parrafo_display(self):
        omml_para = latex_a_omml_parrafo(r"\int_0^1 x^2\,dx")
        self.assertEqual(omml_para.tag, f"{{{OMML_NS}}}oMathPara")

    def test_expresion_invalida_lanza_error(self):
        with self.assertRaises(LatexConversionError):
            latex_a_omml(r"\comando_inexistente{x}")


class TestBloquesDesdeEditor(unittest.TestCase):
    """Conversión del contenido del editor a bloques con ecuaciones."""

    def test_ecuacion_independiente(self):
        bloques = _convertir_a_bloques("$$\\int_0^1 x^2\\,dx=\\frac{1}{3}$$")
        self.assertEqual(len(bloques), 1)
        self.assertEqual(bloques[0]["tipo"], "equation")
        self.assertEqual(bloques[0]["modo"], "display")
        self.assertEqual(
            bloques[0]["latex"], r"\int_0^1 x^2\,dx=\frac{1}{3}"
        )

    def test_ecuacion_en_linea(self):
        bloques = _convertir_a_bloques("La función es $f(x)=x^2+3x+2$.")
        self.assertEqual(len(bloques), 1)
        self.assertEqual(bloques[0]["tipo"], "paragraph")
        partes = bloques[0]["partes"]
        self.assertEqual(len(partes), 3)
        self.assertEqual(partes[0], {"tipo": "texto", "contenido": "La función es "})
        self.assertEqual(
            partes[1],
            {"tipo": "ecuacion", "modo": "inline", "latex": "f(x)=x^2+3x+2"},
        )
        self.assertEqual(partes[2], {"tipo": "texto", "contenido": "."})

    def test_varias_ecuaciones_en_linea(self):
        bloques = _convertir_a_bloques("$a$ y $b$ son variables.")
        partes = bloques[0]["partes"]
        self.assertEqual(len(partes), 4)
        self.assertEqual(partes[0], {"tipo": "ecuacion", "modo": "inline", "latex": "a"})
        self.assertEqual(partes[1], {"tipo": "texto", "contenido": " y "})
        self.assertEqual(partes[2], {"tipo": "ecuacion", "modo": "inline", "latex": "b"})
        self.assertEqual(partes[3], {"tipo": "texto", "contenido": " son variables."})

    def test_texto_tabla_ecuacion_conserva_orden(self):
        bloques = _convertir_a_bloques(
            "Introducción\n"
            "$$\sum_{i=1}^{n} i = \\frac{n(n+1)}{2}$$\n"
            "| Columna | Valor |\n"
            "| A | 1 |\n"
            "Conclusión"
        )
        tipos = [b["tipo"] for b in bloques]
        self.assertEqual(tipos, ["paragraph", "equation", "table", "paragraph"])

    def test_barra_sola_no_es_ecuacion(self):
        """Una línea con una sola $ sin cierre no se trata como ecuación."""
        bloques = _convertir_a_bloques("Precio $ 100")
        self.assertEqual(bloques[0]["tipo"], "paragraph")
        self.assertEqual(bloques[0]["texto"], "Precio $ 100")

    def test_ecuacion_vacia_display_no_valida(self):
        """Una línea $$...$$ vacía no genera bloque equation válido."""
        bloques = _convertir_a_bloques("$$$$")
        self.assertEqual(len(bloques), 1)
        self.assertEqual(bloques[0]["tipo"], "paragraph")
        self.assertEqual(bloques[0]["texto"], "$$$$")


class TestDocumentoConEcuaciones(unittest.TestCase):
    """Generación de DOCX reales con ecuaciones OMML."""

    def setUp(self):
        self.temporal = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporal.cleanup)
        self.directorio = Path(self.temporal.name)

    def test_genera_docx_con_ecuacion_display(self):
        """El bloque equation produce OMML nativo en el DOCX."""
        ruta = crear_documento_profesional(
            "Ecuación",
            _convertir_a_bloques("$$\\int_0^1 x^2\\,dx=\\frac{1}{3}$$"),
            self.directorio / "ecuacion.docx",
        )
        self.assertTrue(ruta.exists())
        self.assertTrue(_tiene_omml(ruta))
        self.assertFalse(_tiene_imagen(ruta))

        documento = Document(ruta)
        self.assertGreaterEqual(len(documento.paragraphs), 2)  # título + ecuación

    def test_genera_docx_con_ecuacion_en_linea(self):
        """La ecuación en línea se incrusta dentro del párrafo."""
        ruta = crear_documento_profesional(
            "Inline",
            _convertir_a_bloques("La función es $f(x)=x^2+3x+2$."),
            self.directorio / "inline.docx",
        )
        self.assertTrue(_tiene_omml(ruta))
        self.assertFalse(_tiene_imagen(ruta))

        documento = Document(ruta)
        # Al menos un párrafo debe contener OMML en su XML
        omml_en_parrafos = any(
            OMML_NS in p._p.xml for p in documento.paragraphs
        )
        self.assertTrue(omml_en_parrafos)

    def test_conserva_texto_alrededor_ecuacion(self):
        """El texto antes y después de la ecuación en línea se conserva."""
        ruta = crear_documento_profesional(
            "Contexto",
            _convertir_a_bloques("Inicio $x^2$ fin."),
            self.directorio / "contexto.docx",
        )
        documento = Document(ruta)
        texto_completo = "\n".join(p.text for p in documento.paragraphs)
        self.assertIn("Inicio", texto_completo)
        self.assertIn("fin", texto_completo)

    def test_ecuacion_entre_texto_y_tabla(self):
        """El orden texto → ecuación → tabla se preserva en el DOCX."""
        ruta = crear_documento_profesional(
            "Orden",
            _convertir_a_bloques(
                "Primera sección\n"
                "$$\sum_{i=1}^{n} i$$\n"
                "| Col | Val |\n"
                "| A | 1 |\n"
                "Última sección"
            ),
            self.directorio / "orden.docx",
        )
        documento = Document(ruta)
        self.assertEqual(len(documento.tables), 1)
        self.assertTrue(_tiene_omml(ruta))
        self.assertFalse(_tiene_imagen(ruta))

    def test_expresion_invalida_no_genera_archivo(self):
        """Una expresión LaTeX inválida impide la generación del DOCX."""
        ruta = self.directorio / "invalido.docx"
        # \comando_inexistente no es un comando soportado
        bloques = [{"tipo": "equation", "modo": "display", "latex": r"\comando_inexistente{x}"}]
        with self.assertRaises(ValueError):
            crear_documento_profesional("Inválido", bloques, ruta)
        self.assertEqual(list(self.directorio.iterdir()), [])

    def test_ecuacion_sin_latex_no_genera_archivo(self):
        """Un bloque equation sin campo 'latex' lanza ValueError."""
        ruta = self.directorio / "vacio.docx"
        bloques = [{"tipo": "equation", "modo": "display"}]
        with self.assertRaises(ValueError):
            crear_documento_profesional("Vacío", bloques, ruta)
        self.assertEqual(list(self.directorio.iterdir()), [])


class TestCompatibilidadAnterior(unittest.TestCase):
    """Verificar que las funcionalidades previas (tablas, imágenes, texto) siguen funcionando."""

    def setUp(self):
        self.temporal = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporal.cleanup)
        self.directorio = Path(self.temporal.name)

    def test_tabla_sigue_funcionando(self):
        ruta = crear_documento_profesional(
            "Tabla",
            _convertir_a_bloques("| A | B |\n| 1 | 2 |"),
            self.directorio / "tabla.docx",
        )
        documento = Document(ruta)
        self.assertEqual(len(documento.tables), 1)
        self.assertFalse(_tiene_omml(ruta))

    def test_texto_plano_sigue_funcionando(self):
        ruta = crear_documento_profesional(
            "Texto",
            _convertir_a_bloques("Párrafo simple sin fórmulas."),
            self.directorio / "texto.docx",
        )
        documento = Document(ruta)
        self.assertEqual(documento.paragraphs[1].text, "Párrafo simple sin fórmulas.")
        self.assertFalse(_tiene_omml(ruta))


class TestInspeccionXmlOmml(unittest.TestCase):
    """Inspección detallada del XML OMML para los casos obligatorios.

    Verifica que NO se pierde contenido: límites, integrando, operadores,
    subíndices, superíndices y el resto de la expresión se conservan en su
    orden correcto dentro de estructuras OMML válidas.
    """

    def _localname(self, elemento):
        return etree.QName(elemento).localname

    def _texto_runs(self, elemento):
        """Concatenar el texto de todos los runs <m:t> dentro de un elemento."""
        return "".join(
            t.text or ""
            for t in elemento.iter(f"{{{OMML_NS}}}t")
        )

    def test_integral_definida_conserva_todo_el_contenido(self):
        """$$\\int_0^1 x^2\\,dx=\\frac{1}{3}$$"""
        omml = latex_a_omml(r"\int_0^1 x^2\,dx=\frac{1}{3}")
        self.assertEqual(omml.tag, f"{{{OMML_NS}}}oMath")

        # El operador integral con sus límites y su integrando.
        func = omml[0]
        self.assertEqual(self._localname(func), "func")
        self.assertEqual(_texto_func_name(func), "∫")

        limites = func.findall(f"{{{OMML_NS}}}lim")
        self.assertEqual(len(limites), 2)
        self.assertEqual(limites[0].get(f"{{{OMML_NS}}}val"), "inf")
        self.assertEqual(self._texto_runs(limites[0]), "0")
        self.assertEqual(limites[1].get(f"{{{OMML_NS}}}val"), "sup")
        self.assertEqual(self._texto_runs(limites[1]), "1")

        arg = func.find(f"{{{OMML_NS}}}arg")
        self.assertIsNotNone(arg)
        arg_texto = self._texto_runs(arg)
        self.assertIn("x", arg_texto)
        self.assertIn("2", arg_texto)  # superíndice de x
        self.assertIn("dx", arg_texto)  # diferencial

        # La expresión después de la integral (= 1/3) sigue presente en el oMath.
        despues = omml[1:]
        self.assertGreater(len(despues), 0, "La expresión tras la integral no se perdió")

    def test_integral_indefinida_conserva_integrando_y_diferencial(self):
        """$$\\int x^2\\,dx=\\frac{x^3}{3}+C$$"""
        omml = latex_a_omml(r"\int x^2\,dx=\frac{x^3}{3}+C")
        self.assertEqual(omml.tag, f"{{{OMML_NS}}}oMath")
        func = omml[0]
        self.assertEqual(self._localname(func), "func")
        self.assertEqual(_texto_func_name(func), "∫")

        # Sin límites en una integral indefinida.
        self.assertEqual(len(func.findall(f"{{{OMML_NS}}}lim")), 0)

        arg = func.find(f"{{{OMML_NS}}}arg")
        self.assertIsNotNone(arg)
        arg_texto = self._texto_runs(arg)
        self.assertIn("x", arg_texto)
        self.assertIn("dx", arg_texto)

    def test_fraccion_conserva_numerador_y_denominador(self):
        """$$\\frac{a+b}{c}$$"""
        omml = latex_a_omml(r"\frac{a+b}{c}")
        self.assertEqual(omml.tag, f"{{{OMML_NS}}}oMath")
        f = omml[0]
        self.assertEqual(self._localname(f), "f")
        num = f.find(f"{{{OMML_NS}}}num")
        den = f.find(f"{{{OMML_NS}}}den")
        self.assertIsNotNone(num)
        self.assertIsNotNone(den)
        self.assertEqual(self._texto_runs(num), "a+b")
        self.assertEqual(self._texto_runs(den), "c")

    def test_potencia_conserva_base_y_exponente(self):
        """$$x^2+y^3$$"""
        omml = latex_a_omml("x^2+y^3")
        self.assertEqual(omml.tag, f"{{{OMML_NS}}}oMath")
        self.assertEqual(self._texto_runs(omml), "x2+y3")
        supers = omml.findall(f".//{{{OMML_NS}}}sup")
        self.assertGreaterEqual(len(supers), 2)

    def test_raiz_conserva_radicando(self):
        """$$\\sqrt{x+1}$$"""
        omml = latex_a_omml(r"\sqrt{x+1}")
        self.assertEqual(omml.tag, f"{{{OMML_NS}}}oMath")
        rad = omml[0]
        self.assertEqual(self._localname(rad), "rad")
        self.assertEqual(self._texto_runs(rad), "x+1")

    def test_sumatoria_conserva_limites_y_sumando(self):
        """$$\\sum_{i=1}^{n} i=\\frac{n(n+1)}{2}$$"""
        omml = latex_a_omml(r"\sum_{i=1}^{n} i")
        self.assertEqual(omml.tag, f"{{{OMML_NS}}}oMath")
        func = omml[0]
        self.assertEqual(self._localname(func), "func")
        self.assertEqual(_texto_func_name(func), "∑")

        limites = func.findall(f"{{{OMML_NS}}}lim")
        self.assertEqual(len(limites), 2)
        self.assertEqual(limites[0].get(f"{{{OMML_NS}}}val"), "inf")
        self.assertEqual(self._texto_runs(limites[0]), "i=1")
        self.assertEqual(limites[1].get(f"{{{OMML_NS}}}val"), "sup")
        self.assertEqual(self._texto_runs(limites[1]), "n")

        arg = func.find(f"{{{OMML_NS}}}arg")
        self.assertIsNotNone(arg)
        self.assertEqual(self._texto_runs(arg), "i")


if __name__ == "__main__":
    unittest.main()
