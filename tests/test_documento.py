"""Pruebas del generador de documentos.

Todas las pruebas escriben en directorios temporales: nunca tocan ``output/``.
"""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from src import documento


class DocumentoBaseTest(unittest.TestCase):
    """Base con un directorio temporal y un destino por defecto aislado."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.tmp = Path(self._tmp.name)

        parche = patch.object(documento, "OUTPUT_DIR", self.tmp)
        parche.start()
        self.addCleanup(parche.stop)

    def _crear(self, titulo: str = "Informe", parrafos=None, **kwargs):
        if parrafos is None:
            parrafos = ["Primer párrafo", "Segundo párrafo"]

        return documento.crear_documento(titulo, parrafos, **kwargs)


class GeneracionBasicaTest(DocumentoBaseTest):
    """Contenido, formato y pie de página."""

    def test_crea_archivo_y_devuelve_ruta(self) -> None:
        ruta = self._crear("Informe de Gestión")

        self.assertIsInstance(ruta, Path)
        self.assertTrue(ruta.exists())
        self.assertEqual(ruta.parent, self.tmp)
        self.assertEqual(ruta.name, "informe_de_gestión.docx")

    def test_titulo_y_parrafos(self) -> None:
        ruta = self._crear("Mi título", ["Uno", "Dos", "Tres"])
        documento_abierto = Document(ruta)

        textos = [(p.style.name, p.text) for p in documento_abierto.paragraphs]

        self.assertEqual(
            textos,
            [
                ("Heading 1", "Mi título"),
                ("Normal", "Uno"),
                ("Normal", "Dos"),
                ("Normal", "Tres"),
            ],
        )

    def test_formato_base(self) -> None:
        ruta = self._crear()
        abierto = Document(ruta)

        normal = abierto.styles["Normal"]
        self.assertEqual(normal.font.name, "Arial")
        self.assertEqual(normal.font.size.pt, 11)

        seccion = abierto.sections[0]
        self.assertAlmostEqual(seccion.top_margin.inches, 0.8, places=2)
        self.assertAlmostEqual(seccion.bottom_margin.inches, 0.8, places=2)
        self.assertAlmostEqual(seccion.left_margin.inches, 1.0, places=2)
        self.assertAlmostEqual(seccion.right_margin.inches, 1.0, places=2)

        self.assertEqual(
            seccion.footer.paragraphs[0].text,
            "Generado con Word-Architect",
        )

    def test_formato_del_titulo(self) -> None:
        ruta = self._crear("Título principal")
        abierto = Document(ruta)

        encabezado = abierto.paragraphs[0]
        self.assertEqual(encabezado.style.name, "Heading 1")
        self.assertEqual(encabezado.alignment, WD_ALIGN_PARAGRAPH.LEFT)

        run = encabezado.runs[0]
        self.assertEqual(run.font.name, "Arial")
        self.assertEqual(run.font.size.pt, 20)
        self.assertEqual(str(run.font.color.rgb), "1F4E79")

        estilo = abierto.styles["Heading 1"]
        self.assertEqual(estilo.font.name, "Arial")
        self.assertEqual(str(estilo.font.color.rgb), "1F4E79")

    def test_titulo_se_recorta(self) -> None:
        ruta = self._crear("   Con espacios   ")
        abierto = Document(ruta)

        self.assertEqual(abierto.paragraphs[0].text, "Con espacios")

    def test_parrafos_en_blanco_se_ignoran(self) -> None:
        ruta = self._crear("T", ["Uno", "", "   ", "Dos"])
        abierto = Document(ruta)

        self.assertEqual(
            [p.text for p in abierto.paragraphs],
            ["T", "Uno", "Dos"],
        )

    def test_acepta_cualquier_iterable(self) -> None:
        ruta = self._crear("T", ("Uno", "Dos"))
        self.assertTrue(ruta.exists())

        ruta_generador = self._crear(
            "Generador",
            (texto for texto in ["A", "B"]),
        )
        self.assertTrue(ruta_generador.exists())


class NoSobrescrituraTest(DocumentoBaseTest):
    """El generador nunca debe reemplazar un archivo existente."""

    def test_destino_por_defecto_no_se_sobrescribe(self) -> None:
        primera = self._crear("Repetido")
        segunda = self._crear("Repetido")
        tercera = self._crear("Repetido")

        self.assertEqual(primera.name, "repetido.docx")
        self.assertEqual(segunda.name, "repetido_2.docx")
        self.assertEqual(tercera.name, "repetido_3.docx")
        self.assertTrue(primera.exists())
        self.assertTrue(segunda.exists())
        self.assertTrue(tercera.exists())

    def test_ruta_explicita_no_se_sobrescribe(self) -> None:
        destino = self.tmp / "elegido.docx"
        destino.write_text("contenido previo", encoding="utf-8")

        ruta = self._crear("T", ruta_destino=destino)

        self.assertEqual(ruta.name, "elegido_2.docx")
        self.assertEqual(
            destino.read_text(encoding="utf-8"),
            "contenido previo",
        )

        # El archivo nuevo debe ser un .docx válido, no texto plano.
        self.assertTrue(Document(ruta).paragraphs)

    def test_extension_distinta_se_conserva(self) -> None:
        destino = self.tmp / "informe.v2"

        ruta = self._crear("T", ruta_destino=destino)

        self.assertEqual(ruta.name, "informe.v2.docx")
        self.assertTrue(ruta.exists())

    def test_extension_mayuscula_se_respeta(self) -> None:
        ruta = self._crear("T", ruta_destino=self.tmp / "MAYUS.DOCX")

        self.assertEqual(ruta.name, "MAYUS.DOCX")


class RutasTest(DocumentoBaseTest):
    """Creación de carpetas y rechazo de destinos inválidos."""

    def test_crea_directorio_padre_inexistente(self) -> None:
        destino = self.tmp / "sub" / "carpeta" / "doc.docx"

        ruta = self._crear("T", ruta_destino=destino)

        self.assertEqual(ruta, destino)
        self.assertTrue(destino.exists())

    def test_destino_es_directorio(self) -> None:
        carpeta = self.tmp / "carpeta"
        carpeta.mkdir()

        with self.assertRaises(IsADirectoryError):
            self._crear("T", ruta_destino=carpeta)

    def test_destino_con_punto_final(self) -> None:
        """Un nombre sin extensión recibe ``.docx``."""

        ruta = self._crear("T", ruta_destino=self.tmp / "sin_extension")

        self.assertEqual(ruta.name, "sin_extension.docx")


class ValidacionTest(DocumentoBaseTest):
    """Entradas inválidas deben fallar con errores claros."""

    def test_titulo_vacio(self) -> None:
        with self.assertRaises(ValueError) as contexto:
            self._crear("")

        self.assertEqual(
            str(contexto.exception),
            documento.MENSAJE_TITULO_VACIO,
        )

    def test_titulo_solo_espacios(self) -> None:
        with self.assertRaises(ValueError):
            self._crear("   ")

    def test_sin_parrafos(self) -> None:
        with self.assertRaises(ValueError) as contexto:
            self._crear("T", [])

        self.assertEqual(
            str(contexto.exception),
            documento.MENSAJE_SIN_PARRAFOS,
        )

    def test_todos_los_parrafos_en_blanco(self) -> None:
        with self.assertRaises(ValueError):
            self._crear("T", ["", "   "])

    def test_titulo_no_es_texto(self) -> None:
        with self.assertRaises(TypeError) as contexto:
            self._crear(123)  # type: ignore[arg-type]

        self.assertEqual(
            str(contexto.exception),
            documento.MENSAJE_TITULO_INVALIDO,
        )

    def test_parrafos_como_texto(self) -> None:
        with self.assertRaises(TypeError) as contexto:
            self._crear("T", "un solo texto")

        self.assertEqual(
            str(contexto.exception),
            documento.MENSAJE_PARRAFOS_INVALIDOS,
        )

    def test_parrafo_no_es_texto(self) -> None:
        with self.assertRaises(TypeError) as contexto:
            self._crear("T", ["bien", 42])  # type: ignore[list-item]

        self.assertEqual(
            str(contexto.exception),
            documento.MENSAJE_PARRAFO_INVALIDO,
        )

    def test_no_crea_archivos_si_falla_la_validacion(self) -> None:
        with self.assertRaises(ValueError):
            self._crear("")

        self.assertEqual(list(self.tmp.iterdir()), [])


class MarcasTest(DocumentoBaseTest):
    """Las marcas solo se interpretan cuando se activan explícitamente."""

    def test_marcas_desactivadas_por_defecto(self) -> None:
        ruta = self._crear("T", ["# Sección", "- Elemento", "1. Primero"])
        abierto = Document(ruta)

        self.assertEqual(
            [(p.style.name, p.text) for p in abierto.paragraphs],
            [
                ("Heading 1", "T"),
                ("Normal", "# Sección"),
                ("Normal", "- Elemento"),
                ("Normal", "1. Primero"),
            ],
        )

    def test_marcas_activadas(self) -> None:
        parrafos = [
            "# Sección 1",
            "## Sección 2",
            "### Sección 3",
            "- Viñeta",
            "* Otra viñeta",
            "1. Primero",
            "2) Segundo",
            "Texto normal",
        ]

        ruta = self._crear("T", parrafos, procesar_marcas=True)
        abierto = Document(ruta)

        self.assertEqual(
            [(p.style.name, p.text) for p in abierto.paragraphs],
            [
                ("Heading 1", "T"),
                ("Heading 1", "Sección 1"),
                ("Heading 2", "Sección 2"),
                ("Heading 3", "Sección 3"),
                ("List Bullet", "Viñeta"),
                ("List Bullet", "Otra viñeta"),
                ("List Number", "Primero"),
                ("List Number", "Segundo"),
                ("Normal", "Texto normal"),
            ],
        )

    def test_marcas_con_espacios_extra(self) -> None:
        ruta = self._crear(
            "T",
            ["  #   Con espacios  ", "  -   Elemento  "],
            procesar_marcas=True,
        )
        abierto = Document(ruta)

        self.assertEqual(abierto.paragraphs[1].text, "Con espacios")
        self.assertEqual(abierto.paragraphs[2].text, "Elemento")

    def test_marcas_sin_texto_se_ignoran(self) -> None:
        ruta = self._crear(
            "T",
            ["# ", "##", "###   ", "- ", "*", "Texto"],
            procesar_marcas=True,
        )
        abierto = Document(ruta)

        self.assertEqual(
            [(p.style.name, p.text) for p in abierto.paragraphs],
            [
                ("Heading 1", "T"),
                ("Normal", "Texto"),
            ],
        )

    def test_numero_sin_punto_no_es_lista(self) -> None:
        ruta = self._crear("T", ["2026 informe"], procesar_marcas=True)
        abierto = Document(ruta)

        self.assertEqual(abierto.paragraphs[1].style.name, "Normal")


class NombreArchivoTest(DocumentoBaseTest):
    """Saneado del título para usarlo como nombre de archivo."""

    def test_acentos_y_guiones(self) -> None:
        ruta = self._crear("Introducción a la Informática")

        self.assertEqual(
            ruta.name,
            "introducción_a_la_informática.docx",
        )

    def test_simbolos_eliminados(self) -> None:
        ruta = self._crear("Informe: gestión/2026")

        self.assertEqual(ruta.name, "informe_gestión2026.docx")

    def test_titulo_sin_caracteres_validos(self) -> None:
        ruta = self._crear("!!!")

        self.assertEqual(ruta.name, "documento.docx")


if __name__ == "__main__":
    unittest.main()
