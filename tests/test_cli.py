"""Pruebas de los dos CLIs de Word-Architect.

Se parchea ``input`` y la salida estándar, y se aísla ``OUTPUT_DIR`` en un
directorio temporal. Ninguna prueba escribe en ``output/``.
"""

from __future__ import annotations

import io
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import patch

from docx import Document

from src import documento, gui, main


class CliBaseTest(unittest.TestCase):
    """Utilidades comunes para ejecutar los ``main()`` sin interacción real."""

    modulo = main

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.tmp = Path(self._tmp.name)

        parche = patch.object(documento, "OUTPUT_DIR", self.tmp)
        parche.start()
        self.addCleanup(parche.stop)

    def _ejecutar(self, respuestas: list[str]) -> str:
        salida = io.StringIO()

        with patch("builtins.input", side_effect=respuestas):
            with redirect_stdout(salida):
                self.modulo.main()

        return salida.getvalue()

    def _documentos(self) -> list[Path]:
        return sorted(self.tmp.glob("*.docx"))


class MainCliTest(CliBaseTest):
    """El CLI básico no interpreta marcas y conserva sus mensajes."""

    modulo = main

    def test_genera_documento(self) -> None:
        salida = self._ejecutar(["Mi informe", "Uno", "Dos", ""])

        archivos = self._documentos()
        self.assertEqual(len(archivos), 1)
        self.assertEqual(archivos[0].name, "mi_informe.docx")

        self.assertIn("¡Documento creado correctamente!", salida)
        self.assertIn("Archivo: mi_informe.docx", salida)
        self.assertIn(f"Ubicación: {archivos[0]}", salida)

        abierto = Document(archivos[0])
        self.assertEqual(
            [p.text for p in abierto.paragraphs],
            ["Mi informe", "Uno", "Dos"],
        )

    def test_no_interpreta_marcas(self) -> None:
        self._ejecutar(["T", "# Sección", "- Elemento", ""])

        abierto = Document(self._documentos()[0])

        self.assertEqual(
            [(p.style.name, p.text) for p in abierto.paragraphs],
            [
                ("Heading 1", "T"),
                ("Normal", "# Sección"),
                ("Normal", "- Elemento"),
            ],
        )

    def test_titulo_vacio(self) -> None:
        salida = self._ejecutar([""])

        self.assertIn("Error: el título no puede estar vacío.", salida)
        self.assertEqual(self._documentos(), [])

    def test_titulo_solo_espacios(self) -> None:
        salida = self._ejecutar(["   "])

        self.assertIn("Error: el título no puede estar vacío.", salida)
        self.assertEqual(self._documentos(), [])

    def test_sin_parrafos(self) -> None:
        salida = self._ejecutar(["T", ""])

        self.assertIn("Error: debes escribir al menos un párrafo.", salida)
        self.assertEqual(self._documentos(), [])

    def test_error_al_guardar(self) -> None:
        with patch.object(
            main,
            "crear_documento",
            side_effect=OSError("disco lleno"),
        ):
            salida = self._ejecutar(["T", "Uno", ""])

        self.assertIn(
            "No se pudo crear el documento: disco lleno",
            salida,
        )
        self.assertEqual(self._documentos(), [])


class GuiDocumentoTest(CliBaseTest):
    """Pruebas del generador utilizado por la interfaz gráfica.

    Estas pruebas no abren la ventana Tkinter.
    Comprueban directamente la creación de documentos.
    """

    def test_wrapper_acepta_ruta_destino(self) -> None:
        """Comprobar que se acepta una ruta personalizada."""

        destino = self.tmp / "sub" / "elegido"

        ruta = gui.crear_documento(
            "T",
            ["# Sección"],
            destino,
        )

        self.assertEqual(
            ruta,
            self.tmp / "sub" / "elegido.docx",
        )

        self.assertTrue(ruta.exists())

        documento = Document(ruta)

        self.assertEqual(
            documento.paragraphs[1].style.name,
            "Heading 1",
        )

    def test_wrapper_no_sobrescribe(self) -> None:
        """Comprobar que se conserva el archivo anterior."""

        primera = gui.crear_documento(
            "T",
            ["Uno"],
            self.tmp / "doc.docx",
        )

        segunda = gui.crear_documento(
            "T",
            ["Dos"],
            self.tmp / "doc.docx",
        )

        self.assertEqual(
            primera.name,
            "doc.docx",
        )

        self.assertEqual(
            segunda.name,
            "doc_2.docx",
        )

        documento_primero = Document(primera)
        documento_segundo = Document(segunda)

        textos_primero = [
            parrafo.text
            for parrafo in documento_primero.paragraphs
        ]

        textos_segundo = [
            parrafo.text
            for parrafo in documento_segundo.paragraphs
        ]

        self.assertIn("Uno", textos_primero)
        self.assertIn("Dos", textos_segundo)


if __name__ == "__main__":
    unittest.main(verbosity=2)
