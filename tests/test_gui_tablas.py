"""Pruebas de la inserción de tablas desde la interfaz gráfica."""

import tempfile
import unittest
from pathlib import Path

from docx import Document

from src.gui import (
    PLANTILLA_TABLA,
    _bloque_tabla,
    _convertir_a_bloques,
    _es_linea_tabla,
)
from src.documento import crear_documento_profesional


class TestConversionTablas(unittest.TestCase):
    """Verificar que el editor reconoce y valida las tablas."""

    def test_convierte_lineas_de_tabla_en_bloque(self):
        """Una cabecera y dos filas producen un único bloque table."""
        bloques = _convertir_a_bloques("| Nombre | Edad |\n| Ana | 34 |\n| Luis | 41 |")

        self.assertEqual(len(bloques), 1)
        self.assertEqual(bloques[0]["tipo"], "table")
        self.assertEqual(bloques[0]["encabezados"], ["Nombre", "Edad"])
        self.assertEqual(bloques[0]["filas"], [["Ana", "34"], ["Luis", "41"]])

    def test_plantilla_insertada_es_una_tabla_valida(self):
        """La plantilla del botón genera cabecera y dos filas de datos."""
        bloques = _convertir_a_bloques(PLANTILLA_TABLA)

        self.assertEqual(len(bloques), 1)
        self.assertEqual(bloques[0]["tipo"], "table")
        self.assertEqual(bloques[0]["encabezados"], ["Concepto", "Detalle"])
        self.assertEqual(bloques[0]["filas"], [["", ""], ["", ""]])

    def test_conserva_acentos_porcentajes_y_moneda(self):
        """Las celdas mantienen acentos, «%» y símbolos monetarios."""
        bloques = _convertir_a_bloques("| Concepto | Valor |\n| Gestión anual | 12,5 % |\n| Presupuesto | 1.250 € |")

        self.assertEqual(
            bloques[0]["filas"],
            [["Gestión anual", "12,5 %"], ["Presupuesto", "1.250 €"]],
        )

    def test_barra_aislada_no_es_tabla(self):
        """Una barra dentro de una frase sigue siendo un párrafo."""
        self.assertFalse(_es_linea_tabla("Entrada | salida"))
        bloques = _convertir_a_bloques("Entrada | salida")

        self.assertEqual(bloques, [{"tipo": "paragraph", "texto": "Entrada | salida"}])

    def test_una_sola_barra_no_es_tabla(self):
        """Una línea con una única barra vertical no se interpreta como tabla."""
        bloques = _convertir_a_bloques("| solo una barra")

        self.assertEqual(bloques, [{"tipo": "paragraph", "texto": "| solo una barra"}])

    def test_rechaza_filas_con_columnas_distintas(self):
        """Una fila con menos celdas que la cabecera lanza ValueError."""
        with self.assertRaises(ValueError) as contexto:
            _convertir_a_bloques("| A | B |\n| 1 | 2 | 3 |")

        self.assertIn("fila 2", str(contexto.exception))

    def test_rechaza_cabecera_vacia(self):
        """Una tabla sin ningún texto en la cabecera no es válida."""
        with self.assertRaises(ValueError):
            _convertir_a_bloques("|  |  |\n| 1 | 2 |")

    def test_rechaza_tabla_de_una_columna(self):
        """_bloque_tabla exige al menos dos columnas."""
        with self.assertRaises(ValueError):
            _bloque_tabla([["Única"]])

    def test_varias_tablas_separadas(self):
        """Dos grupos separados por texto generan dos bloques table."""
        contenido = "| A | B |\n| 1 | 2 |\n\nTexto entre tablas.\n\n| C | D |\n| 3 | 4 |"
        bloques = _convertir_a_bloques(contenido)

        self.assertEqual([bloque["tipo"] for bloque in bloques], ["table", "paragraph", "table"])

    def test_marcas_anteriores_siguen_funcionando(self):
        """Encabezados, listas, citas y párrafos conservan su conversión."""
        contenido = "\n".join(
            [
                "# Título",
                "## Subtítulo",
                "### Sección",
                "- Viñeta",
                "* Otra viñeta",
                "1. Primer elemento",
                "2) Segundo elemento",
                "> Cita",
                "Párrafo normal",
            ]
        )
        bloques = _convertir_a_bloques(contenido)

        self.assertEqual(
            [bloque["tipo"] for bloque in bloques],
            [
                "heading",
                "heading",
                "heading",
                "bullet",
                "bullet",
                "numbered",
                "numbered",
                "quote",
                "paragraph",
            ],
        )
        self.assertEqual(bloques[0]["nivel"], 1)
        self.assertEqual(bloques[1]["nivel"], 2)
        self.assertEqual(bloques[2]["nivel"], 3)


class TestDocumentoConTabla(unittest.TestCase):
    """Generar archivos DOCX reales en un directorio temporal."""

    def setUp(self):
        self.temporal = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporal.cleanup)
        self.directorio = Path(self.temporal.name)

    def test_genera_docx_con_tabla_nativa(self):
        """El bloque table produce una tabla real de Word."""
        ruta = crear_documento_profesional(
            "Informe con tabla",
            _convertir_a_bloques("| Área | Resultado |\n| Ventas | 95 % |\n| Soporte | 1.200 € |"),
            self.directorio / "informe.docx",
        )

        self.assertTrue(ruta.exists())
        documento = Document(ruta)
        self.assertEqual(len(documento.tables), 1)
        self.assertEqual(documento.tables[0].cell(0, 0).text, "Área")
        self.assertEqual(documento.tables[0].cell(0, 1).text, "Resultado")

    def test_dimensiones_de_la_tabla(self):
        """La tabla tiene una fila de cabecera y tres filas de datos."""
        ruta = crear_documento_profesional(
            "Dimensiones",
            _convertir_a_bloques(
                "| Ciudad | Habitantes |\n| Madrid | 3.300.000 |\n| Lima | 9.700.000 |\n| Bogotá | 7.400.000 |"
            ),
            self.directorio / "dimensiones.docx",
        )

        tabla = Document(ruta).tables[0]
        self.assertEqual(len(tabla.rows), 4)
        self.assertEqual(len(tabla.columns), 2)
        self.assertEqual(tabla.cell(3, 0).text, "Bogotá")
        self.assertEqual(tabla.cell(3, 1).text, "7.400.000")

    def test_cabecera_y_celdas_son_editables(self):
        """Se puede modificar una celda, guardar y volver a leer."""
        ruta = crear_documento_profesional(
            "Editable",
            _convertir_a_bloques("| Producto | Precio |\n| Mesa | 120 € |"),
            self.directorio / "editable.docx",
        )

        documento = Document(ruta)
        tabla = documento.tables[0]
        tabla.cell(0, 0).text = "Artículo"
        tabla.cell(1, 1).text = "150 €"
        ruta_editada = self.directorio / "editable_2.docx"
        documento.save(ruta_editada)

        tabla_editada = Document(ruta_editada).tables[0]
        self.assertEqual(tabla_editada.cell(0, 0).text, "Artículo")
        self.assertEqual(tabla_editada.cell(1, 1).text, "150 €")

    def test_error_de_tabla_no_crea_archivo(self):
        """Si la tabla es inválida, no se llega a guardar ningún DOCX."""
        ruta = self.directorio / "invalido.docx"
        contenido = "# Informe\n| A | B |\n| 1 | 2 | 3 |"

        with self.assertRaises(ValueError):
            bloques = _convertir_a_bloques(contenido)
            crear_documento_profesional("Informe", bloques, ruta)

        self.assertEqual(list(self.directorio.iterdir()), [])


if __name__ == "__main__":
    unittest.main()
