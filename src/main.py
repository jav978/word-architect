"""CLI básico de Word-Architect.

Recoge el título y los párrafos por terminal y delega la generación en
``documento.crear_documento`` sin interpretar marcas de formato.
"""

from pathlib import Path
from typing import List

if __package__:
    from .documento import (
        OUTPUT_DIR,
        PROJECT_ROOT,
        crear_documento as _crear_documento,
    )
else:  # ejecución directa: python src/main.py
    from documento import (  # type: ignore[no-redef]
        OUTPUT_DIR,
        PROJECT_ROOT,
        crear_documento as _crear_documento,
    )


__all__ = ["crear_documento", "main", "PROJECT_ROOT", "OUTPUT_DIR"]


def crear_documento(titulo: str, parrafos: List[str]) -> Path:
    """Crear un documento Word con formato profesional."""

    return _crear_documento(titulo, parrafos)


# ==========================================
# PROGRAMA PRINCIPAL
# ==========================================

def main() -> None:
    """Solicitar información y generar el documento."""

    print("=" * 40)
    print("       WORD-ARCHITECT")
    print("   Generador de documentos Word")
    print("=" * 40)

    titulo = input("\nEscribe el título del documento: ").strip()

    if not titulo:
        print("Error: el título no puede estar vacío.")
        return

    print("\nEscribe el contenido del documento.")
    print("Pulsa Enter después de cada párrafo.")
    print("Cuando termines, pulsa Enter en una línea vacía.\n")

    parrafos = []

    while True:
        texto = input("> ")

        if not texto.strip():
            break

        parrafos.append(texto.strip())

    if not parrafos:
        print("Error: debes escribir al menos un párrafo.")
        return

    try:
        ruta = crear_documento(titulo, parrafos)

        print("\n¡Documento creado correctamente!")
        print(f"Archivo: {ruta.name}")
        print(f"Ubicación: {ruta}")

    except Exception as error:
        print(f"\nNo se pudo crear el documento: {error}")


# Ejecutar el programa
if __name__ == "__main__":
    main()
