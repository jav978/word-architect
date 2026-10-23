# AGENTS.md — Word-Architect

Proyecto Python para generar documentos `.docx` con `python-docx`. Todo el código, los comentarios y los mensajes están en español: responde al usuario en español y mantén ese idioma.

## Estado real (verificado)

- La generación está centralizada en `src/documento.py`. `src/main.py` y `src/gui.py` son CLIs delgados que solo recogen datos, validan y delegan. No dupliques formato ni guardado: edita `documento.py`.
- `src/gui.py` **no usa Tkinter** ni abre ninguna ventana. A pesar del nombre, es un CLI que interpreta marcas de formato. No existe diálogo de guardado ni cancelación; la cancelación no se puede implementar ni probar hasta que haya una GUI real.
- Hay pruebas con `unittest` de la biblioteca estándar en `tests/` (`test_documento.py`, `test_cli.py`). No se instaló `pytest`.
- No hay git, README, `requirements.txt` ni `pyproject.toml`. `templates/` y `docs/` están vacíos.
- Hay dos entornos: `.venv` es el válido (python-docx 1.2.0, tkinter 8.6, pymupdf). `.venve` está incompleto (sin `docx`). Usa siempre `.venv/bin/python`.

## Comandos

- CLI básico (sin marcas): `.venv/bin/python src/main.py`
- CLI con marcas de formato: `.venv/bin/python src/gui.py` (no abre ninguna ventana)
- Comprobar sintaxis: `.venv/bin/python -m py_compile src/documento.py src/main.py src/gui.py`
- Pruebas: `.venv/bin/python -m unittest discover -s tests -v`
- Inspeccionar un `.docx`: `.venv/bin/python -c "from docx import Document; d=Document('output/x.docx'); print([(p.style.name, p.text) for p in d.paragraphs])"`

## Arquitectura

- `src/documento.py::crear_documento(titulo, parrafos, ruta_destino=None, *, procesar_marcas=False) -> Path` es la única API pública. Valida, aplica formato, guarda y devuelve la ruta final.
- `src/main.py::crear_documento(titulo, parrafos)` es un wrapper compatible de 2 argumentos y **no** interpreta marcas.
- `src/gui.py::crear_documento(titulo, parrafos, ruta_destino=None)` es un wrapper compatible de 3 argumentos y activa `procesar_marcas=True`.
- Los CLIs importan de forma relativa si hay `__package__` y de forma absoluta (`from documento import ...`) en ejecución directa, para soportar tanto `python src/main.py` como `import src.main`.
- Los mensajes de validación viven en constantes (`MENSAJE_*`) en `documento.py`. `ValueError` para contenido vacío y `TypeError` para tipos incorrectos.
- `_resolver_destino` solo crea `output/` cuando se usa el destino por defecto; con una ruta explícita crea únicamente su directorio padre.
- `PROJECT_ROOT` y `OUTPUT_DIR` se derivan de `__file__` con `pathlib`.

## Guardado

- **Nunca sobrescribe**: `_ruta_libre` añade `_2`, `_3`… también en rutas explícitas.
- Si la extensión no es `.docx`, se añade al nombre completo: `informe.v2` → `informe.v2.docx`; `.DOCX` se conserva.
- Un destino que es una carpeta lanza `IsADirectoryError`.
- Los nombres por defecto se sanean desde el título (`informe de gestión` → `informe_de_gestión.docx`); si no queda ningún carácter válido, se usa `documento.docx`.

## Formato

- Arial; estilo `Normal` a 11 pt; márgenes 0.8/0.8/1/1 pulgadas.
- H1 (`add_heading(level=1)`) Arial 20 pt, color RGB(31, 78, 121), alineado a la izquierda. H2/H3 usan Arial.
- Párrafos normales con `space_after = 8 pt`.
- Pie de página: «Generado con Word-Architect».
- Los párrafos en blanco se ignoran. Con `procesar_marcas=True`, las marcas solas (`#`, `##`, `###`, `-`, `*`) también se ignoran.
- Marcas admitidas solo en `gui.py`: `# ` (H1), `## ` (H2), `### ` (H3), `- `/`* ` (List Bullet) y `N. `/`N) ` (List Number).

## Trampas conocidas

- `output/` contiene documentos reales del usuario y locks `.~lock.*` de OnlyOffice/LibreOffice. No borres, muevas ni sobrescribas nada ahí; tampoco elimines los locks.
- `test_word.py` no es una prueba: escribe `test.docx` en la raíz. Las pruebas nuevas van en `tests/test_*.py` y deben escribir en `tempfile`, nunca en `output/` ni en la raíz.
- Las pruebas parchean `documento.OUTPUT_DIR`; si añades pruebas de guardado por defecto, hazlo igual para no tocar `output/`.
- `pdf_reader.py` usa pymupdf (`fitz`) y tiene un bug: imprime solo el texto de la última página (`page` queda fuera del bucle). No forma parte del generador Word.
- `setup_project.py` crea carpetas relativas al CWD. Ejecútalo solo desde la raíz del proyecto.
- No hay `__init__.py`: `src` funciona como paquete de espacio de nombres; `import src.main` desde la raíz es válido.
- La cancelación del diálogo de guardado sigue pendiente porque no existe GUI. No la simules ni la declares implementada.

## Convenciones

- Rutas con `pathlib`.
- Conserva nombres de funciones, firmas públicas y mensajes en español.
- No añadas dependencias sin justificarlo.
- No afirmes que una prueba pasó sin haberla ejecutado.

## Al cambiar código

1. Lee `src/documento.py` y el CLI afectado antes de editar.
2. Haz cambios pequeños y conserva las firmas y los mensajes existentes.
3. Ejecuta `py_compile` y `unittest`. Si tocas el guardado, comprueba ruta existente, extensión distinta de `.docx`, destino que es carpeta y ausencia de sobrescritura.
4. Si tocas las marcas, revisa también `tests/test_documento.py` y `tests/test_cli.py`.
