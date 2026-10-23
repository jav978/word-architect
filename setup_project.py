from pathlib import Path

# Carpetas principales del proyecto Word-Architect
carpetas = [
    "src",
    "output",
    "templates",
    "tests",
    "docs"
]

# Crear las carpetas del proyecto
for carpeta in carpetas:
    Path(carpeta).mkdir(parents=True, exist_ok=True)
    print(f"Carpeta creada o verificada: {carpeta}")

print("\n¡Estructura de Word-Architect preparada correctamente!")