import fitz

doc = fitz.open("733 TSP 1 2026-2.pdf")
for page in doc: print("Página", page.number+1)
print(page.get_text())

doc.close()
