import fitz  # pymupdf

ruta = "ilovepdf_merged (1).pdf"
doc = fitz.open(ruta)

for i in range(doc.page_count):
    page = doc[i]
    imgs = page.get_images(full=True)
    widgets = list(page.widgets() or [])
    print(f"Página {i+1}: imágenes={len(imgs)}, widgets={len(widgets)}")

