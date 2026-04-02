from pypdf import PdfReader

reader = PdfReader("Soren Energia 07_Censurado.pdf")
texto = ""

for page in reader.pages:
    t = page.extract_text()
    if t:
        texto += t

if texto.strip():
    print("📄 PDF digital (no escaneado)")
else:
    print("🖨️ PDF escaneado (imagen)")
