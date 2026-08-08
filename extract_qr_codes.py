"""Extrait tous les codes QR/numéros de reçu des PDFs de reçus perdus."""
import fitz
import re
import json

pdf_files = [
    'attached_assets/Recus_SC_Lot1_1785509571493.pdf',
    'attached_assets/Recus_SC_Lot1_1785509814262.pdf',
]

# Les codes QR dans les reçus suivent le format ESCIALES-XXXXXXXXXX ou similaire
# Les numéros de reçu sont aussi présents en texte
all_codes = set()
all_numeros = set()

for pdf_path in pdf_files:
    doc = fitz.open(pdf_path)
    print(f"\n=== {pdf_path} : {doc.page_count} pages ===")
    for i, page in enumerate(doc):
        text = page.get_text()
        # Chercher les codes QR (format typique)
        codes = re.findall(r'ESCIALES-[A-Z0-9\-]+', text)
        # Chercher les numéros de reçu (format SC-XXXX-YYYY)
        numeros = re.findall(r'SC-\d{4}-\d+', text)
        # Aussi chercher tout code alphanumérique long
        autres = re.findall(r'[A-Z]{2,4}-\d{4,}-[A-Z0-9]+', text)
        
        for c in codes:
            all_codes.add(c.strip())
        for n in numeros:
            all_numeros.add(n.strip())
        for a in autres:
            all_codes.add(a.strip())
        
        if i < 3:
            print(f"  Page {i+1} extrait (premiers 500 chars):")
            print(text[:500])
            print("  ---")
    doc.close()

print(f"\n=== CODES QR TROUVÉS ({len(all_codes)}) ===")
for c in sorted(all_codes):
    print(c)

print(f"\n=== NUMÉROS REÇUS TROUVÉS ({len(all_numeros)}) ===")
for n in sorted(all_numeros):
    print(n)
