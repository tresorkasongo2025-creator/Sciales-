from pathlib import Path
import fitz

pdf_path = Path("attached_assets/Recus_SCA_Lot2_S2_1785552702779.pdf")
out_dir = Path(".agents/outputs/recus_pdf_inspection")
out_dir.mkdir(parents=True, exist_ok=True)

doc = fitz.open(pdf_path)
print("pages", doc.page_count)
print("metadata", doc.metadata)
for index, page in enumerate(doc):
    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
    image_path = out_dir / f"page_{index + 1:03d}.png"
    pix.save(image_path)
    text = page.get_text("text")
    print(f"--- page {index + 1} ---")
    print(text[:5000])