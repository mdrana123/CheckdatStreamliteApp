
import os
from pathlib import Path
from pdf2image import convert_from_path

pdf_dir = r"C:\Users\admin\Checkdat_Stampel_AI\1_NyaRitningar"

out_dir = r"C:\Users\admin\Checkdat_Stampel_AI\2_output_images"

dpi = 300
fmt = "png"

# Ensure output directory exists
Path(out_dir).mkdir(parents=True, exist_ok=True)

# Find all PDFs
pdf_paths = [p for p in Path(pdf_dir).glob("**/*") if p.suffix.lower() == ".pdf"]

print(f"Found {len(pdf_paths)} PDFs")

for pdf in pdf_paths:
    pdf_name = pdf.stem
    print(f"Converting: {pdf}")

    try:
        pages = convert_from_path(
            str(pdf),
            dpi=dpi,
            fmt=fmt,
        )

        for i, page in enumerate(pages, start=1):
            out_path = Path(out_dir) / f"{pdf_name}_p{i:03d}.{fmt}"
            page.save(out_path, fmt.upper())

    except Exception:
        continue


# In[ ]:


from pathlib import Path
from PIL import Image
import re

# INPUT folder with your page images
in_dir  = Path(r"C:\Users\admin\Checkdat_Stampel_AI\2_output_images")
# SINGLE OUTPUT folder for all crops
out_dir = Path(r"C:\Users\admin\Checkdat_Stampel_AI\3_cropsImages_flat")
out_dir.mkdir(parents=True, exist_ok=True)

# Process only these extensions
exts = {".png"}  

# Crop region fractions
LEFT_FRACTION   = 0.85
TOP_FRACTION    = 0.665
RIGHT_FRACTION  = .985
BOTTOM_FRACTION = 0.97
# Helper to make a safe, unique flat filename from the relative path
def flat_name(rel_path: Path) -> str:
    # e.g. "PDF_A\PDF_A_p001.png" -> "PDF_A__PDF_A_p001.png"
    s = str(rel_path).replace("\\", "/")
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", s)  # keep it filesystem-safe
    return s

files = sorted([p for p in in_dir.rglob("*") if p.suffix.lower() in exts])
print(f"Found {len(files)} images")

for src in files:
    try:
        im = Image.open(src)
        W, H = im.size

        left   = int(W * LEFT_FRACTION)
        top    = int(H * TOP_FRACTION)
        right  = int(W * RIGHT_FRACTION)
        bottom = int(H * BOTTOM_FRACTION)

        # Clamp
        left   = max(0, min(left, W))
        right  = max(0, min(right, W))
        top    = max(0, min(top, H))
        bottom = max(0, min(bottom, H))

        if right <= left or bottom <= top:
            print(f"Skip (empty crop): {src}")
            continue

        crop = im.crop((left, top, right, bottom))

        # Flat filename derived from relative path
        rel = src.relative_to(in_dir)
        base = flat_name(rel.with_suffix(""))  # strip original suffix first
        dst = out_dir / f"{base}_crop.png"

        # If you prefer incremental suffix to avoid overwrite collisions:
        n = 1
        dst_candidate = dst
        while dst_candidate.exists():
            dst_candidate = out_dir / f"{base}_crop_{n:02d}.png"
            n += 1
        dst = dst_candidate

        crop.save(dst, "PNG")
        print("Saved:", dst)
    except Exception as e:
        print(f"Failed on {src}: {e}")

print("Done.")


# In[ ]:




