# extract_raw.py
import os
from pathlib import Path
import re
import numpy as np
import pandas as pd
from PIL import Image
import easyocr
from ultralytics import YOLO
import cv2

model_path = r"C:\Users\admin\Checkdat_Stampel_AI\runs\detect\train\weights\best.pt"
image_dir = Path(r"C:\Users\admin\Checkdat_Stampel_AI\3_cropsImages_flat")
raw_excel_out = r"C:\Users\admin\Checkdat_Stampel_AI\Data_1_RawData.xlsx"

# ===== Debug =====
SAVE_DEBUG_CROPS = False
DEBUG_DIR = r"C:\Users\admin\Checkdat_Stampel_AI\debug_crops"

reader = easyocr.Reader(["sv", "en"], gpu=False)
model = YOLO(model_path)
names = model.names

LABELS = [
    "ANLÄGGNINGDEL",
    "AVDELNING",
    "BESTÄLLARE",
    "BET",
    "BET_REVIDERING",
    "BLAD",
    "Beskrivning_1",
    "Beskrivning_2",
    "Beskrivning_3",
    "Beskrivning_4",
    "DATUM",
    "DELOMRÅDE/BANDEL",
    "FORMAT",
    "FÖRVALTNINGSNUMMER",
    "GODKÄND AV",
    "GRANSKNINGSSTATUS/SYFTE",
    "HANDLINGSTYP",
    "KOMMUN",
    "KONSTRUKTIONSNUMMER",
    "LEVERANS/ÄNDRINGS-PM",
    "LEVERANTÖR",
    "NÄSTA BLAD",
    "OBJEKT",
    "OBJEKTNUMMER/KM",
    "RITNINGSNUMMER",
    "RITNINGSTYP",
    "SKALA",
    "SKAPAD AV",
    "TEKNIKOMRÅDE",
    "UPPDRAGSNUMMER",
]

if SAVE_DEBUG_CROPS:
    os.makedirs(DEBUG_DIR, exist_ok=True)


def list_images(folder: Path):
    exts = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"}
    return [p for p in folder.iterdir() if p.suffix.lower() in exts]


def normalize_text_1line(text: str) -> str:
    """Normalize to a single clean line."""
    if not text:
        return ""
    text = text.replace("\r\n", "\n").replace("\r", "\n").strip()
    # Keep only the first line if any newlines exist
    text = text.split("\n", 1)[0]
    text = re.sub(r"\s+", " ", text).strip()
    return text


def safe_for_filename(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|\-]+', "_", s)


def _easyocr_first_line(arr: np.ndarray) -> str:
    """
    Read OCR and keep ONLY one line (top-most line).
    This avoids EasyOCR "leaking" next label text when boxes are large.
    """
    # detail=1 gives bounding boxes for sorting
    det = reader.readtext(arr, detail=1, paragraph=False)
    if not det:
        return ""

    # det: [ (bbox, text, conf), ... ]
    items = []
    for bbox, text, conf in det:
        if not text:
            continue
        # bbox is 4 points; take top-left y,x
        xs = [p[0] for p in bbox]
        ys = [p[1] for p in bbox]
        x_min, y_min = float(min(xs)), float(min(ys))
        items.append((y_min, x_min, str(text)))

    if not items:
        return ""

    # sort by y then x
    items.sort(key=lambda t: (t[0], t[1]))

    # group into the first line using a y-threshold
    first_y = items[0][0]
    # tolerance based on image height
    H = arr.shape[0]
    y_tol = max(8.0, 0.03 * H)

    first_line_texts = [t for (y, x, t) in items if abs(y - first_y) <= y_tol]

    # join and normalize to one line
    return normalize_text_1line(" ".join(first_line_texts))


def ocr_crop_with_easyocr_one_line(pil_img: Image.Image) -> str:
    w, h = pil_img.size
    if max(w, h) < 300:
        pil_img = pil_img.resize((w * 3, h * 3), Image.BICUBIC)

    arr = np.array(pil_img)
    return _easyocr_first_line(arr)


def clean_beskrivning3_text(text: str) -> str:
    """
    Beskrivning_3:
      - If OCR detects nothing -> keep empty.
      - If it looks like it is actually Beskrivning_4 -> keep empty.
      - Always keep ONLY one line (already handled by OCR function).
    """
    t = normalize_text_1line(text)
    if not t:
        return ""

    up = t.upper()

    # Common "leak" indicators: header text or a leading "4"
    if re.search(r"\bBESKRIVNING\s*[_]?\s*4\b", up):
        return ""
    if re.match(r"^\s*4\b", t):
        return ""

    return t


def adjust_bbox_for_label(label_name: str, x1: int, y1: int, x2: int, y2: int):
    """
    Fix for Beskrivning_3 leaking into Beskrivning_4:
    - Make Beskrivning_3 crop only the TOP part of its box.
    This reduces the chance that the crop includes Beskrivning_4 text.
    """
    if label_name == "Beskrivning_3":
        h = max(0, y2 - y1)
        # keep top 60% of the box
        y2 = y1 + int(h * 0.60)
    return x1, y1, x2, y2


def detect_andring_line_symbol(crop_img: Image.Image, debug_save_basepath: str | None = None) -> str | None:
    img_bgr = cv2.cvtColor(np.array(crop_img), cv2.COLOR_RGB2BGR)
    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)

    h0, w0 = gray.shape[:2]
    scale = 1
    if max(w0, h0) < 220:
        scale = 5
    elif max(w0, h0) < 350:
        scale = 4
    elif max(w0, h0) < 500:
        scale = 3
    if scale > 1:
        gray = cv2.resize(gray, (w0 * scale, h0 * scale), interpolation=cv2.INTER_CUBIC)

    gray_blur = cv2.GaussianBlur(gray, (3, 3), 0)

    _, bw_otsu = cv2.threshold(gray_blur, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    close_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    bw_otsu = cv2.morphologyEx(bw_otsu, cv2.MORPH_CLOSE, close_kernel, iterations=1)

    bw_adapt = cv2.adaptiveThreshold(
        gray_blur, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY_INV,
        31,
        5
    )
    bw_adapt = cv2.morphologyEx(bw_adapt, cv2.MORPH_CLOSE, close_kernel, iterations=1)

    for idx, bw in enumerate([bw_otsu, bw_adapt]):
        H, W = bw.shape[:2]

        vert_len = max(15, int(H * 0.35))
        vert_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, vert_len))
        vertical = cv2.morphologyEx(bw, cv2.MORPH_OPEN, vert_kernel, iterations=1)
        bw_no_vert = cv2.bitwise_and(bw, cv2.bitwise_not(vertical))

        horiz_len = max(18, int(W * 0.18))
        horiz_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (horiz_len, 1))
        horizontal = cv2.morphologyEx(bw_no_vert, cv2.MORPH_OPEN, horiz_kernel, iterations=1)

        if debug_save_basepath:
            cv2.imwrite(f"{debug_save_basepath}_bw{idx}.png", bw)
            cv2.imwrite(f"{debug_save_basepath}_novert{idx}.png", bw_no_vert)
            cv2.imwrite(f"{debug_save_basepath}_horiz{idx}.png", horizontal)

        contours, _ = cv2.findContours(horizontal, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        for cnt in contours:
            x, y, w, h = cv2.boundingRect(cnt)
            aspect = w / max(h, 1)
            if (w >= 0.12 * W) and (aspect >= 5) and (h <= 0.12 * H):
                return "_"

        edges = cv2.Canny(gray_blur, 50, 150)
        lines = cv2.HoughLinesP(
            edges,
            rho=1,
            theta=np.pi / 180,
            threshold=25,
            minLineLength=int(0.12 * W),
            maxLineGap=8
        )

        if debug_save_basepath:
            cv2.imwrite(f"{debug_save_basepath}_edges{idx}.png", edges)

        if lines is not None:
            for x1, y1, x2, y2 in lines[:, 0]:
                dx = abs(x2 - x1)
                dy = abs(y2 - y1)
                if dx >= int(0.12 * W) and dy <= max(2, int(0.02 * H)):
                    return "_"

    return None


def process_folder():
    data_rows = []
    image_paths = list_images(image_dir)

    for img_path in image_paths:
        pil_img = Image.open(img_path).convert("RGB")
        w_img, h_img = pil_img.size

        row = {label: "" for label in LABELS}
        row["Image"] = img_path.name

        results = model(str(img_path))
        result = results[0]

        if result.boxes is None or len(result.boxes) == 0:
            data_rows.append(row)
            continue

        for i, box in enumerate(result.boxes):
            cls_id = int(box.cls[0])
            label_name = names.get(cls_id, str(cls_id))

            if label_name not in LABELS:
                continue

            x1, y1, x2, y2 = box.xyxy[0].cpu().numpy().astype(int)
            x1, y1 = max(0, x1), max(0, y1)
            x2, y2 = min(w_img, x2), min(h_img, y2)

            # ✅ Label-specific crop adjustment (prevents Beskrivning_3 leaking into Beskrivning_4)
            x1, y1, x2, y2 = adjust_bbox_for_label(label_name, x1, y1, x2, y2)

            crop = pil_img.crop((x1, y1, x2, y2))

            debug_base = None
            if SAVE_DEBUG_CROPS:
                safe_stem = safe_for_filename(img_path.stem)
                safe_label = safe_for_filename(label_name)
                debug_base = os.path.join(DEBUG_DIR, f"{safe_stem}_{safe_label}_{i}")
                crop.save(f"{debug_base}_crop.png")

            # ✅ OCR: ALWAYS one line only
            text = ocr_crop_with_easyocr_one_line(crop)

            # ✅ Beskrivning_3: if OCR didn't detect real text, keep EMPTY (no leakage)
            if label_name == "Beskrivning_3":
                text = clean_beskrivning3_text(text)

            # Underscore fallback for BET and BET_REVIDERING
            if label_name in ("BET", "BET_REVIDERING"):
                clean = (text or "").strip()
                if (not clean) or re.fullmatch(r"[_\-\—\–\|I]+", clean or ""):
                    sym = detect_andring_line_symbol(crop, debug_save_basepath=debug_base)
                    if sym:
                        text = sym

            if not text:
                continue

            # Keep one line only per label; if YOLO produces duplicates, keep first and ignore the rest
            if row[label_name]:
                continue

            row[label_name] = text

        data_rows.append(row)

    df = pd.DataFrame(data_rows, columns=["Image"] + LABELS)

    df = df.fillna("").astype(str)

    out_path = Path(raw_excel_out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(out_path, index=False, sheet_name="Raw")

    print("✅ RAW extraction saved:", out_path)


if __name__ == "__main__":
    process_folder()