import os
import time
import re
import tempfile
import hashlib
from pathlib import Path

import pandas as pd
import win32com.client as win32
from PIL import Image, ImageOps

# =========================
# PATHS / SETTINGS
# =========================
EMAIL_PREVIEW_FILE = r"C:\Users\admin\Checkdat_Stampel_AI\Data_3_Mismatch_Report.xlsx"
STAMP_FOLDER       = r"C:\Users\admin\Checkdat_Stampel_AI\3_cropsImages_flat"

OUTLOOK_ACCOUNT_SMTP = "mdalamgirhossain.rana@utb.ecutbildning.se"
TO_EMAIL             = "mdalamgirhossain.rana@utb.ecutbildning.se"

INTRO_TEXT = "Hej,\n\nNedan är fel som hittades i stämplarna. Detta är ett automatiskt utkast.\n"
SEND_NOW = False  # keep False to create draft

# If your images might be inside subfolders, set True
SEARCH_RECURSIVE = True

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}

# =========================
# IMAGE COMPRESSION SETTINGS
# =========================
# Reduce draft size heavily by resizing + JPEG compression
MAX_WIDTH = 300            # try 900 if still big
JPEG_QUALITY = 70           # try 60 if still big
JPEG_OPTIMIZE = True
JPEG_PROGRESSIVE = True

# Put compressed copies here (temp folder)
TEMP_SUBFOLDER = "stamp_mail_compressed"
CLEAN_TEMP_OLDER_THAN_DAYS = 3  # set 0 to disable cleanup

# =========================
# OUTLOOK HELPERS
# =========================
def get_outlook_account(namespace, smtp_address: str):
    smtp_address = smtp_address.lower().strip()
    for acc in namespace.Accounts:
        try:
            if str(acc.SmtpAddress).lower().strip() == smtp_address:
                return acc
        except Exception:
            pass
    return None

def set_inline_attachment_cid(att, cid: str):
    # Content-ID must be stored as <cid> for multi-inline images in Outlook
    att.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F", f"<{cid}>"
    )
    att.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3713001F", cid
    )

# =========================
# HTML HELPERS
# =========================
def html_escape(s: str) -> str:
    s = "" if s is None else str(s)
    return (s.replace("&", "&amp;")
             .replace("<", "&lt;")
             .replace(">", "&gt;")
             .replace('"', "&quot;")
             .replace("'", "&#39;"))

def error_details_to_ul(error_details: str) -> str:
    items = []
    for part in str(error_details).split(";"):
        part = part.strip()
        if part:
            items.append(f"<li>{html_escape(part)}</li>")
    if not items:
        return "<ul><li>Inga detaljer hittades.</li></ul>"
    return "<ul>\n" + "\n".join(items) + "\n</ul>"

def build_section_html(title: str, error_details: str, cid: str, index: int) -> str:
    return f"""
    <hr style="border:none;border-top:1px solid #e5e5e5;margin:14px 0;">
    <div style="padding:8px 0;">
      <div style="font-size:12pt; font-weight:bold; margin-bottom:6px;">
        {index}. {html_escape(title)}
      </div>

      <div style="margin:6px 0;"><b>Bild:</b></div>
      <img src="cid:{cid}" style="max-width:750px; border:1px solid #ddd; display:block; margin:6px 0 10px 0;">

      <div style="margin:6px 0;"><b>Fel:</b></div>
      {error_details_to_ul(error_details)}
    </div>
    """

def build_full_html(intro_text: str, sections_html: str, total_rows: int, matched: int, missing: int) -> str:
    intro_html = "<br>".join(html_escape(intro_text).split("\n"))
    summary = f"""
    <div style="padding:10px; background:#f7f7f7; border:1px solid #eee; border-radius:8px; margin-top:14px;">
      <b>Sammanfattning</b><br>
      Rader i Excel: {total_rows}<br>
      Bilder inkluderade: {matched}<br>
      Bilder saknas: {missing}<br>
    </div>
    """
    return f"""
    <html>
      <body style="font-family:Calibri; font-size:11pt;">
        <p>{intro_html}</p>
        {sections_html}
        {summary}
        <p>Med vänliga hälsningar,<br>Rana</p>
      </body>
    </html>
    """

# =========================
# IMAGE MATCHING
# =========================
def collect_images(folder: str, recursive: bool):
    paths = []
    if recursive:
        for root, _, files in os.walk(folder):
            for fn in files:
                ext = os.path.splitext(fn)[1].lower()
                if ext in IMAGE_EXTS:
                    paths.append(os.path.join(root, fn))
    else:
        for fn in os.listdir(folder):
            p = os.path.join(folder, fn)
            if os.path.isfile(p) and os.path.splitext(fn)[1].lower() in IMAGE_EXTS:
                paths.append(p)
    return paths

def norm_key(s: str) -> str:
    s = "" if s is None else str(s)
    s = os.path.basename(s.strip())
    s = os.path.splitext(s)[0]
    s = re.sub(r"\s+", "", s)
    return s.lower()

def build_indexes(image_paths):
    by_filename = {}  # exact filename lower -> path
    by_base = {}      # base -> list(paths)
    for p in image_paths:
        fn = os.path.basename(p)
        by_filename[fn.lower()] = p
        base = norm_key(fn)
        by_base.setdefault(base, []).append(p)
    return by_filename, by_base

def resolve_image(file_value: str, by_filename, by_base, all_images):
    raw = "" if file_value is None else str(file_value).strip()
    if not raw or raw.lower() == "nan":
        return None

    name_only = os.path.basename(raw)

    # 1) exact filename match
    p = by_filename.get(name_only.lower())
    if p and os.path.exists(p):
        return p

    # 2) exact base match (ignore ext/spaces)
    base = norm_key(name_only)
    if base in by_base:
        return sorted(by_base[base], key=lambda x: len(os.path.basename(x)))[0]

    # 3) contains match (excel base inside image base)
    hits = []
    for ip in all_images:
        ibase = norm_key(os.path.basename(ip))
        if base and base in ibase:
            hits.append(ip)
    if hits:
        hits.sort(key=lambda x: (len(os.path.basename(x)), os.path.basename(x).lower()))
        return hits[0]

    return None

# =========================
# IMAGE DOWNSIZE/COMPRESS
# =========================
def _temp_dir() -> str:
    d = os.path.join(tempfile.gettempdir(), TEMP_SUBFOLDER)
    os.makedirs(d, exist_ok=True)
    return d

def cleanup_temp_folder(days: int):
    if days <= 0:
        return
    d = _temp_dir()
    cutoff = time.time() - days * 86400
    for p in Path(d).glob("*.jpg"):
        try:
            if p.stat().st_mtime < cutoff:
                p.unlink(missing_ok=True)
        except Exception:
            pass

def compressed_copy_path(src_path: str) -> str:
    """
    Deterministic output name so we reuse existing compressed versions.
    """
    d = _temp_dir()
    st = os.stat(src_path)
    key = f"{src_path}|{st.st_size}|{int(st.st_mtime)}|w{MAX_WIDTH}|q{JPEG_QUALITY}"
    h = hashlib.md5(key.encode("utf-8")).hexdigest()  # stable short id
    base = os.path.splitext(os.path.basename(src_path))[0]
    return os.path.join(d, f"{base}_{h}.jpg")

def compress_image_for_email(src_path: str) -> str:
    """
    Creates (or reuses) a resized/compressed JPEG copy and returns its path.
    """
    out_path = compressed_copy_path(src_path)
    if os.path.exists(out_path):
        return out_path

    with Image.open(src_path) as img:
        # Fix orientation using EXIF
        img = ImageOps.exif_transpose(img)

        # Convert to RGB for JPEG
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")

        w, h = img.size
        if w > MAX_WIDTH:
            new_h = int(h * (MAX_WIDTH / w))
            img = img.resize((MAX_WIDTH, new_h), Image.LANCZOS)

        img.save(
            out_path,
            format="JPEG",
            quality=JPEG_QUALITY,
            optimize=JPEG_OPTIMIZE,
            progressive=JPEG_PROGRESSIVE
        )

    return out_path

# =========================
# MAIN
# =========================
def main():
    cleanup_temp_folder(CLEAN_TEMP_OLDER_THAN_DAYS)

    if not os.path.exists(EMAIL_PREVIEW_FILE):
        raise FileNotFoundError(f"Mismatch file not found: {EMAIL_PREVIEW_FILE}")
    if not os.path.isdir(STAMP_FOLDER):
        raise FileNotFoundError(f"Stamp folder not found: {STAMP_FOLDER}")

    df = pd.read_excel(EMAIL_PREVIEW_FILE)
    df.columns = df.columns.astype(str).str.strip()

    # Ensure columns
    if "FILE" not in df.columns:
        df = df.rename(columns={df.columns[0]: "FILE"})
    if "ERROR_DETAILS" not in df.columns and len(df.columns) >= 2:
        df = df.rename(columns={df.columns[1]: "ERROR_DETAILS"})

    if "FILE" not in df.columns or "ERROR_DETAILS" not in df.columns:
        raise Exception(f"Expected columns FILE and ERROR_DETAILS. Found: {list(df.columns)}")

    # Collect images and build indexes
    all_images = collect_images(STAMP_FOLDER, SEARCH_RECURSIVE)
    by_filename, by_base = build_indexes(all_images)

    outlook = win32.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    account = get_outlook_account(namespace, OUTLOOK_ACCOUNT_SMTP)
    if not account:
        raise Exception(f"Outlook account not found for: {OUTLOOK_ACCOUNT_SMTP}")

    # Create Draft
    mail = outlook.CreateItem(0)  # 0 = olMailItem
    mail.SendUsingAccount = account
    mail.To = TO_EMAIL
    mail.Subject = "[Stamp Validation Errors] Samlad rapport"
    mail.BodyFormat = 2  # HTML

    sections = []
    matched = 0
    missing = 0
    idx_display = 0

    for _, row in df.iterrows():
        file_value = "" if pd.isna(row.get("FILE")) else str(row.get("FILE")).strip()
        error_details = "" if pd.isna(row.get("ERROR_DETAILS")) else str(row.get("ERROR_DETAILS")).strip()

        if not file_value or file_value.lower() == "nan":
            continue

        idx_display += 1
        img_path = resolve_image(file_value, by_filename, by_base, all_images)

        if not img_path:
            missing += 1
            sections.append(f"""
              <hr style="border:none;border-top:1px solid #e5e5e5;margin:14px 0;">
              <div style="padding:8px 0;">
                <div style="font-size:12pt; font-weight:bold; margin-bottom:6px;">
                  {idx_display}. {html_escape(file_value)}
                </div>
                <div style="color:#b00020; font-weight:bold; margin:6px 0;">
                  Bild saknas / kunde inte matchas i mappen
                </div>
                <div style="margin:6px 0;"><b>Fel:</b></div>
                {error_details_to_ul(error_details)}
              </div>
            """)
            continue

        # ✅ Compress before attaching (this is the key)
        small_path = compress_image_for_email(img_path)

        matched += 1
        cid = f"stamp_{idx_display}"

        att = mail.Attachments.Add(small_path)
        set_inline_attachment_cid(att, cid)

        # Show original filename (or compressed filename) in the title
        sections.append(build_section_html(os.path.basename(img_path), error_details, cid, idx_display))

    sections_html = "\n".join(sections) if sections else "<p><i>Inga fel att rapportera.</i></p>"
    mail.HTMLBody = build_full_html(INTRO_TEXT, sections_html, total_rows=len(df), matched=matched, missing=missing)

    # Save draft reliably
    mail.Display(False)
    mail.Save()
    time.sleep(0.2)
    mail.Save()

    try:
        print("✅ Saved. Folder:", mail.Parent.FolderPath)
    except Exception:
        print("✅ Saved draft (folder path not available).")

    print("Rows included (with image):", matched)
    print("Rows missing image:", missing)
    print("Compressed images stored in:", _temp_dir())
    print(f"Compression: MAX_WIDTH={MAX_WIDTH}px, JPEG_QUALITY={JPEG_QUALITY}")

    if SEND_NOW:
        mail.Send()
        print("✅ Sent email.")
    else:
        print("✅ Draft created (ONE mail). Check Outlook Drafts / Utkast.")

if __name__ == "__main__":
    main()
