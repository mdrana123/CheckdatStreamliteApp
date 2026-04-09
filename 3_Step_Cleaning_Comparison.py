# Cleaning.py
import re
from pathlib import Path

import pandas as pd
from xlsxwriter.utility import xl_col_to_name


# =========================
# Settings
# =========================
PROJECT_DIR = Path(r"C:\Users\admin\Checkdat_Stampel_AI")

RAW_XLSX = PROJECT_DIR / "Data_1_RawData.xlsx"
OUT_XLSX = PROJECT_DIR / "Data_2_Clean.xlsx"

RAW_SHEET = "Raw"


# =========================
# Helpers
# =========================
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.replace("\r", " ", regex=False)
        .str.replace("\n", " ", regex=False)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
        .str.upper()
    )
    return df


def norm_text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    s = str(v)
    s = s.replace("\r", " ").replace("\n", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def remove_label_spaced(s: str, word: str) -> str:
    letters = list(word)
    pattern = r"\b" + r"\s*".join(map(re.escape, letters)) + r"\b"
    return re.sub(pattern, "", s, flags=re.IGNORECASE).strip()


def _upper_clean(s: str) -> str:
    s = norm_text(s)
    return s.upper().strip()


def _strip_common_labels(s: str, labels: list[str]) -> str:
    out = s
    for w in labels:
        out = remove_label_spaced(out, w)
    return out.strip()


def _norm_for_compare(s: str) -> str:
    s = _upper_clean(s)
    s = s.replace(" ", "")
    s = s.replace("BBPO5", "BBP05").replace("BBPOS", "BBP05")
    s = s.replace("EO4", "E04")
    return s


def _image_base_from_filename(img: str) -> str:
    s = norm_text(img)
    if not s:
        return ""
    s = s.replace(".png", "")
    s = s.replace("_crop", "")
    m = re.search(r"(.*)_pdf_p", s)
    if m:
        s = m.group(1)
    m2 = re.search(r"(.*)_p\d+$", s)
    if m2:
        s = m2.group(1)
    return s.strip()


# =========================
# Cleaning functions
# =========================
ALLOWED_ANLAGGNINGSTYP_VALUES = [
    "TUNNEL", "VÄG", "BYGGNAD", "GEOTEKNIK", "MARK", "KANALISATION",
    "SPÅR", "SPÅRVÄXEL", "FÖRDELNINGSSTATION", "KOPPLINGSCENTRAL",
    "MATARLEDNING", "OMFORMARSTATION", "SEKTIONERINGSSTATION",
    "TRANSFORMATORSTATION", "BELYSNING", "DISTRIBUTIONSNÄT <1000V",
    "ELDRIFTLEDNINGSSYSTEM", "KRAFTFÖRSÖRJNING – TC",
    "KONTAKT-, HJÄLPKRAFT", "MOBILT RESERVELVERK", "NÄTSTATION",
    "TÅG OCH LOKVÄRME", "VÄXELVÄRME",
    "TEKNIKBYGGNAD (GÄLLER TEKNIKHUS, KIOSK, KUR)",
    "ÖVRIG BYGGNAD", "FÖRORENAT OMRÅDE", "MILJÖ (ÖVERGRIPANDE)",
    "DETEKTOR", "KAMERABEVAKNING",
    "PASSAGE-/ INBROTTSLARM", "TRAFIKINFORMATION",
]

ALLOWED_FORMAT_VALUES = [
    "A0", "A1", "A1F", "A2", "A2F", "A2FF",
    "A3", "A3F", "A3FF", "A3FFF",
    "A4",
]
ALLOWED_GRANSKNINGSSTATUS = [
    "UNDER ARBETE",
    "PRELIMINÄR",
    "FÖR GRANSKNING",
    "FÖR FASTSTÄLLELSE",
    "GODKÄND",
]
ALLOWED_RITNINGSTYP = [
    "SAMMANSTÄLLNINGSRITNING",
    "PLAN",
    "PROFIL",
    "DETALJRITNING",
    "SCHEMA",
    "STANDARDRITNING",
    "SEKTION"
]

def clean_anlaggningstyp_column(df: pd.DataFrame) -> pd.DataFrame:
    COL = "ANLÄGGNINGSTYP"
    if COL not in df.columns:
        return df

    s = df[COL].fillna("").astype(str)
    s = s.str.replace("\r", " ", regex=False)
    s = s.str.replace("\n", " ", regex=False)
    s = s.str.replace(r"\s+", " ", regex=True)
    s = s.str.strip()

    s = s.apply(lambda x: remove_label_spaced(x, "ANLÄGGNINGSTYP"))
    s = s.apply(lambda x: remove_label_spaced(x, "ANLAGGNINGSTYP"))

    s = s.str.replace(r"\bTUNNEI\b", "TUNNEL", regex=True)
    s = s.str.replace(r"\bTUNNE\b", "TUNNEL", regex=True)

    s = s.replace({"nan": "", "NaN": "", "NAN": ""})

    df[COL] = s
    return df


def clean_datum_column(df: pd.DataFrame) -> pd.DataFrame:
    COL = "DATUM"
    if COL not in df.columns:
        return df

    def _clean(v) -> str:
        s = norm_text(v)
        s = remove_label_spaced(s, "DATUM").strip()
        s = re.sub(r"^(\d{4}-\d{2}-\d)/$", r"\g<1>1", s)
        s = re.sub(r"^(\d{4}-\d{2}-)\s*(\d)\s*$", r"\g<1>0\g<2>", s)
        return s

    df[COL] = df[COL].apply(_clean)
    return df


def clean_format_column(df: pd.DataFrame) -> pd.DataFrame:
    COL = "FORMAT"
    if COL not in df.columns:
        return df

    def _clean(v) -> str:
        s = _upper_clean(v)
        s = _strip_common_labels(s, ["FORMAT"])
        s = s.replace(" ", "")

        s = s.replace("AO", "A0")
        s = s.replace("A-0", "A0")
        s = s.replace("A1F F", "A1FF").replace("A2F F", "A2FF").replace("A3F F", "A3FF")
        s = s.replace("A3FFFF", "A3FFF")

        return s.strip()

    df[COL] = df[COL].apply(_clean)
    return df


def clean_godkand_av_column(df: pd.DataFrame) -> pd.DataFrame:
    COL = "GODKÄND AV"
    if COL not in df.columns:
        return df

    s = df[COL].astype(str).fillna("")
    s = s.str.replace("\r", " ").str.replace("\n", " ")
    s = s.str.replace(r"\s+", " ", regex=True).str.strip()
    s = s.str.replace(
        r"\bMEHRAN KETABATI SAMMANSATT RITNING SAMMANSATT RITNING\b",
        "MEHRAN KETABATI",
        case=False,
        regex=True,
    )

    def _rm_labels(x: str) -> str:
        x = remove_label_spaced(x, "GODKÄND AV")
        x = remove_label_spaced(x, "GODKAND AV")
        x = remove_label_spaced(x, "ODKÄND AV")
        x = remove_label_spaced(x, "ODKAND AV")
        return x.strip()

    df[COL] = s.apply(_rm_labels)
    return df


COL_STATUS = "GRANSKNINGSSTATUS/SYFTE"


def clean_granskningsstatus_syfte_column(df: pd.DataFrame) -> pd.DataFrame:
    if COL_STATUS not in df.columns:
        return df

    def _clean(v) -> str:
        s = _upper_clean(v)
        s = _strip_common_labels(s, ["GRANSKNINGSSTATUS/SYFTE", "GRANSKNINGSSTATUS", "SYFTE"])
        s = s.replace("PRELIMINAR", "PRELIMINÄR")
        s = s.replace("FOR GRANSKNING", "FÖR GRANSKNING")
        s = s.replace("GODKAND", "GODKÄND")
        s = re.sub(r"\s+", " ", s).strip()
        if s in ["FÖRGRANSKNING", "FORGRANSKNING"]:
            s = "FÖR GRANSKNING"
        return s

    df[COL_STATUS] = df[COL_STATUS].apply(_clean)
    return df


def clean_granskad_av_column(df: pd.DataFrame) -> pd.DataFrame:
    COL = "GRANSKAD AV"
    if COL not in df.columns:
        return df

    def _clean(v) -> str:
        s = norm_text(v)
        s = _strip_common_labels(s, ["GRANSKAD AV", "GRANSKAD", "RANSKAD AV"])
        return s

    df[COL] = df[COL].apply(_clean)
    return df


ALLOWED_HANDLINGSTYP = [
   "SAMRÅDSUNDERLAG",
   "SAMRÅDSHANDLING",
   "GRANSKNINGSHANDLING",
   "FASTSTÄLLELSEHANDLING",
   "FÖRSLAGSHANDLING",
   "SYTEMHANDLING",
   "FÖRVALTNINGSHANDLING",
   "TYPRITNING",
   "STANDARDSRITNING",
   "BYGGHANDLING",
   "RELATIONSHANDLING"
]


def clean_handlingstyp_column(df: pd.DataFrame) -> pd.DataFrame:
    COL = "HANDLINGSTYP"
    if COL not in df.columns:
        return df

    def _clean(v) -> str:
        s = _upper_clean(v)
        s = _strip_common_labels(s, ["HANDLINGSTYP"])
        s = re.sub(r"\s+", " ", s).strip()
        s = s.replace("FORSLAGHANDLING", "FÖRSLAGHANDLING")
        s = s.replace("FORVALTNINGSDATA", "FÖRVALTNINGSDATA")
        return s

    df[COL] = df[COL].apply(_clean)
    return df


def clean_skapad_av_column(df: pd.DataFrame) -> pd.DataFrame:
    COL = "SKAPAD AV"
    if COL not in df.columns:
        return df

    def _clean(v) -> str:
        s = norm_text(v)
        s = _strip_common_labels(s, ["SKAPAD AV", "SKAPAD"])
        s = re.sub(r"\s+", " ", s).strip()
        if not s:
            return ""
        parts = s.split()
        if len(parts) >= 2:
            return f"{parts[0]} {parts[1]}".strip()
        return s

    df[COL] = df[COL].apply(_clean)
    return df


def clean_beskrivning_2_column(df: pd.DataFrame) -> pd.DataFrame:
    COL = "BESKRIVNING_2"
    if COL not in df.columns:
        return df

    def _clean(v) -> str:
        s = norm_text(v)
        s = _strip_common_labels(s, ["BESKRIVNING_2", "BESKRIVNING"])
        s = re.sub(r"\s+", " ", s).strip()
        if not s:
            return ""

        if s.startswith("3"):
            s = "B" + s[1:]

        parts = s.split()
        if len(parts) >= 2:
            return f"{parts[0]} {parts[1]}"
        return s

    df[COL] = df[COL].apply(_clean)
    return df


def _is_garbage_beskrivning3(s: str) -> bool:
    t = norm_text(s)
    if not t:
        return False

    up = t.upper().strip()

    if up in {"BEFINTLIGA LEDNINGAR", "BELYSNING"}:
        return True

    if re.match(r"^(AL|IVY|IV|TV)\b", up):
        return True

    pipe_count = up.count("|")
    letters = re.findall(r"[A-ZÅÄÖ]", up)
    if pipe_count >= 1 and len(letters) <= 6:
        return True

    core = re.sub(r"[^A-Z]", "", up)
    if core:
        junk_ratio = sum(core.count(ch) for ch in "UIV") / max(len(core), 1)
        if len(core) >= 8 and junk_ratio >= 0.65:
            return True

    return False


def clean_beskrivning_3_column(df: pd.DataFrame) -> pd.DataFrame:
    COL = "BESKRIVNING_3"
    if COL not in df.columns:
        return df

    def _clean(v) -> str:
        s = norm_text(v)
        s = _strip_common_labels(s, ["BESKRIVNING_3", "BESKRIVNING"])
        s = re.sub(r"\s+", " ", s).strip()
        if not s:
            return ""
        if _is_garbage_beskrivning3(s):
            return ""
        return s

    df[COL] = df[COL].apply(_clean)
    return df

def clean_delomrade_bandel_column(df: pd.DataFrame) -> pd.DataFrame:
    """
    DELOMRÅDE/BANDEL:
      - If the first word starts with 'I' (OCR artifact),
        remove ONLY that leading 'I' and keep the rest.
      Examples:
        "I604" -> "604"
        "IBANDEL 604" -> "BANDEL 604"
        "I 604" -> "604"
    """
    COL = "DELOMRÅDE/BANDEL"
    if COL not in df.columns:
        return df

    def _clean(v) -> str:
        s = norm_text(v)
        s = re.sub(r"\s+", " ", s).strip()
        if not s:
            return ""

        # Case 1: "I WORD"
        s = re.sub(r"^I\s+", "", s, flags=re.IGNORECASE)

        # Case 2: "IXXXX"
        parts = s.split(" ", 1)
        first = parts[0]
        rest = parts[1] if len(parts) == 2 else ""

        if first.upper().startswith("I") and len(first) > 1 and first[1].isalnum():
            first = first[1:]

        s = (first + (" " + rest if rest else "")).strip()
        return s

    df[COL] = df[COL].apply(_clean)
    return df
def clean_beskrivning_4_column(df: pd.DataFrame) -> pd.DataFrame:
    """
    BESKRIVNING_4:
      - If the first word starts with "I" like "Ixxx" (OCR artifact),
        remove ONLY that leading "I" and keep the rest of the word.
      Examples:
        "IBELYSNING" -> "BELYSNING"
        "I BEFINTLIGA" -> "BEFINTLIGA"  (handles both cases)
    """
    COL = "BESKRIVNING_4"
    if COL not in df.columns:
        return df

    def _clean(v) -> str:
        s = norm_text(v)
        s = _strip_common_labels(s, ["BESKRIVNING_4", "BESKRIVNING"])
        s = re.sub(r"\s+", " ", s).strip()
        if not s:
            return ""

        # Case 1: "I WORD..." -> drop the standalone leading I
        s = re.sub(r"^I\s+", "", s, flags=re.IGNORECASE)

        # Case 2: "IXXXX" -> remove only the first 'I' from the first token
        # (do not touch real words later in the string)
        parts = s.split(" ", 1)
        first = parts[0]
        rest = parts[1] if len(parts) == 2 else ""
        if first.upper().startswith("I") and len(first) > 1 and first[1].isalpha():
            first = first[1:]
        s = (first + (" " + rest if rest else "")).strip()

        return s

    df[COL] = df[COL].apply(_clean)
    return df


TEKNIK_ALLOWED = [
    "VÄGUTFORMNING OCH TRAFIK",
    "VATTEN OCH AVLOPP",
    "EL",
    "BANOMGIVNING",
    "BRO",
    "TUNNEL",
    "GEOTEKNIK",
    "TEKNIKÖVERGRIPANDE",
    "HYDROGEOLOGI GEOTEKNIK",
    "GEOTEKNIK HYDROGEOLOGI",
    
]


def clean_teknikomrade_column(df: pd.DataFrame) -> pd.DataFrame:
    COL = "TEKNIKOMRÅDE"
    if COL not in df.columns:
        return df

    def _clean(v) -> str:
        s = _upper_clean(v)
        s = _strip_common_labels(s, ["TEKNIKOMRÅDE", "TEKNIKOMRADE"])
        s = s.replace("BANGARDSANLAGGNING", "BANGÅRDSANLÄGGNING")
        s = s.replace("ELANLAGGNING", "ELANLÄGGNING")
        s = s.replace("SIGNALANLAGGNING", "SIGNALANLÄGGNING")
        s = s.replace("TELEANLAGGNING", "TELEANLÄGGNING")
        s = s.replace("MILJO", "MILJÖ")
        s = s.replace("VAGKROPP", "VÄGKROPP")
        s = s.replace("BANOVERBYGGNAD", "BANÖVERBYGGNAD")
        s = re.sub(r"\s+", " ", s).strip()
        return s

    df[COL] = df[COL].apply(_clean)
    return df


def clean_uppdragsnummer_column(df: pd.DataFrame) -> pd.DataFrame:
    COL = "UPPDRAGSNUMMER"
    if COL not in df.columns:
        return df

    def _clean(v) -> str:
        s = norm_text(v)
        s = _strip_common_labels(s, ["UPPDRAGSNUMMER", "UPPDRAG", "UPPDRAGS NR", "UPPDRAGSNR"])
        s = s.replace(" ", "")
        s = re.sub(r"[^0-9A-Z\-_]", "", s.upper())
        return s

    df[COL] = df[COL].apply(_clean)
    return df


def clean_kilometer_meter_column(df: pd.DataFrame) -> pd.DataFrame:
    COL = "KILOMETER & METER"
    if COL not in df.columns:
        return df

    def _clean(v) -> str:
        s = norm_text(v)
        s = _strip_common_labels(s, ["KILOMETER & METER", "KILOMETER", "METER"])
        s = s.replace(",", ".")
        s = s.replace(" ", "")

        m = re.fullmatch(r"(\d+)\+(\d{1,3})", s)
        if m:
            km = m.group(1)
            meter = m.group(2).zfill(3)
            return f"{km}+{meter}"

        if re.fullmatch(r"\d{4,}", s):
            km = s[:-3]
            meter = s[-3:]
            return f"{km}+{meter}"

        return s

    df[COL] = df[COL].apply(_clean)
    return df


def clean_andr_column(df: pd.DataFrame) -> pd.DataFrame:
    COL = "ÄNDR"
    if COL not in df.columns:
        return df

    def _clean(v) -> str:
        s = _upper_clean(v)
        s = _strip_common_labels(s, ["ÄNDR", "ANDR", "ÄND"])
        s = s.replace(",", ".")
        s = s.replace(" ", "")
        s = re.sub(r"^F\.?1$", "F.1", s)
        s = re.sub(r"^F\.?2$", "F.2", s)
        s = re.sub(r"^A\.?1$", "A.1", s)
        s = re.sub(r"^A\.?2$", "A.2", s)
        return s

    df[COL] = df[COL].apply(_clean)
    return df


def clean_title_column(df: pd.DataFrame) -> pd.DataFrame:
    COL = "TITLE"
    if COL not in df.columns:
        return df

    def _clean(v) -> str:
        s = norm_text(v)
        s = _strip_common_labels(s, ["TITLE", "T I T L E"])
        s = re.sub(r"\s+", " ", s).strip()
        return s

    df[COL] = df[COL].apply(_clean)
    return df


def clean_ritningsnummer_column(df: pd.DataFrame) -> pd.DataFrame:
    COL = "RITNINGSNUMMER"
    if COL not in df.columns:
        return df

    def _clean(v) -> str:
        s = norm_text(v)
        s = _strip_common_labels(s, ["RITNINGSNUMMER", "RITNING", "RITN NR", "RITNINGSNR"])
        s = s.replace(" ", "").strip()
        if not s:
            return ""
        if s[0] != "0":
            s = s[1:]
        return s

    df[COL] = df[COL].apply(_clean)
    return df


def clean_ritningsnummer_projekt(df: pd.DataFrame) -> pd.DataFrame:
    COL = "RITNINGSNUMMER_PROJEKT"
    IMG = "IMAGE"

    if COL not in df.columns or IMG not in df.columns:
        return df

    def _fix_proj(s: str) -> str:
        s = norm_text(s)
        s = _strip_common_labels(
            s,
            ["RITNINGSNUMMER_PROJEKT", "RITNINGSNUMMER", "RITNING", "RITN NR", "RITNINGSNR"],
        )
        s = s.replace(" ", "")
        s = s.replace("BBPO5", "BBP05").replace("BBPOS", "BBP05")
        s = s.replace("EO4", "E04")
        if s.upper().startswith("BP"):
            s = "BBP" + s[2:]
        s = re.sub(r"(0_0-)(\d{3})$", lambda m: m.group(1) + "0" + m.group(2), s)
        return s.strip()

    new_vals = []
    for proj, img in zip(df[COL].tolist(), df[IMG].tolist()):
        proj_fixed = _fix_proj(proj)
        img_base = _image_base_from_filename(img)

        if not proj_fixed and img_base:
            new_vals.append(img_base)
            continue

        if proj_fixed and img_base:
            if _norm_for_compare(proj_fixed) != _norm_for_compare(img_base):
                new_vals.append(img_base)
            else:
                new_vals.append(proj_fixed)
        else:
            new_vals.append(proj_fixed)

    df[COL] = new_vals
    return df


# =========================
# Excel formatting (conditional formatting)
# =========================
def _col_letter_for(df: pd.DataFrame, col: str) -> str:
    idx = df.columns.get_loc(col)
    return xl_col_to_name(idx)


def apply_format_anlaggningstyp(df, worksheet, workbook, start_row: int, end_row: int):
    COL = "ANLÄGGNINGSTYP"
    if COL not in df.columns:
        return
    red = workbook.add_format({"bg_color": "#E90721", "font_color": "#FFFFFF"})
    col_letter = _col_letter_for(df, COL)
    cell = f"{col_letter}{start_row}"
    or_parts = [f'{cell}="{v}"' for v in ALLOWED_ANLAGGNINGSTYP_VALUES]
    formula = f'=AND({cell}<>"",NOT(OR({",".join(or_parts)})))'
    worksheet.conditional_format(
        f"{col_letter}{start_row}:{col_letter}{end_row}",
        {"type": "formula", "criteria": formula, "format": red},
    )


def apply_format_format(df, worksheet, workbook, start_row: int, end_row: int):
    COL = "FORMAT"
    if COL not in df.columns:
        return
    red = workbook.add_format({"bg_color": "#E90721", "font_color": "#FFFFFF"})
    col_letter = _col_letter_for(df, COL)
    cell = f"{col_letter}{start_row}"
    allowed_parts = [f'{cell}="{v}"' for v in ALLOWED_FORMAT_VALUES]
    formula = f'=AND({cell}<>"",NOT(OR({",".join(allowed_parts)})))'
    worksheet.conditional_format(
        f"{col_letter}{start_row}:{col_letter}{end_row}",
        {"type": "formula", "criteria": formula, "format": red},
    )


def apply_format_bandel(df: pd.DataFrame, worksheet, workbook, start_row: int, end_row: int):
    COL = "BANDEL"
    if COL not in df.columns:
        return
    red_format = workbook.add_format({"bg_color": "#E90721", "font_color": "#FFFFFF"})
    idx = df.columns.get_loc(COL)
    col_letter = xl_col_to_name(idx)
    cell = f"{col_letter}{start_row}"
    formula = f'=AND({cell}<>"",TRIM({cell})<>"604")'
    worksheet.conditional_format(
        f"{col_letter}{start_row}:{col_letter}{end_row}",
        {"type": "formula", "criteria": formula, "format": red_format},
    )


def apply_format_handlingstyp(df, worksheet, workbook, start_row: int, end_row: int):
    COL = "HANDLINGSTYP"
    if COL not in df.columns:
        return
    red = workbook.add_format({"bg_color": "#E90721", "font_color": "#FFFFFF"})
    col_letter = _col_letter_for(df, COL)
    cell = f"{col_letter}{start_row}"
    allowed = [v.upper() for v in ALLOWED_HANDLINGSTYP]
    parts = [f'UPPER(TRIM({cell}))<>"{v}"' for v in allowed]
    formula = f'=AND({cell}<>"",AND({",".join(parts)}))'
    worksheet.conditional_format(
        f"{col_letter}{start_row}:{col_letter}{end_row}",
        {"type": "formula", "criteria": formula, "format": red},
    )


def apply_format_teknikomrade(df: pd.DataFrame, worksheet, workbook, start_row: int, end_row: int):
    COL = "TEKNIKOMRÅDE"
    if COL not in df.columns:
        return
    red_fmt = workbook.add_format({"bg_color": "#E90721", "font_color": "#FFFFFF"})
    col_letter = _col_letter_for(df, COL)
    cell_ref = f"{col_letter}{start_row}"
    norm = f"UPPER(TRIM({cell_ref}))"
    allowed_parts = [f'{norm}="{v.upper()}"' for v in TEKNIK_ALLOWED]
    formula = f'=AND({cell_ref}<>"",NOT(OR({",".join(allowed_parts)})))'
    worksheet.conditional_format(
        f"{col_letter}{start_row}:{col_letter}{end_row}",
        {"type": "formula", "criteria": formula, "format": red_fmt},
    )


def apply_format_datum(df, worksheet, workbook, start_row: int, end_row: int):
    COL = "DATUM"
    if COL not in df.columns:
        return
    red = workbook.add_format({"bg_color": "#E90721", "font_color": "#FFFFFF"})
    col_letter = _col_letter_for(df, COL)
    cell = f"{col_letter}{start_row}"
    formula = (
        f'=AND('
        f'{cell}<>"",'
        f'NOT(AND('
        f'LEN({cell})=10,'
        f'MID({cell},5,1)="-",'
        f'MID({cell},8,1)="-",'
        f'ISNUMBER(--LEFT({cell},4)),'
        f'ISNUMBER(--MID({cell},6,2)),'
        f'ISNUMBER(--RIGHT({cell},2))'
        f'))'
        f')'
    )
    worksheet.conditional_format(
        f"{col_letter}{start_row}:{col_letter}{end_row}",
        {"type": "formula", "criteria": formula, "format": red},
    )


def apply_format_ritningsnummer_vs_image_base(df, worksheet, workbook, start_row: int, end_row: int):
    COL_PROJ = "RITNINGSNUMMER_PROJEKT"
    COL_IMG = "IMAGE"
    if COL_PROJ not in df.columns or COL_IMG not in df.columns:
        return
    red = workbook.add_format({"bg_color": "#E90721", "font_color": "#FFFFFF"})
    proj_letter = _col_letter_for(df, COL_PROJ)
    img_letter = _col_letter_for(df, COL_IMG)
    proj_cell = f"{proj_letter}{start_row}"
    img_cell = f"{img_letter}{start_row}"

    img_nopng = f"SUBSTITUTE({img_cell},\".png\",\"\")"
    img_nocrop = f"SUBSTITUTE({img_nopng},\"_crop\",\"\")"
    img_before_pdfp = f"IFERROR(LEFT({img_nocrop},FIND(\"_pdf_p\",{img_nocrop})-1),{img_nocrop})"
    img_base = f"IFERROR(LEFT({img_before_pdfp},FIND(\"_p\",{img_before_pdfp})-1),{img_before_pdfp})"

    proj_norm = f"SUBSTITUTE(TRIM({proj_cell}),\" \",\"\")"
    proj_norm = f"SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({proj_norm},\"BBPO5\",\"BBP05\"),\"BBPOS\",\"BBP05\"),\"EO4\",\"E04\")"

    img_norm = f"SUBSTITUTE(TRIM({img_base}),\" \",\"\")"
    img_norm = f"SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({img_norm},\"BBPO5\",\"BBP05\"),\"BBPOS\",\"BBP05\"),\"EO4\",\"E04\")"

    formula = f"=AND({proj_cell}<>\"\",{img_cell}<>\"\",{proj_norm}<>{img_norm})"
    worksheet.conditional_format(
        f"{proj_letter}{start_row}:{proj_letter}{end_row}",
        {"type": "formula", "criteria": formula, "format": red},
    )

def apply_format_granskningsstatus(df, worksheet, workbook, start_row: int, end_row: int):
    COL = "GRANSKNINGSSTATUS/SYFTE"
    if COL not in df.columns:
        return

    red = workbook.add_format({"bg_color": "#E90721", "font_color": "#FFFFFF"})
    col_letter = _col_letter_for(df, COL)
    cell = f"{col_letter}{start_row}"

    allowed = [v.upper() for v in ALLOWED_GRANSKNINGSSTATUS]
    parts = [f'UPPER(TRIM({cell}))<>"{v}"' for v in allowed]

    formula = f'=AND({cell}<>"",AND({",".join(parts)}))'

    worksheet.conditional_format(
        f"{col_letter}{start_row}:{col_letter}{end_row}",
        {"type": "formula", "criteria": formula, "format": red},
    )
def apply_format_ritningstyp(df, worksheet, workbook, start_row: int, end_row: int):
    COL = "RITNINGSTYP"
    if COL not in df.columns:
        return

    red = workbook.add_format({"bg_color": "#E90721", "font_color": "#FFFFFF"})
    col_letter = _col_letter_for(df, COL)
    cell = f"{col_letter}{start_row}"

    allowed = [v.upper() for v in ALLOWED_RITNINGSTYP]
    parts = [f'UPPER(TRIM({cell}))<>"{v}"' for v in allowed]
    formula = f'=AND({cell}<>"",AND({",".join(parts)}))'

    worksheet.conditional_format(
        f"{col_letter}{start_row}:{col_letter}{end_row}",
        {"type": "formula", "criteria": formula, "format": red},
    )


def apply_format_blad_vs_proj_lastdigits_skip_red(df, worksheet, workbook, start_row: int, end_row: int):
    COL_BLAD = "BLAD"
    COL_PROJ = "RITNINGSNUMMER_PROJEKT"
    COL_IMG = "IMAGE"
    if COL_BLAD not in df.columns or COL_PROJ not in df.columns or COL_IMG not in df.columns:
        return

    red = workbook.add_format({"bg_color": "#E90721", "font_color": "#FFFFFF"})
    blad_letter = _col_letter_for(df, COL_BLAD)
    proj_letter = _col_letter_for(df, COL_PROJ)
    img_letter = _col_letter_for(df, COL_IMG)

    blad_cell = f"{blad_letter}{start_row}"
    proj_cell = f"{proj_letter}{start_row}"
    img_cell = f"{img_letter}{start_row}"

    img_nopng = f"SUBSTITUTE({img_cell},\".png\",\"\")"
    img_nocrop = f"SUBSTITUTE({img_nopng},\"_crop\",\"\")"
    img_before_pdfp = f"IFERROR(LEFT({img_nocrop},FIND(\"_pdf_p\",{img_nocrop})-1),{img_nocrop})"
    img_base = f"IFERROR(LEFT({img_before_pdfp},FIND(\"_p\",{img_before_pdfp})-1),{img_before_pdfp})"

    proj_norm = f"SUBSTITUTE(TRIM({proj_cell}),\" \",\"\")"
    proj_norm = f"SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({proj_norm},\"BBPO5\",\"BBP05\"),\"BBPOS\",\"BBP05\"),\"EO4\",\"E04\")"

    img_norm = f"SUBSTITUTE(TRIM({img_base}),\" \",\"\")"
    img_norm = f"SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({img_norm},\"BBPO5\",\"BBP05\"),\"BBPOS\",\"BBP05\"),\"EO4\",\"E04\")"

    proj_is_red_condition = f"AND({proj_cell}<>\"\",{img_cell}<>\"\",{proj_norm}<>{img_norm})"
    skip_if_proj_red = f"NOT({proj_is_red_condition})"

    blad_num = f'IFERROR(VALUE(TRIM({blad_cell})),-999999)'
    tail3_num = f'IFERROR(VALUE(RIGHT({proj_cell},3)),-111111)'
    tail4_num = f'IFERROR(VALUE(RIGHT({proj_cell},4)),-222222)'

    formula = (
        f'=AND('
        f'{skip_if_proj_red},'
        f'TRIM({blad_cell})<>"",'
        f'TRIM({proj_cell})<>"",'
        f'NOT(OR({blad_num}={tail3_num},{blad_num}={tail4_num}))'
        f')'
    )
    worksheet.conditional_format(
        f"{blad_letter}{start_row}:{blad_letter}{end_row}",
        {"type": "formula", "criteria": formula, "format": red},
    )


# =========================
# Main pipeline
# =========================
def main():
    if not RAW_XLSX.exists():
        raise FileNotFoundError(f"Hittar inte inputfilen: {RAW_XLSX}")

    df = pd.read_excel(
        RAW_XLSX,
        sheet_name=RAW_SHEET,
        dtype=str,
        keep_default_na=False,
    )
    df = normalize_columns(df)

    df = clean_anlaggningstyp_column(df)
    df = clean_datum_column(df)
    df = clean_format_column(df)
    df = clean_godkand_av_column(df)
    df = clean_granskningsstatus_syfte_column(df)
    df = clean_granskad_av_column(df)
    df = clean_handlingstyp_column(df)
    df = clean_skapad_av_column(df)

    df = clean_beskrivning_2_column(df)
    df = clean_beskrivning_3_column(df)
    df = clean_beskrivning_4_column(df)  
    df = clean_delomrade_bandel_column(df)
    df = clean_teknikomrade_column(df)
    df = clean_uppdragsnummer_column(df)
    df = clean_kilometer_meter_column(df)
    df = clean_andr_column(df)
    df = clean_title_column(df)

    df = clean_ritningsnummer_column(df)
    df = clean_ritningsnummer_projekt(df)

    with pd.ExcelWriter(OUT_XLSX, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Clean", index=False)

        workbook = writer.book
        worksheet = writer.sheets["Clean"]

        text_fmt = workbook.add_format({"num_format": "@"})
        TEXT_COLS = [
            "BLAD",
            "BANDEL",
            "UPPDRAGSNUMMER",
            "RITNINGSNUMMER_PROJEKT",
            "RITNINGSNUMMER",
            "KILOMETER & METER",
        ]
        for c in TEXT_COLS:
            if c in df.columns:
                idx = df.columns.get_loc(c)
                worksheet.set_column(idx, idx, 22, text_fmt)

        start_row = 2
        end_row = len(df) + 1

        apply_format_anlaggningstyp(df, worksheet, workbook, start_row, end_row)
        apply_format_bandel(df, worksheet, workbook, start_row, end_row)
        apply_format_datum(df, worksheet, workbook, start_row, end_row)
        apply_format_format(df, worksheet, workbook, start_row, end_row)
        apply_format_handlingstyp(df, worksheet, workbook, start_row, end_row)
        apply_format_teknikomrade(df, worksheet, workbook, start_row, end_row)

        apply_format_ritningsnummer_vs_image_base(df, worksheet, workbook, start_row, end_row)
        apply_format_blad_vs_proj_lastdigits_skip_red(df, worksheet, workbook, start_row, end_row)
        apply_format_granskningsstatus(df, worksheet, workbook, start_row, end_row)
        apply_format_ritningstyp(df, worksheet, workbook, start_row, end_row)

        worksheet.set_column(0, len(df.columns) - 1, 22)

    print(f"✅ Klar! Sparade: {OUT_XLSX}")


if __name__ == "__main__":
    main()