"""
Detecteer type PDF: FACQ offerte of EPB installatievoorstel
"""
import pdfplumber
from io import BytesIO
from enum import Enum
import re

class PDFType(Enum):
    FACQ_OFFERTE = "facq_offerte"
    EPB_VOORSTEL = "epb_voorstel"
    UNKNOWN = "unknown"

# Sterke markers voor Vaillant/technisch schema + legende
VAILLANT_MARKERS = [
    "vaillant",
    "arotherm",
    "arotherm split",
    "vwl",                 # VWL 8.2 AS/IS
    "vrc720",
    "vr71",
    "vr940",
    "hydraulic module",
    "unistor",
    "vih rw",
    "vp rw",
    "legend",              # engels
    "line legend",
    "pumps",
    "functional valves",
    "safety units",
    "further armatures",
    "sensors vr10",
    "vr10",
]

# Indicator-codes zoals 1a, 2d, 10b komen typisch veel voor in schema-legendes
INDICATOR_RE = re.compile(r"\b\d{1,2}[a-z]{1,2}\b", re.IGNORECASE)

def _norm_text(s: str) -> str:
    s = (s or "").lower()
    s = re.sub(r"\s+", " ", s)
    return s

def _vaillant_schema_score(text: str) -> int:
    t = _norm_text(text)
    marker_hits = sum(1 for m in VAILLANT_MARKERS if m in t)
    indicator_hits = len(INDICATOR_RE.findall(t))
    # markers zwaar laten doorwegen, indicators gecapped
    return marker_hits * 3 + min(indicator_hits, 30)

def detect_pdf_type(pdf_bytes: bytes) -> PDFType:
    """
    Detecteer welk type PDF dit is o.b.v. content.

    FACQ Offerte kenmerken:
    - Bevat vaak: "FACQ", artikelnummers (5-6 cijfers), prijzen, hoeveelheden
    - Tabelvorm met kolommen

    EPB/Vaillant schema kenmerken:
    - Technisch schema + (Line) Legend + Pumps/Valves/Sensors VR10 + Vaillant codes (VWL/VRC/VR/VP/VIH)
    """
    try:
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            # Pak eerste 2 pagina's voor analyse
            text = ""
            for page in pdf.pages[:2]:
                t = page.extract_text()
                if t:
                    text += "\n" + t

            if not text.strip():
                return PDFType.UNKNOWN

            text_l = text.lower()

            # --- NIEUW: Vaillant override ---
            vaillant_score = _vaillant_schema_score(text_l)

            # EPB indicators (houd je bestaande lijst, maar voeg engels toe)
            epb_indicators = ["legende", "legend", "line legend", "installatievoorstel", "epb", "energieprestatie"]
            epb_score = sum(1 for indicator in epb_indicators if indicator in text_l)

            # FACQ indicators
            facq_indicators = ["facq", "artikelnr", "eenheidsprijs", "btw"]
            facq_score = sum(1 for indicator in facq_indicators if indicator in text_l)

            # Check voor tabelvorm (veel getallen + prijzen)
            price_patterns = len(re.findall(r"\d+[.,]\d{2}", text_l))
            article_patterns = len(re.findall(r"\b\d{5,6}\b", text_l))

            print(
                f"DEBUG: vaillant_score={vaillant_score}, EPB score={epb_score}, "
                f"FACQ score={facq_score}, prices={price_patterns}, articles={article_patterns}"
            )

            # ✅ Beslissingslogica (Vaillant krijgt voorrang)
            # Drempel: 2 markers + wat indicatoren = meestal al >10
            if vaillant_score >= 10:
                return PDFType.EPB_VOORSTEL

            # Bestaande logica
            if epb_score >= 2:
                return PDFType.EPB_VOORSTEL
            elif facq_score >= 2 or (price_patterns > 10 and article_patterns > 5):
                return PDFType.FACQ_OFFERTE
            else:
                # Heuristiek: als veel nummers/prijzen → FACQ, anders EPB
                if price_patterns > 5:
                    return PDFType.FACQ_OFFERTE
                else:
                    return PDFType.EPB_VOORSTEL

    except Exception as e:
        print(f"Error detecting PDF type: {e}")
        return PDFType.UNKNOWN
