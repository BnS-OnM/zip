"""
EPB Installatievoorstel → Odoo Sales import
Gebaseerd op ChatGPT's legende-matching code
"""
import re
import unicodedata
from typing import List, Tuple, Dict, Optional, Any
from io import BytesIO
import logging

import pdfplumber
from openpyxl import Workbook

# Configure logging
logger = logging.getLogger(__name__)

# Default Belgian VAT rate for EPB items (can be overridden)
DEFAULT_EPB_TAX_PERCENT = 21

def normalize_text(s: str) -> str:
    """Normaliseer tekst voor betere matching."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s.lower()

def parse_legend_blocks(text: str) -> List[str]:
    """
    Zoek sectie 'Legend' of 'Legende' en pak alleen echte items.
    Filter uit: indicator-only regels (1a, 2d, etc.) en headings.
    """
    txt = text
    m = re.search(r"(^|\n)\s*(legend|legende)\s*[:\n]", txt, flags=re.IGNORECASE)
    if not m:
        return []
    
    start = m.end()
    segment = txt[start:]
    
    # Stop bij "Line Legend" of volgende sectie
    stop_match = re.search(
        r"\n\s*(line legend|bijlage|appendix|notes?|opmerkingen|specificaties|schema|totaal|subtotaal)\s*\n",
        segment,
        flags=re.IGNORECASE
    )
    if stop_match:
        segment = segment[:stop_match.start()]
    
    # Splitsen en normaliseren
    lines = [normalize_text(l) for l in segment.splitlines()]
    lines = [l for l in lines if l and not l.startswith("pagina ") and len(l) >= 2]
    
    # Headings to filter out (case-insensitive)
    headings_to_filter = {
        "heat generators", "storages", "controls", "hydraulical units",
        "pumps", "functional valves", "safety units", "further armatures",
        "sensors vr10", "line legend", "potable water", "heating flow",
        "heating return", "hydraulic units", "sensors", "armatures"
    }
    
    # Filter items
    filtered_items = []
    for l in lines:
        # Skip indicator-only lines (1a, 2d, 3b, etc.)
        if re.match(r"^\d+[a-z]+$", l):
            continue
        
        # Skip heading lines
        if l in headings_to_filter:
            continue
        
        # Keep items with actual content
        if re.search(r"[a-z0-9]{2,}", l):
            # Remove bullet points and numbering
            l2 = re.sub(r"^(•|\-|\*|\d+[a-z]*[\)\.\-\s])+", "", l).strip()
            if l2 and len(l2) > 2:
                filtered_items.append(l2)
    
    # Deduplicatie
    seen = set()
    unique_items = []
    for l in filtered_items:
        if l not in seen:
            seen.add(l)
            unique_items.append(l)
    
    return unique_items


def extract_main_components(full_text: str) -> List[str]:
    """
    Zoek hoofdcomponenten in volledige PDF-tekst.
    
    Zoekt naar bekende Vaillant componenten zoals:
    - aroTHERM Split plus VWL 8.2 AS
    - Hydraulic module VWL 8.2 IS
    - uniSTOR VIH RW
    - VP RW 45/2 B
    - VRC720, VR71, VR940
    
    Returns:
        List van gevonden componenten (originele tekst)
    """
    components = []
    txt_lower = full_text.lower()
    
    # Component patterns with variations
    patterns = [
        # aroTHERM variants
        (r"arotherm\s+split\s+plus\s+vwl\s*8\.?2\s*as", "aroTHERM Split plus VWL 8.2 AS"),
        (r"arotherm.*?vwl\s*8\.?2\s*as", "aroTHERM VWL 8.2 AS"),
        (r"vwl\s*8\.?2\s*as", "VWL 8.2 AS"),
        
        # Hydraulic module variants
        (r"hydraulic\s+module\s+vwl\s*8\.?2\s*is", "Hydraulic module VWL 8.2 IS"),
        (r"vwl\s*8\.?2\s*is", "VWL 8.2 IS"),
        
        # uniSTOR variants
        (r"unistor\s+vih\s*rw", "uniSTOR VIH RW"),
        (r"vih\s*rw", "VIH RW"),
        
        # VP RW variants
        (r"vp\s*rw\s*45\s*/?\s*2\s*b", "VP RW 45/2 B"),
        
        # Controls
        (r"vrc\s*720", "VRC720"),
        (r"vr\s*71", "VR71"),
        (r"vr\s*940", "VR940"),
    ]
    
    seen = set()
    for pattern, name in patterns:
        if re.search(pattern, txt_lower):
            if name.lower() not in seen:
                seen.add(name.lower())
                components.append(name)
    
    logger.info(f"Extracted {len(components)} main components: {components}")
    return components

def extract_qty(item: str) -> Tuple[str, int]:
    """Haal aantal uit een item-regel. Retourneert (clean_text, qty)."""
    qty = 1
    s = item
    patterns = [
        r"\b(\d+)\s*[x×]\b",
        r"\b[x×]\s*(\d+)\b",
        r"\bqty\s*(\d+)\b",
        r"\((\d+)\s*(st|pcs|stuks?)\)",
        r"[-–]\s*(\d+)\b$",
        r"\b(\d+)\s*(st|pcs|stuks?)\b",
    ]
    for pat in patterns:
        m = re.search(pat, s, flags=re.IGNORECASE)
        if m:
            try:
                qty = int(m.group(1))
                s = (s[:m.start()] + s[m.end():]).strip(' -–,;')
                break
            except Exception:
                pass
    return s.strip(), max(1, qty)


def _norm(s: str) -> str:
    """Normalize for matching: lowercase, non-alnum to space, collapse whitespace."""
    if not s:
        return ""
    s = str(s).lower()
    s = re.sub(r'[^a-z0-9]+', ' ', s)
    return ' '.join(s.split())


def _collapse(s: str) -> str:
    """Remove all whitespace, dots, dashes, underscores, slashes."""
    if not s:
        return ""
    return re.sub(r'[\s\.\-_/]+', '', str(s))


def match_item(label: str, catalog_rows: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """
    Match item label to product catalog using keywords/synonyms.
    
    Strategy:
    1. Try exact match on _search or _search_collapse
    2. Try keyword/synonym matching
    3. Fallback to token overlap scoring
    
    Args:
        label: Item description from PDF
        catalog_rows: List of catalog product dicts
        
    Returns:
        Matched catalog dict or None
    """
    if not catalog_rows:
        return None
    
    label_norm = _norm(label)
    label_collapse = _collapse(label_norm)
    
    # Keywords and synonyms for strong matching
    KEYWORDS = {
        # Main components
        "arotherm split plus vwl 8.2 as": ["vwl 8.2 as", "vwl8.2as", "arotherm split plus", "arotherm"],
        "vwl 8.2 as": ["vwl8.2as", "vwl 8 2 as"],
        "hydraulic module vwl 8.2 is": ["vwl 8.2 is", "vwl8.2is", "hydraulic module"],
        "vwl 8.2 is": ["vwl8.2is", "vwl 8 2 is"],
        "unistor vih rw": ["vih rw", "vihrw", "unistor"],
        "vih rw": ["vihrw", "vih"],
        "vp rw 45/2 b": ["vp rw 45/2 b", "vprw45/2b", "vp rw 45 2 b", "vprw"],
        "vrc720": ["vrc720", "vrc 720"],
        "vr71": ["vr71", "vr 71"],
        "vr940": ["vr940", "vr 940"],
        
        # Legend keywords (NL/EN synonyms)
        "pump": ["pomp", "pump", "circulation"],
        "pomp": ["pump", "circulation"],
        "circulation pump": ["circulatiepomp", "pomp"],
        "valve": ["klep", "valve", "mixing"],
        "klep": ["valve"],
        "mixing valve": ["mengklep", "3 port", "3-port"],
        "vessel": ["vat", "vessel", "expansion"],
        "vat": ["vessel", "expansion"],
        "expansion vessel": ["expansievat", "expansie vat"],
        "safety": ["veiligheid", "safety", "assembly"],
        "veiligheid": ["safety"],
        "safety assembly": ["veiligheidsgroep", "safety group"],
        "keerklep": ["non-return", "check valve", "non return"],
        "non-return": ["keerklep", "check valve"],
        "vr10": ["vr10", "vr 10", "sensor"],
        "sensor": ["sensor", "vr10"],
    }
    
    # Build search tokens for label
    label_tokens = set(label_norm.split())
    
    # Try exact match on collapsed search
    for cat_item in catalog_rows:
        if label_collapse and cat_item.get("_search_collapse"):
            if label_collapse in cat_item["_search_collapse"]:
                logger.info(f"Exact collapse match: '{label}' -> '{cat_item['name']}'")
                return cat_item
    
    # Try keyword/synonym matching
    best_match = None
    best_score = 0
    
    for cat_item in catalog_rows:
        cat_search = cat_item.get("_search", "")
        cat_tokens = set(cat_search.split())
        
        # Calculate overlap score
        overlap = len(label_tokens & cat_tokens)
        
        # Boost score for keyword matches
        for keyword, synonyms in KEYWORDS.items():
            keyword_tokens = set(keyword.split())
            # Check if keyword appears in label
            if keyword_tokens.issubset(label_tokens):
                # Check if any synonym appears in catalog item
                for syn in synonyms:
                    syn_tokens = set(_norm(syn).split())
                    if syn_tokens.issubset(cat_tokens):
                        overlap += 5  # Strong boost for keyword match
                        break
            # Check if any synonym appears in label
            for syn in synonyms:
                syn_tokens = set(_norm(syn).split())
                if syn_tokens.issubset(label_tokens):
                    # Check if keyword or other synonyms appear in catalog
                    if keyword_tokens.issubset(cat_tokens):
                        overlap += 5
                        break
                    for syn2 in synonyms:
                        syn2_tokens = set(_norm(syn2).split())
                        if syn2_tokens.issubset(cat_tokens):
                            overlap += 3
                            break
        
        if overlap > best_score:
            best_score = overlap
            best_match = cat_item
    
    # Only return match if score is meaningful
    if best_score >= 2:
        logger.info(f"Token match (score={best_score}): '{label}' -> '{best_match['name']}'")
        return best_match
    
    logger.warning(f"No match found for: '{label}'")
    return None

def epb_pdf_to_xlsx(pdf_bytes: bytes) -> BytesIO:
    """
    Converteer EPB/installatievoorstel PDF naar XLSX met legende items.
    """
    xlsx_file, _ = epb_pdf_to_xlsx_and_data(pdf_bytes)
    return xlsx_file


def epb_pdf_to_xlsx_and_data(pdf_bytes: bytes) -> Tuple[BytesIO, List[Dict]]:
    """
    Converteer EPB/installatievoorstel PDF naar XLSX met legende items en gestructureerde data.
    Output: Odoo Sales format met 6 kolommen.
    
    Returns:
        Tuple[BytesIO, List[Dict]]: XLSX file en lijst van orderregels
    """
    # Import catalog
    from app.xlsx_import import get_product_catalog
    catalog = get_product_catalog()
    
    # Extract text
    full_text = []
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            full_text.append(t)
    
    full_text_str = "\n".join(full_text)
    
    # Extract legend items
    legend_items_raw: List[str] = []
    for chunk in full_text:
        legend_items_raw.extend(parse_legend_blocks(chunk))
    
    # Extract main components
    main_components = extract_main_components(full_text_str)
    
    # Combine and dedupe
    all_items_raw = main_components + legend_items_raw
    
    # Dedupe while preserving order
    seen_lower = set()
    unique_items_raw = []
    for item in all_items_raw:
        item_lower = item.lower()
        if item_lower not in seen_lower:
            seen_lower.add(item_lower)
            unique_items_raw.append(item)
    
    # Fallback: if no items found, use all lines
    if not unique_items_raw:
        logger.warning("No legend or components found, using all lines as fallback")
        for t in full_text:
            lines = [normalize_text(l) for l in t.splitlines() if l.strip()]
            unique_items_raw.extend(lines)
        # Dedupe again
        seen_lower = set()
        deduped = []
        for item in unique_items_raw:
            item_lower = item.lower()
            if item_lower not in seen_lower:
                seen_lower.add(item_lower)
                deduped.append(item)
        unique_items_raw = deduped
    
    # Extract quantities
    items_with_qty: List[Tuple[str, int]] = [extract_qty(it) for it in unique_items_raw]
    
    logger.info(f"Found {len(items_with_qty)} total items (legend + main components)")
    
    # Match to catalog and build rows
    matched_products = {}  # Dict to dedupe on product name
    unmatched_count = 0
    
    for description, qty in items_with_qty:
        if not description or len(description) < 2:
            continue
        
        # Try to match
        matched = match_item(description, catalog) if catalog else None
        
        if matched:
            product_name = matched["name"]
            # Dedupe: if product already in output, add quantities
            if product_name in matched_products:
                matched_products[product_name]["quantity"] += qty
            else:
                matched_products[product_name] = {
                    "product": product_name,
                    "description": matched.get("description", product_name),
                    "quantity": qty,
                    "price": matched.get("list_price", 0.0),
                    "code": matched.get("default_code", "")
                }
        else:
            # No match - use original description
            unmatched_count += 1
            # Use description as key to dedupe unmatched items
            if description not in matched_products:
                matched_products[description] = {
                    "product": description,
                    "description": description,
                    "quantity": qty,
                    "price": 0.0,
                    "code": ""
                }
            else:
                matched_products[description]["quantity"] += qty
    
    logger.info(f"Matched: {len(matched_products) - unmatched_count}, Unmatched: {unmatched_count}")
    
    # Build final rows
    final_rows = list(matched_products.values())
    
    # Build XLSX with Odoo Sales format (6 columns)
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    
    # Headers
    ws.append([
        "Orderreferentie",
        "Partner ID",
        "Product",
        "Product omschrijving",
        "Aantal",
        "Prijs"
    ])
    
    # Data rows
    for idx, row in enumerate(final_rows):
        ws.append([
            "GPT-001" if idx == 0 else "",  # Order reference only on first row
            "" if idx == 0 else "",  # Partner ID only on first row (empty for now)
            row["product"],
            row["description"],
            row["quantity"],
            row["price"]
        ])
    
    # Prepare structured data for Odoo import
    lines_data = []
    for row in final_rows:
        lines_data.append({
            "product_code": row["code"],
            "description": row["product"],  # Use product name as description
            "quantity": row["quantity"],
            "unit_price": row["price"],
            "tax_percent": DEFAULT_EPB_TAX_PERCENT
        })
    
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    logger.info(f"EPB XLSX generated with {len(final_rows)} unique products (after deduplication)")
    return output, lines_data
