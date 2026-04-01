"""
Comprehensive VER Data Extractor
Extracts ALL 20 sections from a VER PDF into a flat dictionary suitable for Excel export.
Uses native text extraction (PyMuPDF) for text PDFs, Tesseract OCR for scanned PDFs.
Supports: English, Hindi, Odia, Tamil, Telugu, Kannada, Marathi, Gujarati.

Reference format: VER_Tsupfume_2025-26.pdf (English, Nagaland)
"""
import re
import json
import fitz  # PyMuPDF
from pathlib import Path
from collections import OrderedDict

# ── Language configuration ────────────────────────────────────

SUPPORTED_LANGUAGES = {
    "English": {"tesseract": "eng", "script": "Latin"},
    "Hindi": {"tesseract": "hin+eng", "script": "Devanagari"},
    "Odia": {"tesseract": "ori+hin+eng", "script": "Odia"},
    "Tamil": {"tesseract": "tam+eng", "script": "Tamil"},
    "Telugu": {"tesseract": "tel+eng", "script": "Telugu"},
    "Kannada": {"tesseract": "kan+eng", "script": "Kannada"},
    "Marathi": {"tesseract": "mar+eng", "script": "Devanagari"},
    "Gujarati": {"tesseract": "guj+eng", "script": "Gujarati"},
    "Auto-detect": {"tesseract": "eng", "script": "Latin"},
}

# State → language mapping for auto-detection
STATE_LANGUAGE_MAP = {
    "odisha": "Odia", "orissa": "Odia",
    "tamil nadu": "Tamil", "tamilnadu": "Tamil",
    "andhra pradesh": "Telugu", "telangana": "Telugu",
    "karnataka": "Kannada",
    "maharashtra": "Marathi",
    "gujarat": "Gujarat",
    "chhattisgarh": "Hindi", "madhya pradesh": "Hindi", "rajasthan": "Hindi",
    "uttar pradesh": "Hindi", "bihar": "Hindi", "jharkhand": "Hindi",
    "nagaland": "English", "meghalaya": "English", "mizoram": "English",
    "manipur": "English", "assam": "English", "arunachal pradesh": "English",
    "sikkim": "English", "tripura": "English",
}


# ── Master field definitions ─────────────────────────────────
# Every field that can appear in the output Excel.
# Grouped by VER section. Order matters for the spreadsheet.

MASTER_FIELDS = OrderedDict([
    # ── Identifiers ──
    ("village_name", ""),
    ("state", ""),
    ("district", ""),
    ("block", ""),
    ("gram_panchayat", ""),
    ("latitude", ""),
    ("longitude", ""),
    ("date_of_survey", ""),
    ("ver_year", ""),
    ("total_pages", 0),
    ("extraction_method", ""),

    # ── S2: General Info ──
    ("total_area_ha", ""),
    ("forest_land_pct", ""),
    ("grazing_land_pct", ""),
    ("revenue_wasteland_pct", ""),
    ("community_conserved_area_pct", ""),
    ("agricultural_land_pct", ""),
    ("other_land_pct", ""),
    ("other_land_details", ""),
    ("total_population", ""),
    ("total_households", ""),
    ("gen_hh", ""),
    ("sc_hh", ""),
    ("st_hh", ""),
    ("obc_hh", ""),
    ("large_farmers_gt10ha", ""),
    ("medium_farmers_4_10ha", ""),
    ("semi_medium_2_4ha", ""),
    ("small_farmers_1_2ha", ""),
    ("marginal_farmers_lt1ha", ""),
    ("landless_farmers", ""),
    ("landholding_remarks", ""),
    ("major_livelihoods", ""),

    # ── S3: Village History ──
    ("village_history_narrative", ""),
    ("traditional_songs", ""),
    ("myths_and_beliefs", ""),

    # ── S4: Agro-ecological ──
    ("kharif_crops", ""),
    ("rabi_crops", ""),
    ("zaid_crops", ""),
    ("traditional_crop_varieties", ""),
    ("farming_practices", ""),
    ("hedge_biodiversity", ""),
    ("soil_type", ""),
    ("soil_fertility_change", ""),
    ("soil_fertility_reason", ""),
    ("soil_health_indicators", ""),
    ("traditional_climate_practices", ""),
    ("pest_incidences", ""),
    ("major_weeds", ""),

    # ── S5: Livestock ──
    ("livestock_summary", ""),  # "Pig:154, Poultry:2580, Rabbit:238"
    ("livestock_detailed", ""),  # Full table as text
    ("indigenous_breeds", ""),
    ("livestock_diseases", ""),
    ("ethno_veterinary_practices", ""),
    ("traditional_livestock_practices", ""),

    # ── S6: Waterscape ──
    ("drinking_water_sources", ""),  # "Tap:288, Spring:2, River:4"
    ("livestock_water_sources", ""),
    ("irrigation_sources", ""),
    ("water_quality_changes", ""),
    ("traditional_water_conservation", ""),
    ("important_water_bodies", ""),

    # ── S7: Forest Lands ──
    ("forest_name", ""),
    ("forest_type", ""),
    ("forest_location_geocode", ""),
    ("forest_size_ha", ""),
    ("forest_tree_species", ""),
    ("forest_shrub_species", ""),
    ("forest_herb_species", ""),
    ("forest_grass_species", ""),
    ("forest_climber_species", ""),
    ("forest_ntfp", ""),

    # ── S8: Grassland ──
    ("grassland_name", ""),
    ("grassland_location", ""),
    ("grassland_size_ha", ""),
    ("grassland_species", ""),
    ("grassland_ntfp", ""),

    # ── S9: Wasteland ──
    ("wasteland_name", ""),
    ("wasteland_location", ""),
    ("wasteland_size_ha", ""),
    ("wasteland_species", ""),

    # ── S10: Sacred Groves ──
    ("sacred_groves", ""),

    # ── S11: Ecologically Important Sites ──
    ("ecological_sites", ""),

    # ── S12: Giant Trees ──
    ("giant_trees", ""),  # "Name|Geocode|Landmark; ..."

    # ── S13: Bee Hives ──
    ("bee_hives", ""),

    # ── S14: Fire Incidence ──
    ("fire_incidence", ""),

    # ── S15: Conservation Ethos ──
    ("conservation_ethos", ""),
    ("bamboo_species", ""),

    # ── S16: Medicinal Plants ──
    ("medicinal_plants", ""),  # "LocalName|ScientificName|Use|Trend; ..."

    # ── S17: Invasive Plants ──
    ("invasive_plants", ""),

    # ── S18: Feral Animals ──
    ("feral_animals", ""),

    # ── S19: Protected Species ──
    ("protected_species", ""),

    # ── S20: Flora & Fauna ──
    ("tree_diversity_count", 0),
    ("tree_diversity", ""),
    ("shrub_diversity_count", 0),
    ("shrub_diversity", ""),
    ("herb_grass_diversity_count", 0),
    ("herb_grass_diversity", ""),
    ("lower_plant_count", 0),
    ("lower_plant_diversity", ""),
    ("mammal_count", 0),
    ("mammal_diversity", ""),
    ("bird_count", 0),
    ("bird_diversity", ""),
    ("reptile_amphibian_count", 0),
    ("reptile_amphibian_diversity", ""),
    ("butterfly_count", 0),
    ("butterfly_diversity", ""),
    ("dragonfly_count", 0),
    ("dragonfly_diversity", ""),
    ("fish_insect_other_count", 0),
    ("fish_insect_other_diversity", ""),
    ("soil_macrofauna_count", 0),
    ("soil_macrofauna_diversity", ""),
    ("total_species_count", 0),
])


def get_empty_record():
    """Return a fresh empty record with all master fields."""
    return OrderedDict((k, v if isinstance(v, int) else "") for k, v in MASTER_FIELDS.items())


# ── PDF text extraction ──────────────────────────────────────

def extract_text_from_pdf(pdf_path: str, language: str = "Auto-detect",
                          progress_callback=None) -> tuple[list[str], bool]:
    """Extract text from PDF. Returns (list_of_page_texts, is_native_text).
    Auto-detects whether PDF has native text or is scanned.
    Falls back to Tesseract OCR for scanned PDFs.
    """
    doc = fitz.open(pdf_path)
    pages = []
    text_pages = 0

    # First pass: try native text extraction
    for i in range(doc.page_count):
        text = doc[i].get_text()
        pages.append(text)
        if len(text.strip()) > 50:
            text_pages += 1

    is_native = text_pages / len(pages) > 0.5 if pages else False

    if is_native:
        doc.close()
        return pages, True

    # Scanned PDF → use Tesseract OCR
    try:
        import pytesseract
        from PIL import Image
        import io
    except ImportError:
        doc.close()
        # Return whatever native text we got (partial)
        return pages, False

    # Determine Tesseract language string
    lang_config = SUPPORTED_LANGUAGES.get(language, SUPPORTED_LANGUAGES["English"])
    tess_lang = lang_config["tesseract"]

    # Try image preprocessing if OpenCV available
    try:
        import cv2
        import numpy as np
        has_cv2 = True
    except ImportError:
        has_cv2 = False

    pages = []  # Reset and OCR all pages
    total = doc.page_count

    for i in range(total):
        if progress_callback and i % 10 == 0:
            progress_callback(i, total, f"OCR page {i+1}/{total} ({tess_lang})...")

        page = doc[i]
        # Render at 300 DPI for better OCR
        pix = page.get_pixmap(dpi=300)
        img_bytes = pix.tobytes("png")
        img = Image.open(io.BytesIO(img_bytes))

        # Convert to grayscale
        if img.mode != "L":
            img = img.convert("L")

        # Apply enhancement if OpenCV available
        if has_cv2:
            img_array = np.array(img)
            denoised = cv2.fastNlMeansDenoising(img_array, None, h=12,
                                                 templateWindowSize=7, searchWindowSize=21)
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
            enhanced = clahe.apply(denoised)
            binary = cv2.adaptiveThreshold(enhanced, 255,
                                           cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                           cv2.THRESH_BINARY, blockSize=15, C=8)
            img = Image.fromarray(binary)

        # Run Tesseract
        try:
            text = pytesseract.image_to_string(img, lang=tess_lang,
                                                config='--oem 3 --psm 6')
        except Exception:
            text = ""

        pages.append(text)

    doc.close()
    return pages, False


def find_section_ranges(pages: list[str]) -> dict:
    """Detect VER section boundaries in the page texts.
    Returns dict: section_key -> (start_page_idx, end_page_idx)
    """
    # Multilingual section detection patterns
    # English + Odia + Hindi + Tamil + Telugu + Kannada + Marathi + Gujarati
    # Uses (?!\d) to prevent "Section - 2" matching "Section - 20"
    def _sp(num, eng_kw, local_kws=""):
        """Build section pattern: Section-N + English keyword + local script variants."""
        parts = [rf'Section\s*[-–]\s*{num}(?!\d)']
        if eng_kw:
            parts.append(eng_kw)
        # Odia numerals
        odia_nums = {"2": "୨(?!୦)", "3": "୩", "4": "୪", "5": "୫", "6": "୬", "7": "୭",
                     "8": "୮", "9": "୯", "10": "୧୦", "11": "୧୧", "12": "୧୨",
                     "13": "୧୩", "14": "୧୪", "15": "୧୫", "16": "୧୬", "17": "୧୭",
                     "18": "୧୮", "19": "୧୯", "20": "୨୦"}
        if num in odia_nums:
            parts.append(rf'ବିଭାଗ\s*[-–]\s*{odia_nums[num]}')
        # Hindi/Marathi (Devanagari numerals)
        hindi_nums = {"2": "२", "3": "३", "4": "४", "5": "५", "6": "६", "7": "७",
                      "8": "८", "9": "९", "10": "१०", "11": "११", "12": "१२",
                      "13": "१३", "14": "१४", "15": "१५", "16": "१६", "17": "१७",
                      "18": "१८", "19": "१९", "20": "२०"}
        if num in hindi_nums:
            parts.append(rf'(?:विभाग|భాగం|பகுதி|ವಿಭಾಗ)\s*[-–]\s*{hindi_nums[num]}')
        # Additional local keywords
        if local_kws:
            parts.append(local_kws)
        return re.compile("|".join(parts), re.I)

    section_patterns = [
        ("s2",  _sp("2",  r'General\s+Information\s+of\s+Village',
                    r'सामान्य जानकारी|ସାଧାରଣ ତଥ୍ୟ|సాధారణ సమాచారం|பொது தகவல்|ಸಾಮಾನ್ಯ ಮಾಹಿತಿ')),
        ("s3",  _sp("3",  r'3\.1\s+Documenting\s+Village|Village\s+History',
                    r'गाँव का इतिहास|ଗ୍ରାମ ଇତିହାସ|గ్రామ చరిత్ర|கிராம வரலாறு|ಹಳ್ಳಿ ಇತಿಹಾಸ')),
        ("s4",  _sp("4",  r'4\.1\.?\s+Cropping|Agro.?ecological\s+Knowledge',
                    r'कृषि-पारिस्थितिक|କୃଷି-ପରିବେଶ|వ్యవసాయ|வேளாண்|ಕೃಷಿ-ಪರಿಸರ')),
        ("s5",  _sp("5",  r'5\.1\s+Livestock',
                    r'पशुधन|ପଶୁ|పశువులు|கால்நடை|ಜಾನುವಾರು')),
        ("s6",  _sp("6",  r'6\.1\s+Availability.*Water|Waterscape',
                    r'जलक्षेत्र|ଜଳ|నీటి వనరులు|நீர்|ಜಲ')),
        ("s7",  _sp("7",  r'7\.1\s+General\s+info|Forest\s+Land',
                    r'वन भूमि|ଜଙ୍ଗଲ|అడవి భూమి|வன நிலங்கள்|ಅರಣ್ಯ')),
        ("s8",  _sp("8",  r'8\.1\s+General\s+info|Grassland|Grazing',
                    r'चरागाह|ଘାସ ଜମି|పచ్చిక|மேய்ச்சல்|ಹುಲ್ಲುಗಾವಲು')),
        ("s9",  _sp("9",  r'9\.1\s+General\s+info|Waste\s*[Ll]and|Revenue\s+Waste',
                    r'राजस्व बंजर|ରାଜସ୍ୱ ଜମି|రెవెన్యూ|கழிவு நிலம்|ಕಂದಾಯ')),
        ("s10", _sp("10", r'Grooves.*Sacred\s+groove|Sacred\s+[Gg]rove',
                    r'पवित्र उपवन|ପବିତ୍ର ବନ|పవిత్ర వనం|புனித தோப்பு|ದೇವರ ಕಾಡು')),
        ("s11", _sp("11", r'Ecologically\s+important\s+sites',
                    r'पारिस्थितिक')),
        ("s12", _sp("12", r'List\s+of\s+Old\s+and\s+Giant|Giant\s+tree',
                    r'पुराने और विशाल पेड़|ପୁରାତନ ଗଛ')),
        ("s13", _sp("13", r'Locations\s+of\s+big\s+bee|[Bb]ee\s+hive',
                    r'मधुमक्खी|ମହୁମାଛି|తేనెటీగ|தேனீ|ಜೇನುಗೂಡು')),
        ("s14", _sp("14", r'Fire\s+incidence\s+in',
                    r'आग|ନିଆଁ|అగ్ని|தீ|ಬೆಂಕಿ')),
        ("s15", _sp("15", r'15\.1\s+Bamboo|Local\s+Conservation\s+ethos',
                    r'संरक्षण|ସଂରକ୍ଷଣ|సంరక్షణ|பாதுகாப்பு|ಸಂರಕ್ಷಣೆ')),
        ("s16", _sp("16", r'Medicinal\s+plants.*uses',
                    r'औषधीय|ଔଷଧୀୟ|ఔషధ|மருத்துவ|ಔಷಧ')),
        ("s17", _sp("17", r'Invasive\s+plants\s+in',
                    r'आक्रामक|ଆକ୍ରମଣକାରୀ|దాడి చేసే|ஊடுருவும்|ಆಕ್ರಮಣಕಾರಿ')),
        ("s18", _sp("18", r'Feral\s+animal\s+in',
                    r'जंगली जानवर|ବଣ୍ୟ ପ୍ରାଣୀ|వన్య జంతువులు|காட்டு விலங்கு|ಕಾಡು ಪ್ರಾಣಿ')),
        ("s19", _sp("19", r'socially.*culturally\s+protected',
                    r'सांस्कृतिक|ସାଂସ୍କୃତିକ|సాంస్కృతిక|பண்பாட்டு|ಸಾಂಸ್ಕೃತಿಕ')),
        ("s20", _sp("20", r'20\.1\s+Tree\s+diversity|List\s+of\s+flora',
                    r'वनस्पति और जीव|ଉଦ୍ଭିଦ ଓ ଜୀବଜନ୍ତୁ|వృక్షజాతి|தாவரங்கள்|ಸಸ್ಯ ಮತ್ತು ಪ್ರಾಣಿ')),
    ]

    # Detect TOC pages: pages where 3+ section headers appear
    toc_pages = set()
    for i, text in enumerate(pages):
        matches = sum(1 for _, pat in section_patterns if pat.search(text))
        if matches >= 3:
            toc_pages.add(i)

    sections = {}
    for key, pat in section_patterns:
        for i, text in enumerate(pages):
            if i in toc_pages:
                continue
            if pat.search(text) and key not in sections:
                sections[key] = {"start": i}
                break

    # Compute end pages
    sorted_keys = sorted(sections.keys(), key=lambda k: sections[k]["start"])
    for i, key in enumerate(sorted_keys):
        if i + 1 < len(sorted_keys):
            sections[key]["end"] = sections[sorted_keys[i + 1]]["start"] - 1
        else:
            sections[key]["end"] = len(pages) - 1

    return sections


def get_section_text(pages: list[str], sections: dict, key: str) -> str:
    """Get concatenated text for a section."""
    if key not in sections:
        return ""
    s = sections[key]
    return "\n".join(pages[s["start"]:s["end"] + 1])


# ── Section parsers ──────────────────────────────────────────

def _clean(text: str) -> str:
    """Clean extracted text: collapse whitespace, strip."""
    return re.sub(r'\s+', ' ', text).strip()


def _extract_between(text: str, start_pat: str, end_pat: str) -> str:
    """Extract text between two patterns."""
    m = re.search(rf'{start_pat}(.*?){end_pat}', text, re.DOTALL | re.I)
    return _clean(m.group(1)) if m else ""


def parse_s2(text: str, record: dict):
    """Parse Section 2: General Information."""
    # Village name
    m = re.search(r'Village\s+Name:\s*(.+?)(?:\n|Village)', text, re.I)
    if m:
        record["village_name"] = _clean(m.group(1))

    # State
    m = re.search(r'State:\s*(.+?)(?:\n|Village)', text, re.I)
    if m:
        record["state"] = _clean(m.group(1).split('Village')[0])

    # Block
    m = re.search(r'Block:\s*(.+?)(?:\n|$)', text, re.I)
    if m:
        record["block"] = _clean(m.group(1))

    # Date
    m = re.search(r'Date:\s*(.+?)(?:\n|$)', text, re.I)
    if m:
        record["date_of_survey"] = _clean(m.group(1))

    # GPS — DMS format: N25°32'58.9" E94°19'16.3" (may span lines, smart quotes)
    dms = re.search(
        r'[NS]?\s*(\d{1,2})[°]\s*(\d{1,2})[\'′\u2018\u2019]\s*(\d{1,2}(?:\.\d+)?)[\"″\u201C\u201D]'
        r'.*?[EW]?\s*(\d{1,3})[°]\s*(\d{1,2})[\'′\u2018\u2019]\s*(\d{1,2}(?:\.\d+)?)[\"″\u201C\u201D]',
        text, re.DOTALL)
    if dms:
        lat = float(dms.group(1)) + float(dms.group(2)) / 60 + float(dms.group(3)) / 3600
        lon = float(dms.group(4)) + float(dms.group(5)) / 60 + float(dms.group(6)) / 3600
        record["latitude"] = round(lat, 6)
        record["longitude"] = round(lon, 6)

    # GPS — Decimal format
    if not record["latitude"]:
        coords = re.findall(r'(\d{1,3}\.\d{4,8})', text)
        if len(coords) >= 2:
            lat, lon = float(coords[0]), float(coords[1])
            if 6 <= lat <= 38 and 68 <= lon <= 98:
                record["latitude"] = lat
                record["longitude"] = lon

    # Total area — handles both percentage format and hectare format
    m = re.search(r'(\d+(?:\.\d+)?)\s*ha\s*\n?\s*\(?\s*100%', text, re.I)
    if m:
        record["total_area_ha"] = m.group(1)
        after = text[m.end():]
        pcts = re.findall(r'(\d+)%|([Nn]il)', after[:200])
        pct_vals = [int(n) if n else 0 for n, nil in pcts]

        # Detect column order: official format has "Revenue Wasteland" 3rd,
        # Nagaland/NE format has "Community Conserved Area" 3rd
        has_cca = bool(re.search(r'Community\s+Conserved', text, re.I))
        has_rwl = bool(re.search(r'Revenue\s+Waste', text, re.I))

        if has_cca:
            fields = ["forest_land_pct", "grazing_land_pct", "community_conserved_area_pct",
                      "agricultural_land_pct", "other_land_pct"]
        else:
            fields = ["forest_land_pct", "grazing_land_pct", "revenue_wasteland_pct",
                      "agricultural_land_pct", "other_land_pct"]

        for i, f in enumerate(fields):
            if i < len(pct_vals):
                record[f] = str(pct_vals[i]) + "%"

    # Also try hectare-based format (official VER: "Forest Land (in ha)")
    if not record["total_area_ha"]:
        m = re.search(r'Total\s+Village\s+Area.*?(\d+(?:\.\d+)?)', text, re.I | re.DOTALL)
        if m:
            record["total_area_ha"] = m.group(1)

    # Other land details
    m = re.search(r'[Ii]f\s+others.*?specific:\s*(.+?)(?:\n\n|\n\d)', text, re.DOTALL)
    if m:
        record["other_land_details"] = _clean(m.group(1))

    # Population
    m = re.search(r'Total\s+Population:\s*([\d,]+)', text, re.I)
    if m:
        record["total_population"] = m.group(1).replace(',', '')

    # Households
    m = re.search(r'households\)?\s*(\d+)', text, re.I)
    if m:
        record["total_households"] = m.group(1)

    # Caste composition
    caste_block = re.search(r'Caste\s+composition.*?Total\s+Population', text, re.DOTALL | re.I)
    if caste_block:
        cb = caste_block.group(0)
        for label, field in [("General", "gen_hh"), (r"Scheduled\s+Caste", "sc_hh"),
                             (r"Scheduled\s+Tribe", "st_hh"), (r"Other\s+Backward", "obc_hh")]:
            m2 = re.search(rf'{label}.*?\n.*?(\d+|N/A|Nil)', cb, re.I | re.DOTALL)
            if m2:
                record[field] = m2.group(1)

    # Landholding
    lh = re.search(r'Landholding.*?Number\s+of\s*\n?\s*Households\s*\n(.*?)(?=Remarks|2\.\d|$)',
                   text, re.I | re.DOTALL)
    if lh:
        vals = re.findall(r'(\d+|Nil)', lh.group(1))
        fields = ["large_farmers_gt10ha", "medium_farmers_4_10ha", "semi_medium_2_4ha",
                  "small_farmers_1_2ha", "marginal_farmers_lt1ha", "landless_farmers"]
        for i, f in enumerate(fields):
            if i < len(vals):
                record[f] = "0" if vals[i] == "Nil" else vals[i]

    # Remarks
    m = re.search(r'Remarks:\s*(.+?)(?:\n\n|2\.\d|$)', text, re.DOTALL)
    if m:
        record["landholding_remarks"] = _clean(m.group(1))

    # Livelihoods
    lv = re.search(r'Major\s+Livelihood.*?(?=Section|$)', text, re.I | re.DOTALL)
    if lv:
        items = re.findall(r'\([a-z]\)\s*([A-Za-z][\w\s.]+?)(?=\s*\([a-z]\)|$)', lv.group(0))
        record["major_livelihoods"] = "; ".join(i.strip() for i in items if len(i.strip()) > 2)


def parse_s3(text: str, record: dict):
    """Parse Section 3: Village History."""
    # History narrative (3.1)
    s31 = re.search(r'3\.1\s+.*?(?:History|ଇତିହାସ).*?(?=3\.2|\Z)', text, re.DOTALL | re.I)
    if s31:
        record["village_history_narrative"] = _clean(s31.group(0))[:2000]

    # Myths & beliefs (3.2)
    s32 = re.search(r'3\.2\s+.*?(?:Myths|Beliefs|ବିଶ୍ବାସ).*?(?=3\.3|\Z)', text, re.DOTALL | re.I)
    if s32:
        record["myths_and_beliefs"] = _clean(s32.group(0))[:2000]

    # Traditional songs within 3.2
    songs = re.findall(r'Song\s*\n(.+?)(?=Song\s*\n|\Z)', text, re.DOTALL | re.I)
    if songs:
        record["traditional_songs"] = "; ".join(_clean(s)[:200] for s in songs[:5])


def parse_s4(text: str, record: dict):
    """Parse Section 4: Agro-ecological."""
    # Cropping pattern
    for season, field in [("Kharif", "kharif_crops"), ("Rabi", "rabi_crops"), ("Zaid", "zaid_crops")]:
        m = re.search(rf'{season}\s+Season\s*\n(.*?)(?=\b(?:Rabi|Zaid|Kharif)\s+Season|4\.\d|Section|\Z)',
                      text, re.DOTALL | re.I)
        if m:
            crops = re.findall(r'^([A-Z][a-z]+(?:\s+[a-z]+)?)', m.group(1), re.MULTILINE)
            record[field] = "; ".join(c for c in crops if len(c) > 2 and c not in ("Rice", "Season"))
            if not record[field] and crops:
                record[field] = "; ".join(crops[:15])

    # Traditional varieties
    tv_block = re.search(r'4\.2\s+Traditional.*?(?=4\.\d|Section|$)', text, re.DOTALL | re.I)
    if tv_block:
        varieties = re.findall(r'^([A-Z][\w\s\']+?)(?=\n(?:Abundant|Scarce|Rare|$))', tv_block.group(0), re.MULTILINE)
        record["traditional_crop_varieties"] = "; ".join(_clean(v) for v in varieties)

    # Farming practices
    fp = re.search(r'(?:4\.2|4\.3)\s+Farming\s+practices.*?(?=4\.\d|Section|$)', text, re.DOTALL | re.I)
    if fp:
        practices = re.findall(r'^([A-Z][\w\s()]+?)(?=\n)', fp.group(0), re.MULTILINE)
        record["farming_practices"] = "; ".join(_clean(p) for p in practices if len(p.strip()) > 3)[:1500]

    # Hedge biodiversity
    hb = re.search(r'(?:4\.3|4\.4)\s+Hedge.*?(?=4\.\d|Section|$)', text, re.DOTALL | re.I)
    if hb:
        record["hedge_biodiversity"] = _clean(hb.group(0))[:1500]

    # Soil health
    m = re.search(r'Soil\s+types?\s*\n?\s*(.+?)(?:\n)', text, re.I)
    if m:
        record["soil_type"] = _clean(m.group(1))
    m = re.search(r'Soil\s+fertility.*?changed.*?\n\s*(.+?)(?:\n)', text, re.I | re.DOTALL)
    if m:
        record["soil_fertility_change"] = _clean(m.group(1))
    m = re.search(r'Reasons?\s+for\s+changes?\s+in\s+soil\s+fertility\s*\n\s*(.+?)(?:\n)', text, re.I)
    if m:
        record["soil_fertility_reason"] = _clean(m.group(1))

    # Soil indicators
    si = re.search(r'(?:4\.5|Soil\s+Health\s+Indicators).*?(?=4\.\d|Section|$)', text, re.DOTALL | re.I)
    if si:
        record["soil_health_indicators"] = _clean(si.group(0))[:1500]

    # Traditional climate practices
    tp = re.search(r'(?:4\.6|Traditional\s+Practices.*?[Cc]limate).*?(?=4\.\d|Section|$)', text, re.DOTALL | re.I)
    if tp:
        record["traditional_climate_practices"] = _clean(tp.group(0))[:2000]

    # Pest incidences
    pi = re.search(r'(?:4\.7|4\.8|Agriculture\s+Pest).*?(?=4\.\d|Section|$)', text, re.DOTALL | re.I)
    if pi:
        record["pest_incidences"] = _clean(pi.group(0))[:1500]

    # Weeds
    wd = re.search(r'(?:4\.8|4\.9|Major\s+weeds).*?(?=Section|$)', text, re.DOTALL | re.I)
    if wd:
        record["major_weeds"] = _clean(wd.group(0))[:1500]


def parse_s5(text: str, record: dict):
    """Parse Section 5: Livestock."""
    # Livestock numbers table
    livestock_types = ["Cows (indigenous)", "Cow (Hybrids)", "Oxen", "Pig", "Buffaloes",
                       "Goat", "Sheep", "Poultry", "Duckery", "Mithun", "Rabbit",
                       "Horse", "Donkey", "Camel"]
    summary_parts = []
    detailed_parts = []

    s51 = re.search(r'5\.1\s+Livestock\s+number.*?(?=5\.2|Section|$)', text, re.I | re.DOTALL)
    if s51:
        s51_text = s51.group(0)
        for lt in livestock_types:
            short = lt.split("(")[0].strip()
            m = re.search(rf'{re.escape(short)}\s*\n?\s*(?:\([^)]*\)\s*\n?)?\s*(\d+|Nil)', s51_text, re.I)
            if m:
                val = 0 if m.group(1).lower() == 'nil' else int(m.group(1))
                # Get trends
                after = s51_text[m.end():m.end() + 150]
                trends = re.findall(r'[↑↓↔]', after[:80])
                trend_str = "/".join(trends[:3]) if trends else ""

                if val > 0:
                    summary_parts.append(f"{short}:{val}")
                detailed_parts.append(f"{short}:{val} [{trend_str}]")

    record["livestock_summary"] = "; ".join(summary_parts)
    record["livestock_detailed"] = "; ".join(detailed_parts)

    # Indigenous breeds
    ib = re.search(r'5\.2\s+.*?[Ii]ndigenous\s+breeds.*?(?=5\.3|Section|$)', text, re.DOTALL | re.I)
    if ib:
        record["indigenous_breeds"] = _clean(ib.group(0))[:1500]

    # Diseases
    dis = re.search(r'5\.3\s+.*?[Dd]iseases.*?(?=5\.4|Section|$)', text, re.DOTALL | re.I)
    if dis:
        record["livestock_diseases"] = _clean(dis.group(0))[:1500]

    # Ethno-vet
    ev = re.search(r'5\.4\s+.*?[Ee]thno.*?(?=5\.5|Section|$)', text, re.DOTALL | re.I)
    if ev:
        record["ethno_veterinary_practices"] = _clean(ev.group(0))[:1500]

    # Traditional practices (5.5 if present)
    tp = re.search(r'5\.5\s+.*?(?=Section|$)', text, re.DOTALL | re.I)
    if tp:
        record["traditional_livestock_practices"] = _clean(tp.group(0))[:1500]


def parse_s6(text: str, record: dict):
    """Parse Section 6: Waterscape."""
    water_types = ["Tap", "Tube-well", "Tubewell", "Open Well", "Spring", "River", "Stream"]

    # 6.1 Drinking
    s61 = re.search(r'6\.1\s+.*?[Dd]rinking.*?(?=6\.2|$)', text, re.DOTALL | re.I)
    if s61:
        parts = []
        for wt in water_types:
            m = re.search(rf'{re.escape(wt)}\s*\n?\s*(\d+|Nil)', s61.group(0), re.I)
            if m:
                val = 0 if m.group(1).lower() == 'nil' else int(m.group(1))
                if val > 0:
                    parts.append(f"{wt}:{val}")
        record["drinking_water_sources"] = "; ".join(parts)

    # 6.2 Livestock water
    s62 = re.search(r'6\.2\s+.*?[Ll]ivestock.*?(?=6\.3|$)', text, re.DOTALL | re.I)
    if s62:
        parts = []
        for wt in water_types:
            m = re.search(rf'{re.escape(wt)}\s*\n?\s*(\d+|Nil)', s62.group(0), re.I)
            if m:
                val = 0 if m.group(1).lower() == 'nil' else int(m.group(1))
                if val > 0:
                    parts.append(f"{wt}:{val}")
        record["livestock_water_sources"] = "; ".join(parts)

    # 6.3 Irrigation
    s63 = re.search(r'6\.3\s+.*?[Ii]rrigation.*?(?=6\.4|$)', text, re.DOTALL | re.I)
    if s63:
        m = re.search(r'Major.*?[Ss]ources.*?:\s*(.+?)(?:\n\n|$)', s63.group(0), re.DOTALL)
        if m:
            record["irrigation_sources"] = _clean(m.group(1))

    # 6.4 Water quality
    s64 = re.search(r'6\.4\s+.*?(?=6\.5|$)', text, re.DOTALL | re.I)
    if s64:
        record["water_quality_changes"] = _clean(s64.group(0))[:1000]

    # 6.5 Traditional conservation
    s65 = re.search(r'6\.5\s+.*?(?=6\.6|$)', text, re.DOTALL | re.I)
    if s65:
        record["traditional_water_conservation"] = _clean(s65.group(0))[:1000]

    # 6.6 Important water bodies
    s66 = re.search(r'6\.6\s+.*?(?=Section|$)', text, re.DOTALL | re.I)
    if s66:
        record["important_water_bodies"] = _clean(s66.group(0))[:2000]


def parse_s7(text: str, record: dict):
    """Parse Section 7: Forest Lands."""
    m = re.search(r'Name\s+of\s+the\s+forest:\s*(.+?)(?:\n|$)', text, re.I)
    if m:
        record["forest_name"] = _clean(m.group(1))

    m = re.search(r'Type\s+of\s+Forest.*?:\s*(.+?)(?:\n|$)', text, re.I)
    if m:
        record["forest_type"] = _clean(m.group(1))

    m = re.search(r'Location.*?[Gg]eo\s*codes?\)?:\s*(.+?)(?:\n|$)', text, re.I)
    if m:
        record["forest_location_geocode"] = _clean(m.group(1))

    m = re.search(r'Size.*?(?:ha|hectare).*?:\s*(.+?)(?:\n|$)', text, re.I)
    if m:
        record["forest_size_ha"] = _clean(m.group(1))

    # Species composition (7.2)
    s72 = re.search(r'7\.2\s+.*?(?=7\.3|Section|$)', text, re.DOTALL | re.I)
    if s72:
        for group_name, field in [("Tree", "forest_tree_species"), ("Shrub", "forest_shrub_species"),
                                   ("Herb", "forest_herb_species"), ("Grass", "forest_grass_species"),
                                   ("Climber", "forest_climber_species")]:
            m2 = re.search(rf'{group_name}s?\s*\n(.*?)(?=\n(?:Tree|Shrub|Herb|Grass|Climber|$))',
                          s72.group(0), re.DOTALL | re.I)
            if m2:
                record[field] = _clean(m2.group(1))[:500]

    # NTFP (7.3)
    s73 = re.search(r'7\.3\s+.*?(?=Section|$)', text, re.DOTALL | re.I)
    if s73:
        record["forest_ntfp"] = _clean(s73.group(0))[:1000]


def parse_s8(text: str, record: dict):
    """Parse Section 8: Grassland."""
    m = re.search(r'Name.*?[Gg]razing.*?:\s*(.+?)(?:\n|$)', text, re.I)
    if m:
        record["grassland_name"] = _clean(m.group(1))
    s82 = re.search(r'8\.2\s+.*?(?=8\.3|Section|$)', text, re.DOTALL | re.I)
    if s82:
        record["grassland_species"] = _clean(s82.group(0))[:1000]
    s83 = re.search(r'8\.3\s+.*?(?=Section|$)', text, re.DOTALL | re.I)
    if s83:
        record["grassland_ntfp"] = _clean(s83.group(0))[:500]


def parse_s9(text: str, record: dict):
    """Parse Section 9: Wasteland."""
    s92 = re.search(r'9\.2\s+.*?(?=Section|$)', text, re.DOTALL | re.I)
    if s92:
        record["wasteland_species"] = _clean(s92.group(0))[:1000]


def parse_simple_section(text: str, record: dict, field: str, max_len: int = 2000):
    """Generic parser for sections that are primarily narrative text."""
    record[field] = _clean(text)[:max_len]


def parse_s12(text: str, record: dict):
    """Parse Section 12: Giant Trees."""
    entries = []
    # Pattern: Name + geocodes
    blocks = re.split(r'(?=\b[A-Z][\w\s]+(?:tree|süh|su)\b)', text, flags=re.I)
    for block in blocks:
        coords = re.findall(r'(\d{1,2}\.\d{4,8})', block)
        name_m = re.match(r'([\w\s()]+?)(?:\n|\d)', block)
        if name_m and coords:
            name = _clean(name_m.group(1))
            lat, lon = coords[0], coords[1] if len(coords) > 1 else ""
            entries.append(f"{name} ({lat},{lon})")

    if entries:
        record["giant_trees"] = "; ".join(entries)
    else:
        record["giant_trees"] = _clean(text)[:1500]


def parse_s16(text: str, record: dict):
    """Parse Section 16: Medicinal Plants."""
    entries = []
    # Look for scientific names — must be genus+species pattern (not instructions)
    sci_re = re.compile(r'([A-Z][a-z]{3,})\s+([a-z]{3,})')
    noise = {"information is", "collected through", "gather data", "group discussion",
             "tabular format", "section through", "instructions the"}

    for m in sci_re.finditer(text):
        sci_name = f"{m.group(1)} {m.group(2)}"
        if sci_name.lower() in noise:
            continue
        # Get surrounding context
        before = text[max(0, m.start() - 80):m.start()]
        after = text[m.end():m.end() + 100]
        # Local name: last word/line before scientific name
        local_lines = [l.strip() for l in before.split('\n') if l.strip()]
        local_name = local_lines[-1] if local_lines else ""
        # Use: text after scientific name until trend arrow
        use_m = re.search(r'\n\s*(.+?)(?:[↑↓↔]|\n\n)', after)
        use = _clean(use_m.group(1)) if use_m else ""
        entries.append(f"{local_name} | {sci_name} | {use}")

    if entries:
        record["medicinal_plants"] = "; ".join(entries[:30])
    else:
        # Fallback: clean raw text
        cleaned = re.sub(r'\(Instructions.*?\)', '', text, flags=re.DOTALL | re.I)
        record["medicinal_plants"] = _clean(cleaned)[:1500]


def parse_s17(text: str, record: dict):
    """Parse Section 17: Invasive Plants."""
    entries = []
    sci_re = re.compile(r'\(([A-Z][a-z]+\s+[a-z]+(?:\s+[a-z]+)?)\)')
    for m in sci_re.finditer(text):
        entries.append(m.group(1))

    if entries:
        record["invasive_plants"] = "; ".join(entries)
    else:
        record["invasive_plants"] = _clean(text)[:1500]


def parse_s20(text: str, record: dict):
    """Parse Section 20: Flora & Fauna — all 11 groups."""
    subsection_map = {
        "1": ("tree_diversity", "tree_diversity_count"),
        "2": ("shrub_diversity", "shrub_diversity_count"),
        "3": ("herb_grass_diversity", "herb_grass_diversity_count"),
        "4": ("lower_plant_diversity", "lower_plant_count"),
        "5": ("mammal_diversity", "mammal_count"),
        "6": ("bird_diversity", "bird_count"),
        "7": ("reptile_amphibian_diversity", "reptile_amphibian_count"),
        "8": ("butterfly_diversity", "butterfly_count"),
        "9": ("dragonfly_diversity", "dragonfly_count"),
        "10": ("fish_insect_other_diversity", "fish_insect_other_count"),
        "11": ("soil_macrofauna_diversity", "soil_macrofauna_count"),
    }

    # Find subsection positions
    sub_pat = re.compile(r'20\.(\d+)\s+(.+?)(?=\n)', re.I)
    matches = list(sub_pat.finditer(text))

    total_species = 0
    for i, m in enumerate(matches):
        sub_num = m.group(1)
        if sub_num not in subsection_map:
            continue

        field, count_field = subsection_map[sub_num]
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        section_text = text[start:end]

        # Extract species entries — look for local names and scientific names
        entries = []
        # Noise words that are NOT species
        noise_re = re.compile(
            r'^(Tick\s+in|Prepare|Local\s+Name|Flower|Forest\s*\(|Grazing|'
            r'CCA|Near\s+Water|In\s+and|Changes\s+in|Use/|Habitat|Group|'
            r'Herbarium|color|traditional|bodies|agriculture|years|over\s+past|'
            r'availability|Name\s+in|Name\s+of|place\s+in|where\s+it|mostly\s+seen|'
            r'photoguide|pocket\s+guide|photo\s+guide|Reasons?\s+for|Remarks|'
            r'Fore\b|st\s*\(\d|land\s*\(\d|Water\b|around\b|availabil|'
            r'past\s+25|off\s+white|Moss|lichen|Local|Number|number)', re.I)

        lines = section_text.split("\n")
        for line in lines:
            line = line.strip()
            if not line or len(line) < 3:
                continue
            if re.match(r'^[√✔↑↓↔()\s\d\-]+$', line):
                continue
            if noise_re.match(line):
                continue

            # Capture meaningful content
            sci_m = re.search(r'([A-Z][a-z]{2,})\s+([a-z]{2,})', line)
            if sci_m:
                name = f"{sci_m.group(1)} {sci_m.group(2)}"
                # Filter out common English false positives
                if name.lower() not in {"off white", "red ish", "olive green", "thick forest",
                                         "village area", "paddy field", "jhum areas", "reserve area"}:
                    entries.append(name)
            elif len(line) > 3 and not line.startswith(("(", "√", "✔")):
                clean_line = re.sub(r'[√✔↑↓↔]', '', line).strip()
                if clean_line and len(clean_line) > 2 and not noise_re.match(clean_line):
                    entries.append(clean_line)

        # Deduplicate while preserving order
        seen = set()
        unique = []
        for e in entries:
            e_lower = e.lower().strip()
            if e_lower not in seen and len(e_lower) > 2:
                seen.add(e_lower)
                unique.append(e)

        record[field] = "; ".join(unique[:100])
        record[count_field] = len(unique)
        total_species += len(unique)

    record["total_species_count"] = total_species


# ── Main extraction function ─────────────────────────────────

def extract_village(pdf_path: str, language: str = "Auto-detect",
                    progress_callback=None) -> dict:
    """Extract ALL data from a VER PDF into a flat record dict.

    Args:
        pdf_path: Path to the PDF file
        language: OCR language — one of SUPPORTED_LANGUAGES keys
        progress_callback: Optional callable(step, total, message) for UI progress

    Returns:
        OrderedDict with all master fields populated
    """
    record = get_empty_record()

    def progress(step, total, msg):
        if progress_callback:
            progress_callback(step, total, msg)

    progress(1, 10, f"Reading PDF ({language})...")

    def ocr_progress(page, total, msg):
        pct = page / total if total else 0
        progress(1 + int(pct * 2), 10, msg)

    pages, is_native = extract_text_from_pdf(pdf_path, language=language,
                                              progress_callback=ocr_progress)
    record["total_pages"] = len(pages)
    record["extraction_method"] = "native_text" if is_native else f"ocr_{language.lower()}"

    progress(3, 10, f"{'Native text' if is_native else f'OCR ({language})'} — {len(pages)} pages")

    progress(3, 10, "Detecting sections...")
    sections = find_section_ranges(pages)

    # Parse each section
    section_parsers = [
        ("s2", parse_s2, "General Information"),
        ("s3", parse_s3, "Village History"),
        ("s4", parse_s4, "Agriculture & Ecology"),
        ("s5", parse_s5, "Livestock"),
        ("s6", parse_s6, "Waterscape"),
        ("s7", parse_s7, "Forest Lands"),
        ("s8", parse_s8, "Grassland"),
        ("s9", lambda t, r: parse_simple_section(t, r, "wasteland_species"), "Wasteland"),
        ("s10", lambda t, r: parse_simple_section(t, r, "sacred_groves"), "Sacred Groves"),
        ("s11", lambda t, r: parse_simple_section(t, r, "ecological_sites"), "Ecological Sites"),
        ("s12", parse_s12, "Giant Trees"),
        ("s13", lambda t, r: parse_simple_section(t, r, "bee_hives"), "Bee Hives"),
        ("s14", lambda t, r: parse_simple_section(t, r, "fire_incidence"), "Fire Incidence"),
        ("s15", lambda t, r: parse_simple_section(t, r, "conservation_ethos"), "Conservation Ethos"),
        ("s16", parse_s16, "Medicinal Plants"),
        ("s17", parse_s17, "Invasive Plants"),
        ("s18", lambda t, r: parse_simple_section(t, r, "feral_animals"), "Feral Animals"),
        ("s19", lambda t, r: parse_simple_section(t, r, "protected_species"), "Protected Species"),
        ("s20", parse_s20, "Flora & Fauna"),
    ]

    for i, (key, parser, label) in enumerate(section_parsers):
        step = 4 + i * 6 // len(section_parsers)
        progress(min(step, 9), 10, f"Parsing {label}...")
        text = get_section_text(pages, sections, key)
        if text:
            try:
                parser(text, record)
            except Exception as e:
                print(f"Warning: Error parsing {key}: {e}")

    # Bamboo (part of section 15)
    s15_text = get_section_text(pages, sections, "s15")
    bamboo = re.search(r'(?:15\.1|Bamboo).*?(?=Section|$)', s15_text, re.DOTALL | re.I)
    if bamboo:
        record["bamboo_species"] = _clean(bamboo.group(0))[:1000]

    # Infer VER year from filename
    year_m = re.search(r'(\d{4}[-–]\d{2,4})', str(pdf_path))
    if year_m:
        record["ver_year"] = year_m.group(1)

    progress(10, 10, "Extraction complete!")
    return record


# ── CLI interface ─────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python comprehensive_extract.py <pdf_path> [<pdf_path2> ...]")
        sys.exit(1)

    for pdf_path in sys.argv[1:]:
        print(f"\nExtracting: {pdf_path}")
        record = extract_village(pdf_path)

        # Print non-empty fields
        for k, v in record.items():
            if v and v != 0:
                display = str(v)[:100] + "..." if len(str(v)) > 100 else str(v)
                print(f"  {k:40s} = {display}")

        print(f"\n  Total fields populated: {sum(1 for v in record.values() if v and v != 0)}/{len(record)}")
