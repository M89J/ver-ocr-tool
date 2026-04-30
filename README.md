# VER Data Extraction Tool

A web-based tool to extract comprehensive data from **Village Ecological Register (VER)** PDFs into structured Excel/CSV/GeoJSON formats.

## What it does

Upload VER PDFs and get a master spreadsheet with **113 fields** covering all 20 VER sections:

| Section | Data Extracted |
|---------|---------------|
| S2: General Info | Village name, state, GPS, area, land use %, population, households, caste, landholding, livelihoods |
| S3: History | Village history narrative, myths, traditional songs |
| S4: Agriculture | Kharif/Rabi/Zaid crops, traditional varieties, farming practices, soil health, pests, weeds |
| S5: Livestock | Numbers with 10/25/50yr trends, indigenous breeds, diseases, ethno-vet practices |
| S6: Water | Drinking/livestock sources with counts, irrigation, quality changes, traditional conservation |
| S7-S9: Land | Forest/grassland/wasteland name, size, species composition, NTFP |
| S10-S15 | Sacred groves, ecological sites, giant trees (GPS), bee hives, fire, conservation, bamboo |
| S16-S19 | Medicinal plants, invasive plants, feral animals, protected species |
| S20: Biodiversity | 11 groups: trees, shrubs, herbs, mammals, birds, reptiles, butterflies, dragonflies, fish/insects, soil macrofauna |

## Output formats

- **Excel (.xlsx)** — Two sheets: Overview + full 113-column master data
- **CSV** — Flat table, one row per village
- **GeoJSON** — Point features with all properties, ready for GIS/mapping
- **JSON** — Raw structured data

## How to use

### Online (Streamlit Cloud)
Visit the deployed app (link in repo description) and upload your PDFs.

### Local
```bash
pip install -r requirements.txt
streamlit run app.py
```

## How it works

1. **Native text detection** — Auto-detects if a PDF has extractable text (typed/digital) vs scanned images
2. **For native text PDFs** — Direct extraction via PyMuPDF at 99% accuracy
3. **For scanned PDFs** — Falls back to Tesseract OCR (requires local Tesseract installation)
4. **Section detection** — Identifies all 20 VER sections using multilingual header patterns (English, Odia, Hindi, Marathi)
5. **Comprehensive parsing** — Each section has a dedicated parser that extracts structured fields
6. **Multi-village support** — Upload multiple PDFs and get all villages in a single master sheet

## Supported languages

| State | Language | Script | Status |
|-------|----------|--------|--------|
| Nagaland / NE India | English | Latin | Excellent (native text) |
| Odisha | Odia + Hindi + English | Odia | Partial (scanned handwriting is challenging) |
| Maharashtra | Marathi + English | Devanagari | Partial |
| Gujarat | Gujarati + English | Gujarati | Planned |
| Chhattisgarh | Hindi + English | Devanagari | Planned |
| Rajasthan | Hindi + English | Devanagari | Planned |

## Optional: better OCR for Indic scripts

Two fallback engines are wired in. Both kick in **only when Tesseract confidence is low** — so they never slow down high-quality pages. Choose one via `OCR_FALLBACK`.

### Option 1 — EasyOCR (local, no API)

```bash
pip install -r requirements-easyocr.txt
export OCR_FALLBACK=easyocr
streamlit run app.py
```

- Languages: Hindi, Marathi, Tamil, Telugu, Kannada, Bengali (Odia/Gujarati not supported)
- Adds ~500MB model + ~1GB RAM — **does not fit free Streamlit Cloud**
- Apache 2.0, no API keys, no cost

### Option 2 — Bhashini (Indian Govt API, free for citizen-science)

```bash
export OCR_FALLBACK=bhashini
export BHASHINI_USER_ID="<your-ulca-user-id>"
export BHASHINI_API_KEY="<your-ulca-api-key>"
streamlit run app.py
```

- **Free** for individuals/research at https://bhashini.gov.in/ulca (sign up for ULCA, get user ID + API key)
- Languages: Hindi, Marathi, Tamil, Telugu, Kannada, Bengali, **Odia**, Gujarati
- Runs as HTTP API — no extra RAM, fits free Streamlit Cloud
- Falls back silently to Tesseract if the Bhashini service is unreachable
- Note: pipeline IDs and service availability change occasionally; if accuracy drops, check https://bhashini.gov.in/ulca for current OCR pipelines

## About VER

The Village Ecological Register is a citizen science document prepared by community members across ~200 Indian villages in 7+ states. It captures traditional ecological knowledge, biodiversity data, and socio-ecological changes spanning 50+ years.

## License

Open source. Part of the [VER Digital Platform](https://github.com/M89J/ver-platform) project.
