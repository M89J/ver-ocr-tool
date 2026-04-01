"""
VER Data Extraction Tool — Streamlit Web Interface
Upload Village Ecological Register (VER) PDFs → Extract all 20 sections →
Download master Excel/CSV/GeoJSON sheet with all villages.

Run locally:  streamlit run app.py
Deploy:       Push to GitHub → Connect to Streamlit Community Cloud
"""
import streamlit as st
import pandas as pd
import os
import sys
import json
import tempfile
from pathlib import Path
from io import BytesIO
from datetime import datetime
from collections import OrderedDict

# Add ETL directory to path
sys.path.insert(0, str(Path(__file__).parent))
from comprehensive_extract import extract_village, get_empty_record, MASTER_FIELDS, SUPPORTED_LANGUAGES

# ── Page config ──────────────────────────────────────────────
st.set_page_config(
    page_title="VER Data Extractor",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Styling ──────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header { font-size: 2rem; font-weight: 700; color: #2d6a4f; margin-bottom: 0.5rem; }
    .sub-header { font-size: 1rem; color: #555; margin-bottom: 2rem; }
    .stat-card { background: #f8f9fa; border-radius: 8px; padding: 1rem; text-align: center; border: 1px solid #e0e0e0; }
    .stat-number { font-size: 1.8rem; font-weight: 700; color: #2d6a4f; }
    .stat-label { font-size: 0.85rem; color: #666; }
    .section-header { background: #2d6a4f; color: white; padding: 0.5rem 1rem; border-radius: 4px; margin: 1rem 0 0.5rem 0; }
    .stDataFrame { font-size: 0.85rem; }
</style>
""", unsafe_allow_html=True)

# ── Session state ────────────────────────────────────────────
if "extracted_data" not in st.session_state:
    st.session_state.extracted_data = []  # list of OrderedDicts
if "processing" not in st.session_state:
    st.session_state.processing = False


def generate_excel(records: list[dict]) -> bytes:
    """Generate Excel file from list of village records."""
    if not records:
        return b""

    df = pd.DataFrame(records)

    # Reorder columns to match MASTER_FIELDS order
    ordered_cols = [c for c in MASTER_FIELDS.keys() if c in df.columns]
    extra_cols = [c for c in df.columns if c not in ordered_cols]
    df = df[ordered_cols + extra_cols]

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="VER Master Data", index=False)

        workbook = writer.book
        worksheet = writer.sheets["VER Master Data"]

        # Header format
        header_fmt = workbook.add_format({
            "bold": True, "bg_color": "#2d6a4f", "font_color": "white",
            "border": 1, "text_wrap": True, "valign": "top",
        })
        for col_idx, col_name in enumerate(df.columns):
            worksheet.write(0, col_idx, col_name, header_fmt)

        # Column widths
        for i, col in enumerate(df.columns):
            max_len = max(len(str(col)), df[col].astype(str).str.len().max())
            worksheet.set_column(i, i, min(max(max_len * 0.9, 12), 50))

        # Freeze top row and first 2 columns (village_name, state)
        worksheet.freeze_panes(1, 2)

        # Add an overview sheet
        overview_data = []
        for rec in records:
            overview_data.append({
                "Village": rec.get("village_name", ""),
                "State": rec.get("state", ""),
                "Lat": rec.get("latitude", ""),
                "Lon": rec.get("longitude", ""),
                "Area (ha)": rec.get("total_area_ha", ""),
                "Population": rec.get("total_population", ""),
                "Households": rec.get("total_households", ""),
                "Livestock": rec.get("livestock_summary", ""),
                "Water Sources": rec.get("drinking_water_sources", ""),
                "Total Species": rec.get("total_species_count", 0),
                "Trees": rec.get("tree_diversity_count", 0),
                "Birds": rec.get("bird_count", 0),
                "Mammals": rec.get("mammal_count", 0),
                "Extraction": rec.get("extraction_method", ""),
                "Pages": rec.get("total_pages", 0),
            })
        overview_df = pd.DataFrame(overview_data)
        overview_df.to_excel(writer, sheet_name="Overview", index=False)

        ws2 = writer.sheets["Overview"]
        for col_idx, col_name in enumerate(overview_df.columns):
            ws2.write(0, col_idx, col_name, header_fmt)
            ws2.set_column(col_idx, col_idx, 18)
        ws2.freeze_panes(1, 1)

    return output.getvalue()


def generate_csv(records: list[dict]) -> str:
    """Generate CSV from list of village records."""
    if not records:
        return ""
    df = pd.DataFrame(records)
    ordered_cols = [c for c in MASTER_FIELDS.keys() if c in df.columns]
    return df[ordered_cols].to_csv(index=False)


def generate_geojson(records: list[dict]) -> str:
    """Generate GeoJSON FeatureCollection from village records."""
    features = []
    for rec in records:
        lat = rec.get("latitude")
        lon = rec.get("longitude")

        # Build geometry (null if no coords)
        geometry = None
        if lat and lon:
            try:
                geometry = {
                    "type": "Point",
                    "coordinates": [float(lon), float(lat)]
                }
            except (ValueError, TypeError):
                pass

        # All fields as properties (skip empty ones for cleaner output)
        properties = OrderedDict()
        for k, v in rec.items():
            if k in ("latitude", "longitude"):
                continue
            if v and v != 0:
                properties[k] = v

        features.append({
            "type": "Feature",
            "geometry": geometry,
            "properties": properties,
        })

    geojson = {
        "type": "FeatureCollection",
        "name": "VER_Master_Data",
        "features": features,
    }
    return json.dumps(geojson, indent=2, ensure_ascii=False, default=str)


# ── Sidebar ──────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🌿 VER Data Extractor")
    st.markdown("Upload Village Ecological Register PDFs and extract comprehensive data into a single master sheet.")
    st.divider()

    st.markdown("**Extracted Villages**")
    if st.session_state.extracted_data:
        for i, rec in enumerate(st.session_state.extracted_data):
            name = rec.get("village_name", f"Village {i+1}")
            state = rec.get("state", "")
            species = rec.get("total_species_count", 0)
            method = rec.get("extraction_method", "")
            icon = "📄" if method == "native_text" else "🔍"
            st.markdown(f"{icon} **{name}** ({state}) — {species} species")
    else:
        st.markdown("*No villages extracted yet*")

    st.divider()
    if st.session_state.extracted_data:
        if st.button("🗑️ Clear All Data", use_container_width=True):
            st.session_state.extracted_data = []
            st.rerun()

    st.markdown("---")
    st.markdown("**Languages**")
    st.markdown("English, Hindi, Odia, Tamil, Telugu, Kannada, Marathi, Gujarati")
    st.markdown("**Fields:** 114 columns per village")
    st.markdown("**Sections:** All 20 VER sections")


# ── Main content ─────────────────────────────────────────────
st.markdown('<div class="main-header">VER Data Extraction Tool</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Upload VER PDFs → Extract all 20 sections → Download master Excel sheet with all villages</div>', unsafe_allow_html=True)

# Upload section
col_upload, col_lang, col_info = st.columns([2, 1, 1])

with col_upload:
    uploaded_files = st.file_uploader(
        "Upload VER PDF(s)",
        type=["pdf"],
        accept_multiple_files=True,
        help="Upload one or more VER PDF files. Each PDF will be processed and added to the master sheet.",
    )

with col_lang:
    selected_language = st.selectbox(
        "PDF Language",
        options=list(SUPPORTED_LANGUAGES.keys()),
        index=0,
        help="Select the language of the scanned PDF. For typed/digital PDFs, language is auto-detected. For scanned PDFs, this sets the OCR language.",
    )
    if selected_language != "Auto-detect":
        lang_info = SUPPORTED_LANGUAGES[selected_language]
        st.caption(f"Script: {lang_info['script']} | OCR: `{lang_info['tesseract']}`")

with col_info:
    st.info(
        f"**{len(st.session_state.extracted_data)}** village(s) extracted so far.\n\n"
        "Upload PDFs and click **Extract** to process them."
    )

# Extract button
if uploaded_files:
    if st.button("🚀 Extract Data from PDFs", type="primary", use_container_width=True):
        progress_bar = st.progress(0, text="Starting extraction...")
        status_text = st.empty()

        for file_idx, uploaded_file in enumerate(uploaded_files):
            # Save to temp file
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                tmp.write(uploaded_file.read())
                tmp_path = tmp.name

            file_label = uploaded_file.name

            def progress_cb(step, total, msg):
                pct = ((file_idx + step / total) / len(uploaded_files))
                progress_bar.progress(min(pct, 1.0), text=f"[{file_idx+1}/{len(uploaded_files)}] {file_label}: {msg}")

            status_text.markdown(f"Processing **{file_label}**...")

            try:
                record = extract_village(tmp_path, language=selected_language,
                                        progress_callback=progress_cb)

                # Use filename as fallback village name
                if not record.get("village_name"):
                    name_from_file = Path(file_label).stem.replace("VER_", "").replace("_", " ")
                    record["village_name"] = name_from_file.split(" ")[0]

                st.session_state.extracted_data.append(record)
            except Exception as e:
                st.error(f"Error processing {file_label}: {e}")
            finally:
                os.unlink(tmp_path)

        progress_bar.progress(1.0, text="All PDFs processed!")
        status_text.success(f"Extracted {len(uploaded_files)} PDF(s) successfully!")
        st.rerun()

# ── Results ──────────────────────────────────────────────────
if st.session_state.extracted_data:
    st.divider()

    # Stats row
    records = st.session_state.extracted_data
    cols = st.columns(5)
    with cols[0]:
        st.markdown(f'<div class="stat-card"><div class="stat-number">{len(records)}</div><div class="stat-label">Villages</div></div>', unsafe_allow_html=True)
    with cols[1]:
        total_species = sum(r.get("total_species_count", 0) for r in records)
        st.markdown(f'<div class="stat-card"><div class="stat-number">{total_species}</div><div class="stat-label">Total Species</div></div>', unsafe_allow_html=True)
    with cols[2]:
        total_pop = sum(int(r.get("total_population", 0) or 0) for r in records)
        st.markdown(f'<div class="stat-card"><div class="stat-number">{total_pop:,}</div><div class="stat-label">Population</div></div>', unsafe_allow_html=True)
    with cols[3]:
        fields_filled = sum(sum(1 for v in r.values() if v and v != 0) for r in records)
        total_fields = len(records) * len(MASTER_FIELDS)
        pct = round(fields_filled / total_fields * 100) if total_fields else 0
        st.markdown(f'<div class="stat-card"><div class="stat-number">{pct}%</div><div class="stat-label">Fields Populated</div></div>', unsafe_allow_html=True)
    with cols[4]:
        states = set(r.get("state", "") for r in records if r.get("state"))
        st.markdown(f'<div class="stat-card"><div class="stat-number">{len(states)}</div><div class="stat-label">States</div></div>', unsafe_allow_html=True)

    # Download buttons
    st.markdown("### Download Master Sheet")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    dl_col1, dl_col2, dl_col3, dl_col4 = st.columns(4)
    with dl_col1:
        excel_data = generate_excel(records)
        st.download_button(
            label="📥 Excel (.xlsx)",
            data=excel_data,
            file_name=f"VER_Master_Data_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    with dl_col2:
        csv_data = generate_csv(records)
        st.download_button(
            label="📥 CSV",
            data=csv_data,
            file_name=f"VER_Master_Data_{timestamp}.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with dl_col3:
        geojson_data = generate_geojson(records)
        st.download_button(
            label="🗺️ GeoJSON",
            data=geojson_data,
            file_name=f"VER_Master_Data_{timestamp}.geojson",
            mime="application/geo+json",
            use_container_width=True,
        )
    with dl_col4:
        json_data = json.dumps(records, indent=2, ensure_ascii=False, default=str)
        st.download_button(
            label="📥 JSON",
            data=json_data,
            file_name=f"VER_Master_Data_{timestamp}.json",
            mime="application/json",
            use_container_width=True,
        )

    # Data preview
    st.markdown("### Data Preview")

    # Village selector
    village_names = [r.get("village_name", f"Village {i+1}") for i, r in enumerate(records)]
    view_mode = st.radio("View", ["All Villages (Overview)", "Single Village (Detail)"], horizontal=True)

    if view_mode == "All Villages (Overview)":
        # Overview table with key fields
        overview_cols = [
            "village_name", "state", "latitude", "longitude", "total_area_ha",
            "total_population", "total_households", "major_livelihoods",
            "livestock_summary", "drinking_water_sources", "total_species_count",
            "tree_diversity_count", "bird_count", "mammal_count",
        ]
        df = pd.DataFrame(records)[overview_cols]
        st.dataframe(df, use_container_width=True, height=400)

    else:
        selected = st.selectbox("Select Village", village_names)
        idx = village_names.index(selected)
        rec = records[idx]

        # Section tabs
        tabs = st.tabs([
            "📋 General", "📜 History", "🌾 Agriculture", "🐄 Livestock",
            "💧 Water", "🌲 Forest", "🦋 Biodiversity", "🌿 Conservation", "📊 All Fields"
        ])

        with tabs[0]:  # General
            st.markdown(f"**Village:** {rec.get('village_name', '')} | **State:** {rec.get('state', '')} | **Block:** {rec.get('block', '')}")
            st.markdown(f"**GPS:** {rec.get('latitude', '')}, {rec.get('longitude', '')} | **Date:** {rec.get('date_of_survey', '')}")
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**Land Use**")
                for f in ["total_area_ha", "forest_land_pct", "grazing_land_pct", "community_conserved_area_pct", "agricultural_land_pct", "other_land_pct"]:
                    if rec.get(f):
                        st.markdown(f"- {f.replace('_', ' ').title()}: **{rec[f]}**")
            with c2:
                st.markdown("**Demographics**")
                for f in ["total_population", "total_households", "large_farmers_gt10ha", "medium_farmers_4_10ha", "small_farmers_1_2ha", "marginal_farmers_lt1ha"]:
                    if rec.get(f):
                        st.markdown(f"- {f.replace('_', ' ').title()}: **{rec[f]}**")
            if rec.get("major_livelihoods"):
                st.markdown(f"**Livelihoods:** {rec['major_livelihoods']}")

        with tabs[1]:  # History
            for f in ["village_history_narrative", "myths_and_beliefs", "traditional_songs"]:
                if rec.get(f):
                    st.markdown(f"**{f.replace('_', ' ').title()}**")
                    st.text_area("", rec[f], height=150, key=f"hist_{f}", disabled=True)

        with tabs[2]:  # Agriculture
            for f in ["kharif_crops", "rabi_crops", "zaid_crops", "traditional_crop_varieties", "farming_practices",
                       "soil_type", "soil_fertility_change", "soil_fertility_reason", "pest_incidences", "major_weeds"]:
                if rec.get(f):
                    st.markdown(f"**{f.replace('_', ' ').title()}:** {rec[f][:300]}")

        with tabs[3]:  # Livestock
            if rec.get("livestock_summary"):
                st.markdown(f"**Summary:** {rec['livestock_summary']}")
            if rec.get("livestock_detailed"):
                st.markdown(f"**Detailed:** {rec['livestock_detailed']}")
            for f in ["indigenous_breeds", "livestock_diseases", "ethno_veterinary_practices"]:
                if rec.get(f):
                    st.text_area(f.replace("_", " ").title(), rec[f][:500], height=100, key=f"ls_{f}", disabled=True)

        with tabs[4]:  # Water
            for f in ["drinking_water_sources", "livestock_water_sources", "irrigation_sources",
                       "water_quality_changes", "important_water_bodies"]:
                if rec.get(f):
                    st.markdown(f"**{f.replace('_', ' ').title()}:** {rec[f][:300]}")

        with tabs[5]:  # Forest
            for f in ["forest_name", "forest_type", "forest_size_ha", "forest_location_geocode",
                       "forest_tree_species", "forest_shrub_species", "forest_herb_species", "forest_ntfp"]:
                if rec.get(f):
                    st.markdown(f"**{f.replace('_', ' ').title()}:** {rec[f][:300]}")

        with tabs[6]:  # Biodiversity
            st.markdown(f"**Total Species: {rec.get('total_species_count', 0)}**")
            bio_fields = [
                ("tree_diversity_count", "tree_diversity", "Trees"),
                ("shrub_diversity_count", "shrub_diversity", "Shrubs"),
                ("herb_grass_diversity_count", "herb_grass_diversity", "Herbs & Grasses"),
                ("mammal_count", "mammal_diversity", "Mammals"),
                ("bird_count", "bird_diversity", "Birds"),
                ("reptile_amphibian_count", "reptile_amphibian_diversity", "Reptiles & Amphibians"),
                ("butterfly_count", "butterfly_diversity", "Butterflies"),
                ("dragonfly_count", "dragonfly_diversity", "Dragonflies"),
            ]
            for count_f, list_f, label in bio_fields:
                count = rec.get(count_f, 0)
                if count:
                    with st.expander(f"{label} ({count})"):
                        st.markdown(rec.get(list_f, ""))

        with tabs[7]:  # Conservation
            for f in ["sacred_groves", "conservation_ethos", "bamboo_species", "medicinal_plants",
                       "invasive_plants", "protected_species", "feral_animals", "fire_incidence"]:
                if rec.get(f):
                    st.text_area(f.replace("_", " ").title(), rec[f][:500], height=80, key=f"cons_{f}", disabled=True)

        with tabs[8]:  # All Fields
            all_data = [(k, str(v)[:200]) for k, v in rec.items() if v and v != 0]
            st.dataframe(pd.DataFrame(all_data, columns=["Field", "Value"]), use_container_width=True, height=600)
