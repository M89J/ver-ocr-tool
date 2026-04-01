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
from db import save_village, load_all_villages, delete_village, delete_all_villages, get_village_count
import folium
from streamlit_folium import st_folium

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

# ── Session state (load from DB on first visit) ─────────────
if "extracted_data" not in st.session_state:
    st.session_state.extracted_data = load_all_villages()
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

    # ── Import / Export ──────────────────────────────────────
    st.markdown("**Data Backup**")
    if st.session_state.extracted_data:
        # Export
        export_data = json.dumps(
            [{k: v for k, v in r.items() if not k.startswith("_")} for r in st.session_state.extracted_data],
            ensure_ascii=False, default=str,
        )
        st.download_button(
            label="💾 Export Backup (JSON)",
            data=export_data,
            file_name=f"VER_backup_{datetime.now().strftime('%Y%m%d')}.json",
            mime="application/json",
            use_container_width=True,
        )

    # Import
    imported_file = st.file_uploader("📂 Import Backup", type=["json"], help="Import a previously exported JSON backup to restore village data.")
    if imported_file:
        try:
            imported_records = json.loads(imported_file.read().decode("utf-8"))
            if isinstance(imported_records, list) and len(imported_records) > 0:
                if st.button(f"Restore {len(imported_records)} village(s)", type="primary", use_container_width=True):
                    delete_all_villages()
                    for rec in imported_records:
                        save_village(rec)
                    st.session_state.extracted_data = load_all_villages()
                    st.success(f"Restored {len(imported_records)} village(s)!")
                    st.rerun()
            else:
                st.warning("Invalid backup file — expected a JSON list of village records.")
        except (json.JSONDecodeError, UnicodeDecodeError):
            st.error("Could not read this file. Please use a valid JSON backup.")

    st.divider()
    if st.session_state.extracted_data:
        if st.button("🗑️ Clear All Data", use_container_width=True):
            delete_all_villages()
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

                # Save to database and add DB metadata
                db_id = save_village(record)
                record["_db_id"] = db_id
                record["_created_at"] = datetime.now().isoformat()
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

    all_records = st.session_state.extracted_data

    # ── Search & Filter ─────────────────────────────────────
    st.markdown("### Search & Filter")
    filter_col1, filter_col2, filter_col3 = st.columns([2, 1, 1])

    with filter_col1:
        search_query = st.text_input(
            "Search villages",
            placeholder="Search by village name, crop, species, livestock, water source...",
            help="Searches across all text fields in the extracted data.",
        )

    with filter_col2:
        all_states = sorted(set(r.get("state", "") for r in all_records if r.get("state")))
        selected_states = st.multiselect("Filter by State", options=all_states, default=[])

    with filter_col3:
        species_counts = [r.get("total_species_count", 0) for r in all_records]
        min_sp, max_sp = min(species_counts, default=0), max(species_counts, default=500)
        if min_sp < max_sp:
            species_range = st.slider("Species Count", min_value=min_sp, max_value=max_sp, value=(min_sp, max_sp))
        else:
            species_range = (min_sp, max_sp)

    # Apply filters
    records = all_records
    if search_query:
        query_lower = search_query.lower()
        records = [
            r for r in records
            if any(query_lower in str(v).lower() for v in r.values())
        ]
    if selected_states:
        records = [r for r in records if r.get("state", "") in selected_states]
    if min_sp < max_sp:
        records = [r for r in records if species_range[0] <= r.get("total_species_count", 0) <= species_range[1]]

    if len(records) < len(all_records):
        st.caption(f"Showing **{len(records)}** of {len(all_records)} villages")

    if not records:
        st.warning("No villages match your filters. Try adjusting your search or filters.")
        st.stop()

    # Stats row
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

    # ── Interactive Map ──────────────────────────────────────
    villages_with_coords = [r for r in records if r.get("latitude") and r.get("longitude")]

    if villages_with_coords:
        st.markdown("### Village Map")

        # Calculate map center from average of all coordinates
        lats = []
        lons = []
        for r in villages_with_coords:
            try:
                lats.append(float(r["latitude"]))
                lons.append(float(r["longitude"]))
            except (ValueError, TypeError):
                continue

        if lats and lons:
            center_lat = sum(lats) / len(lats)
            center_lon = sum(lons) / len(lons)

            m = folium.Map(location=[center_lat, center_lon], zoom_start=6,
                           tiles="OpenStreetMap")

            # Color based on species count
            def get_marker_color(species_count):
                if species_count >= 200:
                    return "darkgreen"
                elif species_count >= 100:
                    return "green"
                elif species_count >= 50:
                    return "orange"
                else:
                    return "red"

            for r in villages_with_coords:
                try:
                    lat = float(r["latitude"])
                    lon = float(r["longitude"])
                except (ValueError, TypeError):
                    continue

                name = r.get("village_name", "Unknown")
                state = r.get("state", "")
                species = r.get("total_species_count", 0)
                pop = r.get("total_population", "N/A")
                area = r.get("total_area_ha", "N/A")
                trees = r.get("tree_diversity_count", 0)
                birds = r.get("bird_count", 0)
                mammals = r.get("mammal_count", 0)

                popup_html = f"""
                <div style="font-family: sans-serif; min-width: 200px;">
                    <h4 style="margin:0; color:#2d6a4f;">{name}</h4>
                    <p style="margin:2px 0; color:#555;">{state}</p>
                    <hr style="margin:4px 0;">
                    <table style="font-size:12px;">
                        <tr><td><b>Population</b></td><td>{pop}</td></tr>
                        <tr><td><b>Area (ha)</b></td><td>{area}</td></tr>
                        <tr><td><b>Total Species</b></td><td>{species}</td></tr>
                        <tr><td><b>Trees</b></td><td>{trees}</td></tr>
                        <tr><td><b>Birds</b></td><td>{birds}</td></tr>
                        <tr><td><b>Mammals</b></td><td>{mammals}</td></tr>
                    </table>
                </div>
                """

                folium.Marker(
                    location=[lat, lon],
                    popup=folium.Popup(popup_html, max_width=300),
                    tooltip=f"{name} ({species} species)",
                    icon=folium.Icon(color=get_marker_color(species), icon="leaf", prefix="fa"),
                ).add_to(m)

            # Legend
            legend_html = """
            <div style="position:fixed; bottom:30px; left:30px; z-index:1000;
                        background:white; padding:10px; border-radius:5px;
                        border:1px solid #ccc; font-size:12px;">
                <b>Species Count</b><br>
                <i class="fa fa-leaf" style="color:darkgreen;"></i> 200+&nbsp;
                <i class="fa fa-leaf" style="color:green;"></i> 100-199&nbsp;
                <i class="fa fa-leaf" style="color:orange;"></i> 50-99&nbsp;
                <i class="fa fa-leaf" style="color:red;"></i> &lt;50
            </div>
            """
            m.get_root().html.add_child(folium.Element(legend_html))

            st_folium(m, width=None, height=500, use_container_width=True)
    elif len(records) > 0:
        st.info("No GPS coordinates found in extracted data. Map will appear when villages have latitude/longitude.")

    # ── Charts & Visualization ──────────────────────────────
    st.markdown("### Charts & Visualization")
    chart_tabs = st.tabs(["🦋 Biodiversity", "🌍 Land Use", "👥 Demographics", "🌾 Agriculture"])

    with chart_tabs[0]:  # Biodiversity
        import plotly.graph_objects as go

        village_names_chart = [r.get("village_name", f"Village {i+1}") for i, r in enumerate(records)]
        bio_categories = [
            ("Trees", "tree_diversity_count", "#2d6a4f"),
            ("Shrubs", "shrub_diversity_count", "#52b788"),
            ("Herbs & Grasses", "herb_grass_diversity_count", "#95d5b2"),
            ("Mammals", "mammal_count", "#d4a373"),
            ("Birds", "bird_count", "#e9c46a"),
            ("Reptiles & Amphibians", "reptile_amphibian_count", "#e76f51"),
            ("Butterflies", "butterfly_count", "#f4a261"),
            ("Dragonflies", "dragonfly_count", "#264653"),
        ]

        fig = go.Figure()
        for label, field, color in bio_categories:
            values = [r.get(field, 0) for r in records]
            if any(v > 0 for v in values):
                fig.add_trace(go.Bar(name=label, x=village_names_chart, y=values, marker_color=color))

        fig.update_layout(
            barmode="stack",
            title="Biodiversity Breakdown by Village",
            xaxis_title="Village",
            yaxis_title="Species Count",
            height=450,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(t=80, b=40),
        )
        st.plotly_chart(fig, use_container_width=True)

    with chart_tabs[1]:  # Land Use
        import plotly.express as px

        land_fields = [
            ("Forest", "forest_land_pct"),
            ("Grazing", "grazing_land_pct"),
            ("Community Conserved", "community_conserved_area_pct"),
            ("Agricultural", "agricultural_land_pct"),
            ("Other", "other_land_pct"),
        ]

        if len(records) == 1:
            # Pie chart for single village
            rec = records[0]
            pie_labels, pie_values = [], []
            for label, field in land_fields:
                val = rec.get(field, "")
                try:
                    num = float(str(val).replace("%", "").strip())
                    if num > 0:
                        pie_labels.append(label)
                        pie_values.append(num)
                except (ValueError, TypeError):
                    pass
            if pie_values:
                fig = px.pie(names=pie_labels, values=pie_values,
                             title=f"Land Use — {rec.get('village_name', 'Village')}",
                             color_discrete_sequence=["#2d6a4f", "#95d5b2", "#b7e4c7", "#e9c46a", "#ccc"])
                fig.update_layout(height=400)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No land use data available.")
        else:
            # Grouped bar for multiple villages
            fig = go.Figure()
            colors = ["#2d6a4f", "#95d5b2", "#b7e4c7", "#e9c46a", "#ccc"]
            for (label, field), color in zip(land_fields, colors):
                values = []
                for r in records:
                    val = r.get(field, "")
                    try:
                        values.append(float(str(val).replace("%", "").strip()))
                    except (ValueError, TypeError):
                        values.append(0)
                if any(v > 0 for v in values):
                    fig.add_trace(go.Bar(name=label, x=village_names_chart, y=values, marker_color=color))
            fig.update_layout(
                barmode="group",
                title="Land Use Comparison (%)",
                xaxis_title="Village", yaxis_title="Percentage",
                height=450,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                margin=dict(t=80, b=40),
            )
            st.plotly_chart(fig, use_container_width=True)

    with chart_tabs[2]:  # Demographics
        pop_data = []
        for r in records:
            name = r.get("village_name", "?")
            try:
                pop = int(r.get("total_population", 0) or 0)
            except (ValueError, TypeError):
                pop = 0
            try:
                hh = int(r.get("total_households", 0) or 0)
            except (ValueError, TypeError):
                hh = 0
            pop_data.append({"Village": name, "Population": pop, "Households": hh})

        if any(d["Population"] > 0 for d in pop_data):
            df_pop = pd.DataFrame(pop_data)
            fig = go.Figure()
            fig.add_trace(go.Bar(name="Population", x=df_pop["Village"], y=df_pop["Population"], marker_color="#2d6a4f"))
            fig.add_trace(go.Bar(name="Households", x=df_pop["Village"], y=df_pop["Households"], marker_color="#e9c46a"))
            fig.update_layout(
                barmode="group",
                title="Population & Households",
                xaxis_title="Village", yaxis_title="Count",
                height=400,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                margin=dict(t=80, b=40),
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No population data available.")

    with chart_tabs[3]:  # Agriculture
        crop_types = [
            ("Kharif Crops", "kharif_crops", "#2d6a4f"),
            ("Rabi Crops", "rabi_crops", "#e9c46a"),
            ("Zaid Crops", "zaid_crops", "#e76f51"),
        ]
        has_crop_data = False
        for r in records:
            for _, field, _ in crop_types:
                if r.get(field):
                    has_crop_data = True
                    break

        if has_crop_data:
            for r in records:
                name = r.get("village_name", "Village")
                st.markdown(f"**{name}**")
                for label, field, color in crop_types:
                    crops = r.get(field, "")
                    if crops:
                        st.markdown(f"- **{label}:** {crops[:300]}")
            if any(r.get("traditional_crop_varieties") for r in records):
                st.markdown("---")
                st.markdown("**Traditional Varieties**")
                for r in records:
                    if r.get("traditional_crop_varieties"):
                        st.markdown(f"- **{r.get('village_name', '?')}:** {r['traditional_crop_varieties'][:300]}")
        else:
            st.info("No agriculture data available.")

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

    # ── Cross-Village Comparison ────────────────────────────
    if len(records) >= 2:
        st.divider()
        st.markdown("### Cross-Village Comparison")

        compare_names = [r.get("village_name", f"Village {i+1}") for i, r in enumerate(records)]
        selected_compare = st.multiselect(
            "Select villages to compare",
            options=compare_names,
            default=compare_names[:min(3, len(compare_names))],
            help="Pick 2 or more villages for side-by-side comparison.",
        )

        if len(selected_compare) >= 2:
            compare_records = [records[compare_names.index(n)] for n in selected_compare]
            compare_aspect = st.radio(
                "Compare by",
                ["Overview", "Biodiversity", "Land Use", "Water & Livestock", "All Fields"],
                horizontal=True,
                key="compare_aspect",
            )

            if compare_aspect == "Overview":
                rows = [
                    ("State", "state"),
                    ("Population", "total_population"),
                    ("Households", "total_households"),
                    ("Area (ha)", "total_area_ha"),
                    ("Total Species", "total_species_count"),
                    ("Livelihoods", "major_livelihoods"),
                    ("GPS", None),
                ]
                table_data = {"Field": [r[0] for r in rows]}
                for rec in compare_records:
                    col_vals = []
                    for label, field in rows:
                        if field is None:
                            col_vals.append(f"{rec.get('latitude', '')}, {rec.get('longitude', '')}")
                        else:
                            col_vals.append(str(rec.get(field, "") or ""))
                    table_data[rec.get("village_name", "?")] = col_vals
                st.dataframe(pd.DataFrame(table_data).set_index("Field"), use_container_width=True)

            elif compare_aspect == "Biodiversity":
                import plotly.graph_objects as go
                bio_cats = [
                    ("Trees", "tree_diversity_count"),
                    ("Shrubs", "shrub_diversity_count"),
                    ("Herbs", "herb_grass_diversity_count"),
                    ("Mammals", "mammal_count"),
                    ("Birds", "bird_count"),
                    ("Reptiles", "reptile_amphibian_count"),
                    ("Butterflies", "butterfly_count"),
                    ("Dragonflies", "dragonfly_count"),
                ]
                fig = go.Figure()
                colors = ["#2d6a4f", "#52b788", "#95d5b2", "#d4a373", "#e9c46a", "#e76f51", "#f4a261", "#264653"]
                for rec in compare_records:
                    name = rec.get("village_name", "?")
                    values = [rec.get(f, 0) for _, f in bio_cats]
                    fig.add_trace(go.Bar(name=name, x=[c[0] for c in bio_cats], y=values))
                fig.update_layout(
                    barmode="group", title="Biodiversity Comparison",
                    yaxis_title="Species Count", height=450,
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                )
                st.plotly_chart(fig, use_container_width=True)

            elif compare_aspect == "Land Use":
                rows = [
                    ("Forest %", "forest_land_pct"),
                    ("Grazing %", "grazing_land_pct"),
                    ("Community Conserved %", "community_conserved_area_pct"),
                    ("Agricultural %", "agricultural_land_pct"),
                    ("Other %", "other_land_pct"),
                    ("Total Area (ha)", "total_area_ha"),
                ]
                table_data = {"Field": [r[0] for r in rows]}
                for rec in compare_records:
                    table_data[rec.get("village_name", "?")] = [str(rec.get(f, "") or "") for _, f in rows]
                st.dataframe(pd.DataFrame(table_data).set_index("Field"), use_container_width=True)

            elif compare_aspect == "Water & Livestock":
                rows = [
                    ("Drinking Water", "drinking_water_sources"),
                    ("Livestock Water", "livestock_water_sources"),
                    ("Irrigation", "irrigation_sources"),
                    ("Livestock Summary", "livestock_summary"),
                    ("Indigenous Breeds", "indigenous_breeds"),
                ]
                table_data = {"Field": [r[0] for r in rows]}
                for rec in compare_records:
                    table_data[rec.get("village_name", "?")] = [str(rec.get(f, "") or "")[:150] for _, f in rows]
                st.dataframe(pd.DataFrame(table_data).set_index("Field"), use_container_width=True)

            elif compare_aspect == "All Fields":
                all_fields = [k for k in MASTER_FIELDS.keys()]
                table_data = {"Field": all_fields}
                for rec in compare_records:
                    table_data[rec.get("village_name", "?")] = [str(rec.get(f, "") or "")[:100] for f in all_fields]
                df_compare = pd.DataFrame(table_data).set_index("Field")
                # Only show rows where at least one village has data
                df_compare = df_compare[df_compare.apply(lambda row: any(v.strip() and v.strip() != "0" for v in row), axis=1)]
                st.dataframe(df_compare, use_container_width=True, height=600)
        else:
            st.caption("Select at least 2 villages to compare.")

    # ── PDF Report Generation ───────────────────────────────
    st.divider()
    st.markdown("### Generate Report")

    report_col1, report_col2 = st.columns([2, 1])
    with report_col1:
        report_villages = st.multiselect(
            "Select villages for report",
            options=[r.get("village_name", f"Village {i+1}") for i, r in enumerate(records)],
            default=[records[0].get("village_name", "Village 1")] if records else [],
            key="report_villages",
        )
    with report_col2:
        report_format = st.radio("Format", ["Summary", "Detailed"], horizontal=True, key="report_format")

    if report_villages and st.button("Generate Report", type="primary", use_container_width=True):
        report_recs = [r for r in records if r.get("village_name") in report_villages]

        report_lines = []
        report_lines.append("=" * 60)
        report_lines.append("VER DATA EXTRACTION REPORT")
        report_lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        report_lines.append(f"Villages: {len(report_recs)}")
        report_lines.append("=" * 60)

        for rec in report_recs:
            name = rec.get("village_name", "Unknown")
            report_lines.append("")
            report_lines.append("-" * 60)
            report_lines.append(f"VILLAGE: {name}")
            report_lines.append("-" * 60)

            # General
            report_lines.append(f"  State: {rec.get('state', 'N/A')}")
            report_lines.append(f"  Block: {rec.get('block', 'N/A')}")
            report_lines.append(f"  GPS: {rec.get('latitude', 'N/A')}, {rec.get('longitude', 'N/A')}")
            report_lines.append(f"  Survey Date: {rec.get('date_of_survey', 'N/A')}")
            report_lines.append(f"  Total Area: {rec.get('total_area_ha', 'N/A')} ha")
            report_lines.append(f"  Population: {rec.get('total_population', 'N/A')}")
            report_lines.append(f"  Households: {rec.get('total_households', 'N/A')}")

            # Land Use
            report_lines.append("")
            report_lines.append("  LAND USE:")
            for label, field in [("Forest", "forest_land_pct"), ("Grazing", "grazing_land_pct"),
                                  ("Community Conserved", "community_conserved_area_pct"),
                                  ("Agricultural", "agricultural_land_pct"), ("Other", "other_land_pct")]:
                val = rec.get(field, "")
                if val:
                    report_lines.append(f"    {label}: {val}%")

            # Biodiversity
            report_lines.append("")
            report_lines.append(f"  BIODIVERSITY (Total Species: {rec.get('total_species_count', 0)}):")
            for label, field in [("Trees", "tree_diversity_count"), ("Shrubs", "shrub_diversity_count"),
                                  ("Herbs & Grasses", "herb_grass_diversity_count"), ("Mammals", "mammal_count"),
                                  ("Birds", "bird_count"), ("Reptiles & Amphibians", "reptile_amphibian_count"),
                                  ("Butterflies", "butterfly_count"), ("Dragonflies", "dragonfly_count")]:
                val = rec.get(field, 0)
                if val:
                    report_lines.append(f"    {label}: {val}")

            # Water & Livestock
            report_lines.append("")
            report_lines.append("  WATER SOURCES:")
            if rec.get("drinking_water_sources"):
                report_lines.append(f"    Drinking: {rec['drinking_water_sources']}")
            if rec.get("livestock_water_sources"):
                report_lines.append(f"    Livestock: {rec['livestock_water_sources']}")
            if rec.get("irrigation_sources"):
                report_lines.append(f"    Irrigation: {rec['irrigation_sources']}")

            report_lines.append("")
            report_lines.append("  LIVESTOCK:")
            if rec.get("livestock_summary"):
                report_lines.append(f"    Summary: {rec['livestock_summary']}")

            if report_format == "Detailed":
                # Agriculture
                report_lines.append("")
                report_lines.append("  AGRICULTURE:")
                for label, field in [("Kharif Crops", "kharif_crops"), ("Rabi Crops", "rabi_crops"),
                                      ("Zaid Crops", "zaid_crops"), ("Traditional Varieties", "traditional_crop_varieties"),
                                      ("Soil Type", "soil_type"), ("Farming Practices", "farming_practices")]:
                    val = rec.get(field, "")
                    if val:
                        report_lines.append(f"    {label}: {val[:200]}")

                # Forest
                report_lines.append("")
                report_lines.append("  FOREST:")
                for label, field in [("Name", "forest_name"), ("Type", "forest_type"),
                                      ("Size", "forest_size_ha"), ("Location", "forest_location_geocode")]:
                    val = rec.get(field, "")
                    if val:
                        report_lines.append(f"    {label}: {val[:200]}")

                # Conservation
                report_lines.append("")
                report_lines.append("  CONSERVATION:")
                for label, field in [("Sacred Groves", "sacred_groves"), ("Conservation Ethos", "conservation_ethos"),
                                      ("Medicinal Plants", "medicinal_plants"), ("Protected Species", "protected_species")]:
                    val = rec.get(field, "")
                    if val:
                        report_lines.append(f"    {label}: {val[:300]}")

                # History
                if rec.get("village_history_narrative"):
                    report_lines.append("")
                    report_lines.append("  VILLAGE HISTORY:")
                    report_lines.append(f"    {rec['village_history_narrative'][:500]}")

        # Summary stats
        report_lines.append("")
        report_lines.append("=" * 60)
        report_lines.append("SUMMARY STATISTICS")
        report_lines.append("=" * 60)
        report_lines.append(f"  Total Villages: {len(report_recs)}")
        report_lines.append(f"  Total Species: {sum(r.get('total_species_count', 0) for r in report_recs)}")
        report_lines.append(f"  Total Population: {sum(int(r.get('total_population', 0) or 0) for r in report_recs)}")
        states = set(r.get("state", "") for r in report_recs if r.get("state"))
        report_lines.append(f"  States Covered: {', '.join(sorted(states)) if states else 'N/A'}")
        report_lines.append("")
        report_lines.append("--- End of Report ---")

        report_text = "\n".join(report_lines)

        st.text_area("Report Preview", report_text, height=400, disabled=True)
        st.download_button(
            label="📥 Download Report (.txt)",
            data=report_text,
            file_name=f"VER_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
            mime="text/plain",
            use_container_width=True,
        )

    # ── AI-Powered Analysis ─────────────────────────────────
    st.divider()
    st.markdown("### AI-Powered Analysis")

    ai_village_names = [r.get("village_name", f"Village {i+1}") for i, r in enumerate(records)]
    ai_selected = st.selectbox("Select village for analysis", ai_village_names, key="ai_village")
    ai_rec = records[ai_village_names.index(ai_selected)]

    ai_type = st.radio(
        "Analysis type",
        ["Ecological Summary", "Conservation Priority", "Biodiversity Health", "Recommendations"],
        horizontal=True,
        key="ai_type",
    )

    # Build context for the AI prompt
    def build_village_context(rec):
        """Build a concise text summary of village data for AI analysis."""
        parts = []
        parts.append(f"Village: {rec.get('village_name', 'Unknown')}, State: {rec.get('state', 'N/A')}")
        parts.append(f"Population: {rec.get('total_population', 'N/A')}, Households: {rec.get('total_households', 'N/A')}")
        parts.append(f"Total Area: {rec.get('total_area_ha', 'N/A')} ha")
        parts.append(f"Forest: {rec.get('forest_land_pct', 'N/A')}%, Agricultural: {rec.get('agricultural_land_pct', 'N/A')}%")
        parts.append(f"Total Species: {rec.get('total_species_count', 0)}")
        parts.append(f"Trees: {rec.get('tree_diversity_count', 0)}, Birds: {rec.get('bird_count', 0)}, Mammals: {rec.get('mammal_count', 0)}")
        parts.append(f"Butterflies: {rec.get('butterfly_count', 0)}, Dragonflies: {rec.get('dragonfly_count', 0)}")
        if rec.get("livestock_summary"):
            parts.append(f"Livestock: {rec['livestock_summary']}")
        if rec.get("drinking_water_sources"):
            parts.append(f"Water Sources: {rec['drinking_water_sources']}")
        if rec.get("conservation_ethos"):
            parts.append(f"Conservation: {rec['conservation_ethos'][:300]}")
        if rec.get("sacred_groves"):
            parts.append(f"Sacred Groves: {rec['sacred_groves'][:200]}")
        if rec.get("medicinal_plants"):
            parts.append(f"Medicinal Plants: {rec['medicinal_plants'][:200]}")
        if rec.get("invasive_plants"):
            parts.append(f"Invasive Plants: {rec['invasive_plants'][:200]}")
        if rec.get("forest_name"):
            parts.append(f"Forest: {rec['forest_name']}, Size: {rec.get('forest_size_ha', 'N/A')}")
        if rec.get("village_history_narrative"):
            parts.append(f"History: {rec['village_history_narrative'][:300]}")
        return "\n".join(parts)

    ai_prompts = {
        "Ecological Summary": "Provide a concise ecological summary of this village. Cover biodiversity richness, key habitats, water resources, and overall ecological significance. Keep it to 3-4 paragraphs.",
        "Conservation Priority": "Assess the conservation priority of this village. Consider species diversity, forest coverage, sacred groves, traditional practices, and any threats. Classify as High/Medium/Low priority with justification.",
        "Biodiversity Health": "Analyze the biodiversity health of this village. Comment on species counts across groups (trees, birds, mammals, butterflies), any notable patterns, potential indicator species, and overall ecosystem balance.",
        "Recommendations": "Based on the village data, provide 5-7 specific, actionable conservation and sustainable development recommendations. Consider biodiversity, water resources, agriculture, livestock, and community practices.",
    }

    # Check for API key
    api_key = st.text_input("Anthropic API Key", type="password", help="Enter your Anthropic API key to use Claude for analysis. Key is not stored.")

    if st.button("Analyze", type="primary", use_container_width=True, key="ai_analyze"):
        if not api_key:
            st.warning("Please enter your Anthropic API key above.")
        else:
            village_context = build_village_context(ai_rec)
            prompt = f"""You are an ecologist analyzing Village Ecological Register (VER) data from India.

Village Data:
{village_context}

Task: {ai_prompts[ai_type]}"""

            with st.spinner("Analyzing with Claude..."):
                try:
                    import anthropic
                    client = anthropic.Anthropic(api_key=api_key)
                    message = client.messages.create(
                        model="claude-sonnet-4-20250514",
                        max_tokens=1024,
                        messages=[{"role": "user", "content": prompt}],
                    )
                    result = message.content[0].text
                    st.markdown(f"**{ai_type} — {ai_rec.get('village_name', '')}**")
                    st.markdown(result)
                except ImportError:
                    st.error("The `anthropic` package is not installed. Add it to requirements.txt or run: `pip install anthropic`")
                except Exception as e:
                    st.error(f"Analysis failed: {e}")
