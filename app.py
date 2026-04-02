"""
VER Data Portal — Village Ecological Register Insights Dashboard
A single-page tabbed portal for exploring, visualizing, and managing
VER data extracted from PDFs across Indian villages.

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

sys.path.insert(0, str(Path(__file__).parent))
from comprehensive_extract import extract_village, get_empty_record, MASTER_FIELDS, SUPPORTED_LANGUAGES
from github_db import (
    load_all_villages, upsert_village,
    import_villages, sync_to_github, get_village_count,
)
import plotly.graph_objects as go
import plotly.express as px

try:
    import folium
    from streamlit_folium import st_folium
    HAS_FOLIUM = True
except ImportError:
    HAS_FOLIUM = False

# ── Page config ─────────────────────────────────────────────
st.set_page_config(
    page_title="VER Data Portal",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Styling ─────────────────────────────────────────────────
st.markdown("""
<style>
    .block-container { padding-top: 4rem !important; }
    .portal-title {
        font-size: 1.5rem; font-weight: 700; color: #2d6a4f;
        margin: 0; line-height: 1.2;
    }
    .portal-subtitle { font-size: 0.8rem; color: #666; margin: 0 0 0.3rem 0; }
    .stat-row { display: flex; gap: 0.5rem; margin: 0.3rem 0 0.5rem 0; flex-wrap: wrap; }
    .stat-card {
        background: linear-gradient(135deg, #f0fdf4, #dcfce7);
        border: 1px solid #bbf7d0; border-radius: 8px;
        padding: 0.4rem 0.6rem; text-align: center; flex: 1; min-width: 100px;
    }
    .stat-number { font-size: 1.3rem; font-weight: 700; color: #166534; line-height: 1.3; }
    .stat-label { font-size: 0.65rem; color: #555; text-transform: uppercase; letter-spacing: 0.5px; }
    .stTabs [data-baseweb="tab-list"] { gap: 0; }
    .stTabs [data-baseweb="tab"] { padding: 0.5rem 1.25rem; font-weight: 600; }
    iframe { border: none !important; }
    .welcome-card {
        background: linear-gradient(135deg, #ecfdf5, #d1fae5);
        border: 1px solid #a7f3d0; border-radius: 12px;
        padding: 2rem; text-align: center; margin: 1rem 0;
    }
    .welcome-title { font-size: 1.8rem; font-weight: 700; color: #065f46; margin-bottom: 0.5rem; }
    .welcome-text { font-size: 1rem; color: #047857; }
    .section-hdr {
        font-size: 1rem; font-weight: 600; color: #2d6a4f;
        border-bottom: 2px solid #bbf7d0; padding-bottom: 0.25rem;
        margin: 0.75rem 0 0.5rem 0;
    }
    .village-pill {
        background: #f0fdf4; border: 1px solid #d1fae5; border-radius: 6px;
        padding: 0.3rem 0.6rem; margin: 0.2rem 0; font-size: 0.8rem;
    }
</style>
""", unsafe_allow_html=True)

# ── GitHub config from secrets ──────────────────────────────
def _get_github_config():
    """Get GitHub token and repo from Streamlit secrets."""
    try:
        token = st.secrets.get("github", {}).get("token", "")
        repo = st.secrets.get("github", {}).get("repo", "")
        return token, repo
    except Exception:
        return "", ""

GH_TOKEN, GH_REPO = _get_github_config()

# ── Session state ───────────────────────────────────────────
if "extracted_data" not in st.session_state:
    st.session_state.extracted_data = load_all_villages()
if "processing" not in st.session_state:
    st.session_state.processing = False


# ── Helper functions ────────────────────────────────────────
def _display_clean(text):
    """Clean text for display — fix broken Unicode arrows/symbols."""
    import re as _re
    s = str(text)
    # Replace arrow symbols with readable text
    s = s.replace('\u2191', ' Up ').replace('\u2193', ' Down ').replace('\u2194', ' Stable ')
    s = s.replace('↑', ' Up ').replace('↓', ' Down ').replace('↔', ' Stable ')
    # Remove common broken/banned Unicode symbols
    s = _re.sub(r'[\ufffd\ufffe\uffff]', '', s)  # replacement chars
    s = _re.sub(r'[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]', '', s)  # control chars
    s = _re.sub(r'[\u2400-\u243f]', '', s)  # control pictures (look like banned icons)
    s = _re.sub(r'[\u2300-\u23ff]', '', s)  # misc technical symbols
    s = _re.sub(r'[\u2b00-\u2bff]', '', s)  # misc symbols and arrows
    s = _re.sub(r'[\u25a0-\u25ff]', '', s)  # geometric shapes
    s = _re.sub(r'[\u2600-\u26ff]', '', s)  # misc symbols (includes ⊘ ⛔ etc)
    s = _re.sub(r'[\u2190-\u21ff]', lambda m: {'↑':' Up ','↓':' Down ','↔':' Stable ','→':' > ','←':' < '}.get(m.group(), ''), s)  # arrows block
    s = _re.sub(r'\s+', ' ', s).strip()
    return s


def _safe_int(val):
    """Safely convert a value to int, extracting leading digits if needed."""
    if isinstance(val, (int, float)):
        return int(val)
    try:
        return int(str(val).strip())
    except (ValueError, TypeError):
        import re
        m = re.match(r'(\d+)', str(val).strip())
        return int(m.group(1)) if m else 0


def admin_hierarchy_filter(records, key_prefix=""):
    """Render cascading State > District > Block > Village dropdowns. Returns filtered list."""
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        all_states = sorted(set(r.get("state", "") for r in records if r.get("state")))
        sel_state = st.selectbox("State", ["All"] + all_states, key=f"{key_prefix}_state")
    state_recs = [r for r in records if r.get("state") == sel_state] if sel_state != "All" else records
    with c2:
        all_districts = sorted(set(r.get("district", "") for r in state_recs if r.get("district")))
        sel_district = st.selectbox("District", ["All"] + all_districts, key=f"{key_prefix}_district")
    dist_recs = [r for r in state_recs if r.get("district") == sel_district] if sel_district != "All" else state_recs
    with c3:
        all_blocks = sorted(set(r.get("block", "") for r in dist_recs if r.get("block")))
        sel_block = st.selectbox("Block", ["All"] + all_blocks, key=f"{key_prefix}_block")
    block_recs = [r for r in dist_recs if r.get("block") == sel_block] if sel_block != "All" else dist_recs
    with c4:
        all_villages = sorted(set(r.get("village_name", "") for r in block_recs if r.get("village_name")))
        sel_village = st.selectbox("Village", ["All"] + all_villages, key=f"{key_prefix}_village")
    if sel_village != "All":
        return [r for r in block_recs if r.get("village_name") == sel_village]
    return block_recs


def generate_excel(records: list[dict]) -> bytes:
    """Generate Excel file from list of village records."""
    if not records:
        return b""
    df = pd.DataFrame(records)
    ordered_cols = [c for c in MASTER_FIELDS.keys() if c in df.columns]
    extra_cols = [c for c in df.columns if c not in ordered_cols]
    df = df[ordered_cols + extra_cols]

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="VER Master Data", index=False)
        workbook = writer.book
        worksheet = writer.sheets["VER Master Data"]
        header_fmt = workbook.add_format({
            "bold": True, "bg_color": "#2d6a4f", "font_color": "white",
            "border": 1, "text_wrap": True, "valign": "top",
        })
        for col_idx, col_name in enumerate(df.columns):
            worksheet.write(0, col_idx, col_name, header_fmt)
        for i, col in enumerate(df.columns):
            max_len = max(len(str(col)), df[col].astype(str).str.len().max())
            worksheet.set_column(i, i, min(max(max_len * 0.9, 12), 50))
        worksheet.freeze_panes(1, 2)

        # Overview sheet
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
                "Total Species": rec.get("total_species_count", 0),
                "Trees": rec.get("tree_diversity_count", 0),
                "Birds": rec.get("bird_count", 0),
                "Mammals": rec.get("mammal_count", 0),
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
    if not records:
        return ""
    df = pd.DataFrame(records)
    ordered_cols = [c for c in MASTER_FIELDS.keys() if c in df.columns]
    return df[ordered_cols].to_csv(index=False)


def generate_geojson(records: list[dict]) -> str:
    features = []
    for rec in records:
        lat, lon = rec.get("latitude"), rec.get("longitude")
        geometry = None
        if lat and lon:
            try:
                geometry = {"type": "Point", "coordinates": [float(lon), float(lat)]}
            except (ValueError, TypeError):
                pass
        properties = OrderedDict()
        for k, v in rec.items():
            if k in ("latitude", "longitude") or k.startswith("_"):
                continue
            if v and v != 0:
                properties[k] = v
        features.append({"type": "Feature", "geometry": geometry, "properties": properties})
    return json.dumps({"type": "FeatureCollection", "name": "VER_Master_Data", "features": features},
                       indent=2, ensure_ascii=False, default=str)


def build_village_context(rec):
    """Build a concise text summary of village data for AI analysis."""
    parts = [
        f"Village: {rec.get('village_name', 'Unknown')}, State: {rec.get('state', 'N/A')}",
        f"Population: {rec.get('total_population', 'N/A')}, Households: {rec.get('total_households', 'N/A')}",
        f"Total Area: {rec.get('total_area_ha', 'N/A')} ha",
        f"Forest: {rec.get('forest_land_pct', 'N/A')}%, Agricultural: {rec.get('agricultural_land_pct', 'N/A')}%",
        f"Total Species: {rec.get('total_species_count', 0)}",
        f"Trees: {rec.get('tree_diversity_count', 0)}, Birds: {rec.get('bird_count', 0)}, Mammals: {rec.get('mammal_count', 0)}",
        f"Butterflies: {rec.get('butterfly_count', 0)}, Dragonflies: {rec.get('dragonfly_count', 0)}",
    ]
    for field, label in [
        ("livestock_summary", "Livestock"), ("drinking_water_sources", "Water Sources"),
        ("conservation_ethos", "Conservation"), ("sacred_groves", "Sacred Groves"),
        ("medicinal_plants", "Medicinal Plants"), ("invasive_plants", "Invasive Plants"),
        ("village_history_narrative", "History"),
    ]:
        if rec.get(field):
            parts.append(f"{label}: {str(rec[field])[:300]}")
    if rec.get("forest_name"):
        parts.append(f"Forest: {rec['forest_name']}, Size: {rec.get('forest_size_ha', 'N/A')}")
    return "\n".join(parts)


# ════════════════════════════════════════════════════════════
# HEADER — Always visible: title + stats
# ════════════════════════════════════════════════════════════
records = st.session_state.extracted_data

# Header
st.markdown('<p class="portal-title">VER Data Portal</p><p class="portal-subtitle">Village Ecological Register — Insights Dashboard</p>', unsafe_allow_html=True)

st.caption(f"{len(records)} village(s) in database" if records else "No villages yet")


# ── Sidebar — minimal village list ──────────────────────────
with st.sidebar:
    st.markdown("### Villages")
    if records:
        for i, rec in enumerate(records):
            name = rec.get("village_name", f"Village {i+1}")
            state = rec.get("state", "")
            species = rec.get("total_species_count", 0)
            method = rec.get("extraction_method", "")
            icon = "📄" if method == "native_text" else "🔍"
            st.markdown(f'<div class="village-pill">{icon} <b>{name}</b> ({state}) — {species} spp</div>', unsafe_allow_html=True)
    else:
        st.caption("No villages yet. Go to **Manage Data** tab to upload PDFs.")

    st.divider()
    st.caption(f"**114** fields per village | **20** VER sections")
    st.caption("Languages: EN, HI, OR, TA, TE, KN, MR, GU")
    if GH_TOKEN:
        st.caption("GitHub sync: Connected")
    else:
        st.caption("GitHub sync: Not configured")


# ════════════════════════════════════════════════════════════
# MAIN TABS
# ════════════════════════════════════════════════════════════
tab_dashboard, tab_explore, tab_charts, tab_reports, tab_manage = st.tabs([
    "\U0001f3e0 Dashboard", "\U0001f50d Explore", "\U0001f4ca Charts", "\U0001f4cb Reports & AI", "\U0001f4c2 Manage Data"
])


# ────────────────────────────────────────────────────────────
# TAB 1: DASHBOARD
# ────────────────────────────────────────────────────────────
with tab_dashboard:
    # ── India Map — always visible ──
    if HAS_FOLIUM:
        st.markdown('<div class="section-hdr">Village Map — India</div>', unsafe_allow_html=True)

        try:
            # Center on villages if available, otherwise center on India
            villages_with_coords = [r for r in records if r.get("latitude") and r.get("longitude")] if records else []
            lats, lons = [], []
            for r in villages_with_coords:
                try:
                    lats.append(float(r["latitude"]))
                    lons.append(float(r["longitude"]))
                except (ValueError, TypeError):
                    continue

            # Always show India extent on load
            m = folium.Map(
                location=[22.5, 82.0],
                zoom_start=5, tiles="OpenStreetMap",
                min_zoom=4,
            )

            # India Country Boundary overlay
            # Source: Survey of India via ESRI India Living Atlas (IAB_Country_2024)
            try:
                india_geojson_path = Path(__file__).parent / "data" / "india_boundary.geojson"
                if india_geojson_path.exists():
                    with open(india_geojson_path, "r", encoding="utf-8") as _f:
                        india_geojson = json.load(_f)
                    folium.GeoJson(
                        india_geojson,
                        name="India Boundary (Survey of India)",
                        style_function=lambda x: {
                            "fillColor": "#d1fae5",
                            "color": "#065f46",
                            "weight": 2.5,
                            "fillOpacity": 0.1,
                        },
                    ).add_to(m)
            except Exception:
                pass  # Boundary overlay is non-critical — map still works without it

            # Village markers
            for r in villages_with_coords:
                try:
                    lat, lon = float(r["latitude"]), float(r["longitude"])
                except (ValueError, TypeError):
                    continue
                name = r.get("village_name", "Unknown")
                species = r.get("total_species_count", 0)
                pop = r.get("total_population", "N/A")
                color = "darkgreen" if _safe_int(species) >= 200 else "green" if _safe_int(species) >= 100 else "orange" if _safe_int(species) >= 50 else "red"
                popup_html = f"""
                <div style="font-family:sans-serif;min-width:180px">
                    <h4 style="margin:0;color:#2d6a4f">{name}</h4>
                    <p style="margin:2px 0;color:#555">{r.get('state','')}</p>
                    <hr style="margin:4px 0">
                    <b>Population:</b> {pop}<br>
                    <b>Species:</b> {species}<br>
                    <b>Trees:</b> {r.get('tree_diversity_count',0)} |
                    <b>Birds:</b> {r.get('bird_count',0)} |
                    <b>Mammals:</b> {r.get('mammal_count',0)}
                </div>"""
                folium.Marker(
                    location=[lat, lon],
                    popup=folium.Popup(popup_html, max_width=300),
                    tooltip=f"{name} ({species} spp)",
                    icon=folium.Icon(color=color, icon="leaf", prefix="fa"),
                ).add_to(m)

            folium.LayerControl(collapsed=True).add_to(m)

            # Legend
            legend_html = """
            <div style="position:fixed;bottom:30px;left:30px;z-index:1000;
                        background:white;padding:8px 12px;border-radius:5px;
                        border:1px solid #ccc;font-size:12px">
                <b>Species Count</b><br>
                <i class="fa fa-leaf" style="color:darkgreen"></i> 200+&nbsp;
                <i class="fa fa-leaf" style="color:green"></i> 100+&nbsp;
                <i class="fa fa-leaf" style="color:orange"></i> 50+&nbsp;
                <i class="fa fa-leaf" style="color:red"></i> &lt;50
            </div>"""
            m.get_root().html.add_child(folium.Element(legend_html))
            st_folium(m, height=450, use_container_width=True, returned_objects=[])

            if not villages_with_coords:
                st.caption("Upload VER PDFs in the **Manage Data** tab to see village markers on the map.")

        except Exception as e:
            st.warning(f"Map could not be rendered: {e}")

    if not records:
        st.markdown("""
        <div class="welcome-card">
            <div class="welcome-title">Welcome to the VER Data Portal</div>
            <div class="welcome-text">
                Upload Village Ecological Register PDFs to get started.<br>
                Go to the <b>Manage Data</b> tab to upload your first PDF.<br><br>
                The map above will show village locations once data is loaded.
            </div>
        </div>
        """, unsafe_allow_html=True)
    else:
        # ── Admin Hierarchy Filter ──
        st.markdown('<div class="section-hdr">Select Village</div>', unsafe_allow_html=True)
        dash_records = admin_hierarchy_filter(records, key_prefix="dash")

        # ── Stat cards — reflect filtered selection ──
        d_species = sum(_safe_int(r.get("total_species_count", 0)) for r in dash_records)
        d_pop = sum(_safe_int(r.get("total_population", 0)) for r in dash_records)
        d_states = len(set(r.get("state", "") for r in dash_records if r.get("state")))
        d_area = sum(_safe_int(r.get("total_area_ha", 0)) for r in dash_records)
        mc1, mc2, mc3, mc4, mc5 = st.columns(5)
        mc1.metric("Villages", len(dash_records))
        mc2.metric("Species", d_species)
        mc3.metric("Population", f"{d_pop:,}")
        mc4.metric("Area (ha)", d_area)
        mc5.metric("States", d_states)

        # ── Key Insights (two columns) ──
        dash_left, dash_right = st.columns(2)

        with dash_left:
            # ── Village Overview Table ──
            st.markdown('<div class="section-hdr">Village Overview</div>', unsafe_allow_html=True)
            overview_data = []
            for rec in dash_records:
                overview_data.append({
                    "Village": rec.get("village_name", ""),
                    "State": rec.get("state", ""),
                    "Population": _safe_int(rec.get("total_population", 0)),
                    "Area (ha)": rec.get("total_area_ha", ""),
                    "Species": _safe_int(rec.get("total_species_count", 0)),
                    "Trees": _safe_int(rec.get("tree_diversity_count", 0)),
                    "Birds": _safe_int(rec.get("bird_count", 0)),
                    "Mammals": _safe_int(rec.get("mammal_count", 0)),
                })
            st.dataframe(pd.DataFrame(overview_data), use_container_width=True, hide_index=True, height=min(200 + len(dash_records) * 35, 400))

        with dash_right:
            # ── Land Use Summary ──
            st.markdown('<div class="section-hdr">Land Use Summary</div>', unsafe_allow_html=True)
            land_fields = [("Forest", "forest_land_pct"), ("Grazing", "grazing_land_pct"),
                           ("Community Conserved", "community_conserved_area_pct"),
                           ("Agricultural", "agricultural_land_pct"), ("Other", "other_land_pct")]
            if len(dash_records) == 1:
                rec = dash_records[0]
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
                                 color_discrete_sequence=["#2d6a4f", "#95d5b2", "#b7e4c7", "#e9c46a", "#ccc"])
                    fig.update_layout(height=300, margin=dict(t=10, b=10, l=10, r=10), showlegend=True)
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("No land use data available.")
            else:
                for rec in dash_records:
                    name = rec.get("village_name", "")
                    forest = rec.get("forest_land_pct", "")
                    agri = rec.get("agricultural_land_pct", "")
                    area = rec.get("total_area_ha", "")
                    st.markdown(f"**{name}** — {area} ha | Forest: {forest}% | Agricultural: {agri}%")

        # ── Village History ──
        histories = [(rec.get("village_name", ""), rec.get("village_history_narrative", "")) for rec in dash_records]
        histories = [(n, h) for n, h in histories if h and str(h).strip()]
        if histories:
            st.markdown('<div class="section-hdr">\U0001f4dc Village History</div>', unsafe_allow_html=True)
            for name, history in histories:
                with st.expander(f"**{name}**", expanded=len(histories) == 1):
                    st.markdown(str(history)[:1000])

        # ── Key Ecological Data ──
        st.markdown('<div class="section-hdr">Key Ecological Data</div>', unsafe_allow_html=True)
        eco_left, eco_right = st.columns(2)
        with eco_left:
            water_data = [(rec.get("village_name", ""), rec.get("drinking_water_sources", "")) for rec in dash_records]
            water_data = [(n, w) for n, w in water_data if w and str(w).strip()]
            if water_data:
                st.markdown("**\U0001f4a7 Water Sources**")
                for name, water in water_data:
                    st.markdown(f"- **{name}:** {str(water)[:200]}")

            livestock_data = [(rec.get("village_name", ""), rec.get("livestock_summary", "")) for rec in dash_records]
            livestock_data = [(n, l) for n, l in livestock_data if l and str(l).strip()]
            if livestock_data:
                st.markdown("**\U0001f404 Livestock**")
                for name, livestock in livestock_data:
                    st.markdown(f"- **{name}:** {str(livestock)[:200]}")

        with eco_right:
            # Conservation
            cons_data = [(rec.get("village_name", ""), rec.get("conservation_ethos", "")) for rec in dash_records]
            cons_data = [(n, c) for n, c in cons_data if c and str(c).strip()]
            if cons_data:
                st.markdown("**\U0001f33f Conservation**")
                for name, cons in cons_data:
                    st.markdown(f"- **{name}:** {str(cons)[:200]}")

            grove_data = [(rec.get("village_name", ""), rec.get("sacred_groves", "")) for rec in dash_records]
            grove_data = [(n, g) for n, g in grove_data if g and str(g).strip()]
            if grove_data:
                st.markdown("**\U0001f333 Sacred Groves**")
                for name, grove in grove_data:
                    st.markdown(f"- **{name}:** {str(grove)[:200]}")


# ────────────────────────────────────────────────────────────
# TAB 2: EXPLORE
# ────────────────────────────────────────────────────────────
with tab_explore:
    if not records:
        st.info("No village data loaded yet. Go to **Manage Data** tab to upload PDFs.")
    else:
        # Admin Hierarchy Filter
        st.markdown('<div class="section-hdr">Navigate by Admin Hierarchy</div>', unsafe_allow_html=True)
        filtered = admin_hierarchy_filter(records, key_prefix="explore")

        # Additional filters
        f1, f2 = st.columns([2, 1])
        with f1:
            search_q = st.text_input("Search", placeholder="Crop, species, keyword...", key="explore_search")
        with f2:
            sp_counts = [_safe_int(r.get("total_species_count", 0)) for r in filtered]
            min_sp, max_sp = min(sp_counts, default=0), max(sp_counts, default=500)
            if min_sp < max_sp:
                sp_range = st.slider("Species", min_value=min_sp, max_value=max_sp, value=(min_sp, max_sp), key="explore_sp")
            else:
                sp_range = (min_sp, max_sp)

        if search_q:
            q = search_q.lower()
            filtered = [r for r in filtered if any(q in str(v).lower() for v in r.values())]
        if min_sp < max_sp:
            filtered = [r for r in filtered if sp_range[0] <= _safe_int(r.get("total_species_count", 0)) <= sp_range[1]]

        if len(filtered) < len(records):
            st.caption(f"Showing **{len(filtered)}** of {len(records)} villages")
        if not filtered:
            st.warning("No villages match your filters.")
        else:
            # View mode
            view_mode = st.radio("View", ["Overview", "Village Detail", "Compare"], horizontal=True, key="explore_view")

            if view_mode == "Overview":
                overview_cols = [
                    "village_name", "state", "latitude", "longitude", "total_area_ha",
                    "total_population", "total_households", "major_livelihoods",
                    "livestock_summary", "drinking_water_sources", "total_species_count",
                    "tree_diversity_count", "bird_count", "mammal_count",
                ]
                df = pd.DataFrame(filtered)
                available = [c for c in overview_cols if c in df.columns]
                st.dataframe(df[available], use_container_width=True, height=500, hide_index=True)

            elif view_mode == "Village Detail":
                v_names = [r.get("village_name", f"Village {i+1}") for i, r in enumerate(filtered)]
                selected = st.selectbox("Select Village", v_names, key="explore_detail_sel")
                rec = filtered[v_names.index(selected)]

                detail_tabs = st.tabs([
                    "\U0001f4cb General", "\U0001f4dc History", "\U0001f33e Agriculture", "\U0001f404 Livestock",
                    "\U0001f4a7 Water", "\U0001f332 Forest", "\U0001f98b Biodiversity", "\U0001f33f Conservation", "\U0001f4ca All Fields"
                ])

                with detail_tabs[0]:
                    st.markdown(f"**{rec.get('village_name','')}** | {rec.get('state','')} | Block: {rec.get('block','')}")
                    st.markdown(f"GPS: {rec.get('latitude','')}, {rec.get('longitude','')} | Survey: {rec.get('date_of_survey','')}")
                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown("**Land Use**")
                        for f in ["total_area_ha", "forest_land_pct", "grazing_land_pct", "community_conserved_area_pct", "agricultural_land_pct", "other_land_pct"]:
                            if rec.get(f):
                                st.markdown(f"- {f.replace('_',' ').title()}: **{rec[f]}**")
                    with c2:
                        st.markdown("**Demographics**")
                        for f in ["total_population", "total_households", "large_farmers_gt10ha", "medium_farmers_4_10ha", "small_farmers_1_2ha", "marginal_farmers_lt1ha"]:
                            if rec.get(f):
                                st.markdown(f"- {f.replace('_',' ').title()}: **{rec[f]}**")
                    if rec.get("major_livelihoods"):
                        st.markdown(f"**Livelihoods:** {rec['major_livelihoods']}")

                with detail_tabs[1]:
                    for f in ["village_history_narrative", "myths_and_beliefs", "traditional_songs"]:
                        if rec.get(f):
                            st.markdown(f"**{f.replace('_',' ').title()}**")
                            st.text_area("", _display_clean(rec[f]), height=150, key=f"det_h_{f}_{selected}", disabled=True)

                with detail_tabs[2]:
                    for f in ["kharif_crops", "rabi_crops", "zaid_crops", "traditional_crop_varieties", "farming_practices",
                               "soil_type", "soil_fertility_change", "soil_fertility_reason", "pest_incidences", "major_weeds"]:
                        if rec.get(f):
                            st.markdown(f"**{f.replace('_',' ').title()}:** {_display_clean(rec[f])[:300]}")

                with detail_tabs[3]:
                    if rec.get("livestock_summary"):
                        st.markdown(f"**Summary:** {_display_clean(rec['livestock_summary'])}")
                    if rec.get("livestock_detailed"):
                        st.markdown(f"**Detailed:** {_display_clean(rec['livestock_detailed'])}")
                    for f in ["indigenous_breeds", "livestock_diseases", "ethno_veterinary_practices"]:
                        if rec.get(f):
                            st.text_area(f.replace("_"," ").title(), _display_clean(rec[f])[:500], height=100, key=f"det_l_{f}_{selected}", disabled=True)

                with detail_tabs[4]:
                    for f in ["drinking_water_sources", "livestock_water_sources", "irrigation_sources",
                               "water_quality_changes", "important_water_bodies"]:
                        if rec.get(f):
                            st.markdown(f"**{f.replace('_',' ').title()}:** {_display_clean(rec[f])[:300]}")

                with detail_tabs[5]:
                    for f in ["forest_name", "forest_type", "forest_size_ha", "forest_location_geocode",
                               "forest_tree_species", "forest_shrub_species", "forest_herb_species", "forest_ntfp"]:
                        if rec.get(f):
                            st.markdown(f"**{f.replace('_',' ').title()}:** {_display_clean(rec[f])[:300]}")

                with detail_tabs[6]:
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
                                st.markdown(_display_clean(rec.get(list_f, "")))

                with detail_tabs[7]:
                    for f in ["sacred_groves", "conservation_ethos", "bamboo_species", "medicinal_plants",
                               "invasive_plants", "protected_species", "feral_animals", "fire_incidence"]:
                        if rec.get(f):
                            st.text_area(f.replace("_"," ").title(), _display_clean(rec[f])[:500], height=80, key=f"det_c_{f}_{selected}", disabled=True)

                with detail_tabs[8]:
                    all_data = [(k, str(v)[:200]) for k, v in rec.items() if v and v != 0 and not k.startswith("_")]
                    st.dataframe(pd.DataFrame(all_data, columns=["Field", "Value"]), use_container_width=True, height=600, hide_index=True)

            elif view_mode == "Compare":
                if len(filtered) < 2:
                    st.info("Need at least 2 villages to compare.")
                else:
                    compare_names = [r.get("village_name", f"V{i+1}") for i, r in enumerate(filtered)]
                    selected_compare = st.multiselect(
                        "Select villages", options=compare_names,
                        default=compare_names[:min(3, len(compare_names))], key="explore_compare",
                    )
                    if len(selected_compare) >= 2:
                        compare_recs = [filtered[compare_names.index(n)] for n in selected_compare]
                        aspect = st.radio("Compare by", ["Overview", "Biodiversity", "Land Use", "Water & Livestock", "All Fields"],
                                         horizontal=True, key="compare_aspect")

                        if aspect == "Overview":
                            rows = [("State","state"),("Population","total_population"),("Households","total_households"),
                                    ("Area (ha)","total_area_ha"),("Total Species","total_species_count"),("Livelihoods","major_livelihoods"),("GPS",None)]
                            table_data = {"Field": [r[0] for r in rows]}
                            for rec in compare_recs:
                                vals = []
                                for label, field in rows:
                                    if field is None:
                                        vals.append(f"{rec.get('latitude','')}, {rec.get('longitude','')}")
                                    else:
                                        vals.append(str(rec.get(field, "") or ""))
                                table_data[rec.get("village_name","?")] = vals
                            st.dataframe(pd.DataFrame(table_data).set_index("Field"), use_container_width=True)

                        elif aspect == "Biodiversity":
                            bio_cats = [("Trees","tree_diversity_count"),("Shrubs","shrub_diversity_count"),
                                        ("Herbs","herb_grass_diversity_count"),("Mammals","mammal_count"),
                                        ("Birds","bird_count"),("Reptiles","reptile_amphibian_count"),
                                        ("Butterflies","butterfly_count"),("Dragonflies","dragonfly_count")]
                            fig = go.Figure()
                            for rec in compare_recs:
                                fig.add_trace(go.Bar(name=rec.get("village_name","?"),
                                                     x=[c[0] for c in bio_cats],
                                                     y=[_safe_int(rec.get(f,0)) for _,f in bio_cats]))
                            fig.update_layout(barmode="group", title="Biodiversity Comparison",
                                              yaxis_title="Species Count", height=400,
                                              margin=dict(t=40,b=40), legend=dict(orientation="h",y=1.1))
                            st.plotly_chart(fig, use_container_width=True)

                        elif aspect == "Land Use":
                            rows = [("Forest %","forest_land_pct"),("Grazing %","grazing_land_pct"),
                                    ("Community Conserved %","community_conserved_area_pct"),
                                    ("Agricultural %","agricultural_land_pct"),("Other %","other_land_pct"),
                                    ("Total Area (ha)","total_area_ha")]
                            table_data = {"Field": [r[0] for r in rows]}
                            for rec in compare_recs:
                                table_data[rec.get("village_name","?")] = [str(rec.get(f,"") or "") for _,f in rows]
                            st.dataframe(pd.DataFrame(table_data).set_index("Field"), use_container_width=True)

                        elif aspect == "Water & Livestock":
                            rows = [("Drinking Water","drinking_water_sources"),("Livestock Water","livestock_water_sources"),
                                    ("Irrigation","irrigation_sources"),("Livestock","livestock_summary"),
                                    ("Indigenous Breeds","indigenous_breeds")]
                            table_data = {"Field": [r[0] for r in rows]}
                            for rec in compare_recs:
                                table_data[rec.get("village_name","?")] = [str(rec.get(f,"") or "")[:150] for _,f in rows]
                            st.dataframe(pd.DataFrame(table_data).set_index("Field"), use_container_width=True)

                        elif aspect == "All Fields":
                            all_fields = list(MASTER_FIELDS.keys())
                            table_data = {"Field": all_fields}
                            for rec in compare_recs:
                                table_data[rec.get("village_name","?")] = [str(rec.get(f,"") or "")[:100] for f in all_fields]
                            df_cmp = pd.DataFrame(table_data).set_index("Field")
                            df_cmp = df_cmp[df_cmp.apply(lambda row: any(v.strip() and v.strip() != "0" for v in row), axis=1)]
                            st.dataframe(df_cmp, use_container_width=True, height=600)
                    else:
                        st.caption("Select at least 2 villages to compare.")


# ────────────────────────────────────────────────────────────
# TAB 3: CHARTS
# ────────────────────────────────────────────────────────────
with tab_charts:
    if not records:
        st.info("No village data loaded yet. Go to **Manage Data** tab to upload PDFs.")
    else:
        chart_sub = st.tabs(["\U0001f98b Biodiversity", "\U0001f30d Land Use", "\U0001f465 Demographics", "\U0001f33e Agriculture"])

        with chart_sub[0]:
            village_names_c = [r.get("village_name", f"V{i+1}") for i, r in enumerate(records)]
            bio_cats = [
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
            for label, field, color in bio_cats:
                values = [_safe_int(r.get(field, 0)) for r in records]
                if any(v > 0 for v in values):
                    fig.add_trace(go.Bar(name=label, x=village_names_c, y=values, marker_color=color))
            fig.update_layout(
                barmode="stack", title="Biodiversity Breakdown by Village",
                xaxis_title="Village", yaxis_title="Species Count", height=500,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                margin=dict(t=60, b=40),
            )
            st.plotly_chart(fig, use_container_width=True)

        with chart_sub[1]:
            land_fields = [("Forest","forest_land_pct"),("Grazing","grazing_land_pct"),
                           ("Community Conserved","community_conserved_area_pct"),
                           ("Agricultural","agricultural_land_pct"),("Other","other_land_pct")]
            if len(records) == 1:
                rec = records[0]
                pie_labels, pie_values = [], []
                for label, field in land_fields:
                    val = rec.get(field, "")
                    try:
                        num = float(str(val).replace("%","").strip())
                        if num > 0:
                            pie_labels.append(label)
                            pie_values.append(num)
                    except (ValueError, TypeError):
                        pass
                if pie_values:
                    fig = px.pie(names=pie_labels, values=pie_values,
                                 title=f"Land Use — {rec.get('village_name','')}",
                                 color_discrete_sequence=["#2d6a4f","#95d5b2","#b7e4c7","#e9c46a","#ccc"])
                    fig.update_layout(height=450)
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("No land use data available.")
            else:
                fig = go.Figure()
                colors = ["#2d6a4f","#95d5b2","#b7e4c7","#e9c46a","#ccc"]
                for (label, field), color in zip(land_fields, colors):
                    values = []
                    for r in records:
                        try:
                            values.append(float(str(r.get(field,"")).replace("%","").strip()))
                        except (ValueError, TypeError):
                            values.append(0)
                    if any(v > 0 for v in values):
                        fig.add_trace(go.Bar(name=label, x=village_names_c, y=values, marker_color=color))
                fig.update_layout(barmode="group", title="Land Use Comparison (%)",
                                  xaxis_title="Village", yaxis_title="Percentage", height=500,
                                  legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1),
                                  margin=dict(t=60,b=40))
                st.plotly_chart(fig, use_container_width=True)

        with chart_sub[2]:
            pop_data = []
            for r in records:
                pop_data.append({
                    "Village": r.get("village_name","?"),
                    "Population": _safe_int(r.get("total_population",0)),
                    "Households": _safe_int(r.get("total_households",0)),
                })
            if any(d["Population"] > 0 for d in pop_data):
                df_pop = pd.DataFrame(pop_data)
                fig = go.Figure()
                fig.add_trace(go.Bar(name="Population", x=df_pop["Village"], y=df_pop["Population"], marker_color="#2d6a4f"))
                fig.add_trace(go.Bar(name="Households", x=df_pop["Village"], y=df_pop["Households"], marker_color="#e9c46a"))
                fig.update_layout(barmode="group", title="Population & Households",
                                  xaxis_title="Village", yaxis_title="Count", height=450,
                                  legend=dict(orientation="h",y=1.1), margin=dict(t=60,b=40))
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No population data available.")

        with chart_sub[3]:
            crop_types = [("Kharif Crops","kharif_crops"),("Rabi Crops","rabi_crops"),("Zaid Crops","zaid_crops")]
            has_crop = any(r.get(f) for r in records for _,f in crop_types)
            if has_crop:
                for r in records:
                    name = r.get("village_name","Village")
                    st.markdown(f"**{name}**")
                    for label, field in crop_types:
                        if r.get(field):
                            st.markdown(f"- **{label}:** {str(r[field])[:300]}")
                if any(r.get("traditional_crop_varieties") for r in records):
                    st.markdown("---")
                    st.markdown("**Traditional Varieties**")
                    for r in records:
                        if r.get("traditional_crop_varieties"):
                            st.markdown(f"- **{r.get('village_name','?')}:** {str(r['traditional_crop_varieties'])[:300]}")
            else:
                st.info("No agriculture data available.")


# ────────────────────────────────────────────────────────────
# TAB 4: REPORTS & AI
# ────────────────────────────────────────────────────────────
with tab_reports:
    if not records:
        st.info("No village data loaded yet. Go to **Manage Data** tab to upload PDFs.")
    else:
        report_sub = st.tabs(["\U0001f4c4 Generate Report", "\U0001f916 AI Analysis"])

        # ── Report Generation ──
        with report_sub[0]:
            st.markdown('<div class="section-hdr">Generate Report</div>', unsafe_allow_html=True)
            rc1, rc2 = st.columns([2, 1])
            with rc1:
                report_villages = st.multiselect(
                    "Villages", options=[r.get("village_name", f"V{i+1}") for i,r in enumerate(records)],
                    default=[records[0].get("village_name","Village 1")] if records else [],
                    key="report_villages",
                )
            with rc2:
                report_fmt = st.radio("Format", ["Summary","Detailed"], horizontal=True, key="report_fmt")

            if report_villages and st.button("Generate Report", type="primary", use_container_width=True):
                report_recs = [r for r in records if r.get("village_name") in report_villages]
                lines = []
                lines.append("=" * 60)
                lines.append("VER DATA EXTRACTION REPORT")
                lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
                lines.append(f"Villages: {len(report_recs)}")
                lines.append("=" * 60)

                for rec in report_recs:
                    name = rec.get("village_name","Unknown")
                    lines += ["", "-"*60, f"VILLAGE: {name}", "-"*60]
                    lines.append(f"  State: {rec.get('state','N/A')}")
                    lines.append(f"  Block: {rec.get('block','N/A')}")
                    lines.append(f"  GPS: {rec.get('latitude','N/A')}, {rec.get('longitude','N/A')}")
                    lines.append(f"  Survey Date: {rec.get('date_of_survey','N/A')}")
                    lines.append(f"  Total Area: {rec.get('total_area_ha','N/A')} ha")
                    lines.append(f"  Population: {rec.get('total_population','N/A')}")
                    lines.append(f"  Households: {rec.get('total_households','N/A')}")

                    lines += ["", "  LAND USE:"]
                    for label, field in [("Forest","forest_land_pct"),("Grazing","grazing_land_pct"),
                                          ("Community Conserved","community_conserved_area_pct"),
                                          ("Agricultural","agricultural_land_pct"),("Other","other_land_pct")]:
                        if rec.get(field):
                            lines.append(f"    {label}: {rec[field]}%")

                    lines += ["", f"  BIODIVERSITY (Total Species: {rec.get('total_species_count',0)}):"]
                    for label, field in [("Trees","tree_diversity_count"),("Shrubs","shrub_diversity_count"),
                                          ("Herbs & Grasses","herb_grass_diversity_count"),("Mammals","mammal_count"),
                                          ("Birds","bird_count"),("Reptiles","reptile_amphibian_count"),
                                          ("Butterflies","butterfly_count"),("Dragonflies","dragonfly_count")]:
                        if rec.get(field, 0):
                            lines.append(f"    {label}: {rec[field]}")

                    lines += ["", "  WATER SOURCES:"]
                    for label, field in [("Drinking","drinking_water_sources"),("Livestock","livestock_water_sources"),("Irrigation","irrigation_sources")]:
                        if rec.get(field):
                            lines.append(f"    {label}: {rec[field]}")

                    lines += ["", "  LIVESTOCK:"]
                    if rec.get("livestock_summary"):
                        lines.append(f"    Summary: {rec['livestock_summary']}")

                    if report_fmt == "Detailed":
                        lines += ["", "  AGRICULTURE:"]
                        for label, field in [("Kharif","kharif_crops"),("Rabi","rabi_crops"),("Zaid","zaid_crops"),
                                              ("Traditional","traditional_crop_varieties"),("Soil","soil_type"),("Practices","farming_practices")]:
                            if rec.get(field):
                                lines.append(f"    {label}: {str(rec[field])[:200]}")
                        lines += ["", "  FOREST:"]
                        for label, field in [("Name","forest_name"),("Type","forest_type"),("Size","forest_size_ha")]:
                            if rec.get(field):
                                lines.append(f"    {label}: {str(rec[field])[:200]}")
                        lines += ["", "  CONSERVATION:"]
                        for label, field in [("Sacred Groves","sacred_groves"),("Ethos","conservation_ethos"),
                                              ("Medicinal","medicinal_plants"),("Protected","protected_species")]:
                            if rec.get(field):
                                lines.append(f"    {label}: {str(rec[field])[:300]}")
                        if rec.get("village_history_narrative"):
                            lines += ["", "  HISTORY:", f"    {str(rec['village_history_narrative'])[:500]}"]

                lines += ["", "="*60, "SUMMARY STATISTICS", "="*60]
                lines.append(f"  Total Villages: {len(report_recs)}")
                lines.append(f"  Total Species: {sum(_safe_int(r.get('total_species_count',0)) for r in report_recs)}")
                lines.append(f"  Total Population: {sum(_safe_int(r.get('total_population',0)) for r in report_recs)}")
                sts = set(r.get("state","") for r in report_recs if r.get("state"))
                lines.append(f"  States: {', '.join(sorted(sts)) if sts else 'N/A'}")
                lines += ["", "--- End of Report ---"]

                report_text = "\n".join(lines)
                st.text_area("Report Preview", report_text, height=400, disabled=True)
                st.download_button("Download Report (.txt)", data=report_text,
                                   file_name=f"VER_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                                   mime="text/plain", use_container_width=True)

        # ── AI Analysis ──
        with report_sub[1]:
            st.markdown('<div class="section-hdr">AI-Powered Analysis (Gemini)</div>', unsafe_allow_html=True)
            ai_names = [r.get("village_name", f"V{i+1}") for i,r in enumerate(records)]
            ai_sel = st.selectbox("Village", ai_names, key="ai_village")
            ai_rec = records[ai_names.index(ai_sel)]

            ai_type = st.radio("Analysis type",
                ["Ecological Summary","Conservation Priority","Biodiversity Health","Recommendations"],
                horizontal=True, key="ai_type")

            ai_prompts = {
                "Ecological Summary": "Provide a concise ecological summary of this village. Cover biodiversity richness, key habitats, water resources, and overall ecological significance. Keep it to 3-4 paragraphs.",
                "Conservation Priority": "Assess the conservation priority of this village. Consider species diversity, forest coverage, sacred groves, traditional practices, and any threats. Classify as High/Medium/Low priority with justification.",
                "Biodiversity Health": "Analyze the biodiversity health of this village. Comment on species counts across groups, notable patterns, potential indicator species, and overall ecosystem balance.",
                "Recommendations": "Based on the village data, provide 5-7 specific, actionable conservation and sustainable development recommendations.",
            }

            api_key = ""
            try:
                api_key = st.secrets.get("GEMINI_API_KEY", "")
            except Exception:
                pass
            if not api_key:
                api_key = st.text_input("Gemini API Key", type="password",
                    help="Get free key from aistudio.google.com. Or add GEMINI_API_KEY to Streamlit secrets.", key="ai_key")

            if st.button("Analyze", type="primary", use_container_width=True, key="ai_analyze"):
                if not api_key:
                    st.warning("Enter your Gemini API key above, or add GEMINI_API_KEY in Streamlit Cloud secrets.")
                else:
                    context = build_village_context(ai_rec)
                    prompt = f"""You are an ecologist analyzing Village Ecological Register (VER) data from India.

Village Data:
{context}

Task: {ai_prompts[ai_type]}"""
                    with st.spinner("Analyzing with Gemini..."):
                        try:
                            import google.generativeai as genai
                            genai.configure(api_key=api_key)
                            model = genai.GenerativeModel("gemini-2.0-flash-lite")
                            response = model.generate_content(prompt)
                            st.markdown(f"**{ai_type} — {ai_rec.get('village_name','')}**")
                            st.markdown(response.text)
                        except ImportError:
                            st.error("google-generativeai package not installed.")
                        except Exception as e:
                            st.error(f"Analysis failed: {e}")


# ────────────────────────────────────────────────────────────
# TAB 5: MANAGE DATA
# ────────────────────────────────────────────────────────────
with tab_manage:
    manage_sub = st.tabs(["\U0001f4e4 Upload PDFs", "\U0001f4e5 Export Data", "\U0001f504 Backup & Sync"])

    # ── Upload ──
    with manage_sub[0]:
        st.markdown('<div class="section-hdr">Upload VER PDFs</div>', unsafe_allow_html=True)
        st.caption("Upload Village Ecological Register PDFs to extract data. Each PDF will be processed and added to the database.")

        u1, u2 = st.columns([2, 1])
        with u1:
            uploaded_files = st.file_uploader("Upload PDF(s)", type=["pdf"], accept_multiple_files=True,
                help="Upload one or more VER PDF files.", key="pdf_upload")
        with u2:
            selected_language = st.selectbox("PDF Language", options=list(SUPPORTED_LANGUAGES.keys()), index=0, key="lang_sel")
            if selected_language != "Auto-detect":
                lang_info = SUPPORTED_LANGUAGES[selected_language]
                st.caption(f"Script: {lang_info['script']} | OCR: `{lang_info['tesseract']}`")

        if uploaded_files:
            if st.button("Extract Data from PDFs", type="primary", use_container_width=True):
                progress_bar = st.progress(0, text="Starting extraction...")
                status_text = st.empty()

                for file_idx, uploaded_file in enumerate(uploaded_files):
                    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                        tmp.write(uploaded_file.read())
                        tmp_path = tmp.name

                    file_label = uploaded_file.name

                    def progress_cb(step, total, msg, _idx=file_idx, _label=file_label):
                        pct = ((_idx + step / total) / len(uploaded_files))
                        progress_bar.progress(min(pct, 1.0), text=f"[{_idx+1}/{len(uploaded_files)}] {_label}: {msg}")

                    status_text.markdown(f"Processing **{file_label}**...")
                    new_count, update_count = getattr(st.session_state, '_upload_new', 0), getattr(st.session_state, '_upload_update', 0)
                    try:
                        record = extract_village(tmp_path, language=selected_language, progress_callback=progress_cb)
                        if not record.get("village_name"):
                            name_from_file = Path(file_label).stem.replace("VER_","").replace("_"," ")
                            record["village_name"] = name_from_file.split(" ")[0]

                        vid, was_update = upsert_village(record, github_token=GH_TOKEN, github_repo=GH_REPO)
                        if was_update:
                            update_count += 1
                        else:
                            new_count += 1
                    except Exception as e:
                        st.error(f"Error processing {file_label}: {e}")
                    finally:
                        os.unlink(tmp_path)

                progress_bar.progress(1.0, text="All PDFs processed!")
                # Reload from database to ensure session matches persistent store
                st.session_state.extracted_data = load_all_villages()
                parts = []
                if new_count: parts.append(f"{new_count} new village(s)")
                if update_count: parts.append(f"{update_count} updated village(s)")
                status_text.success(f"Processed {len(uploaded_files)} PDF(s): {', '.join(parts) if parts else 'done'}")
                st.rerun()

        # Show current database info
        if records:
            st.divider()
            st.info(f"**{len(records)} village(s)** in the database. Data is append-only — uploading a PDF for an existing village will update its record.")

    # ── Export ──
    with manage_sub[1]:
        st.markdown('<div class="section-hdr">Export Data</div>', unsafe_allow_html=True)
        if not records:
            st.info("No data to export yet.")
        else:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M")
            e1, e2, e3, e4 = st.columns(4)
            with e1:
                st.download_button("Excel (.xlsx)", data=generate_excel(records),
                    file_name=f"VER_Master_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary", use_container_width=True)
            with e2:
                st.download_button("CSV", data=generate_csv(records),
                    file_name=f"VER_Master_{timestamp}.csv", mime="text/csv", use_container_width=True)
            with e3:
                st.download_button("GeoJSON", data=generate_geojson(records),
                    file_name=f"VER_Master_{timestamp}.geojson", mime="application/geo+json", use_container_width=True)
            with e4:
                json_str = json.dumps(records, indent=2, ensure_ascii=False, default=str)
                st.download_button("JSON", data=json_str,
                    file_name=f"VER_Master_{timestamp}.json", mime="application/json", use_container_width=True)

    # ── Backup & Sync ──
    with manage_sub[2]:
        st.markdown('<div class="section-hdr">Backup & Sync</div>', unsafe_allow_html=True)

        # JSON backup export
        if records:
            export_data = json.dumps(
                [{k: v for k, v in r.items() if not k.startswith("_")} for r in records],
                ensure_ascii=False, default=str,
            )
            st.download_button("Export Backup (JSON)", data=export_data,
                file_name=f"VER_backup_{datetime.now().strftime('%Y%m%d')}.json",
                mime="application/json", use_container_width=True)

        # JSON backup import
        imported_file = st.file_uploader("Import Backup (JSON)", type=["json"],
            help="Import a previously exported JSON backup.", key="import_backup")
        if imported_file:
            try:
                imported_records = json.loads(imported_file.read().decode("utf-8"))
                if isinstance(imported_records, list) and len(imported_records) > 0:
                    if st.button(f"Restore {len(imported_records)} village(s)", type="primary", use_container_width=True):
                        import_villages(imported_records, github_token=GH_TOKEN, github_repo=GH_REPO)
                        st.session_state.extracted_data = load_all_villages()
                        st.success(f"Restored {len(imported_records)} village(s)!")
                        st.rerun()
                else:
                    st.warning("Invalid backup — expected a JSON list of village records.")
            except (json.JSONDecodeError, UnicodeDecodeError):
                st.error("Could not read this file. Use a valid JSON backup.")

        # GitHub sync
        st.divider()
        st.markdown("**GitHub Sync**")
        if GH_TOKEN and GH_REPO:
            st.success(f"Connected to `{GH_REPO}`")
            if st.button("Sync to GitHub Now", use_container_width=True):
                with st.spinner("Syncing..."):
                    ok, msg = sync_to_github(GH_TOKEN, GH_REPO)
                if ok:
                    st.success(msg)
                else:
                    st.error(msg)
        else:
            st.markdown("""
            **Not configured.** To enable GitHub sync:
            1. Create a [Fine-grained token](https://github.com/settings/tokens?type=beta) with **Contents: Read & Write** on your repo
            2. In Streamlit Cloud, go to **Settings > Secrets** and add:
            ```toml
            [github]
            token = "github_pat_xxxx"
            repo = "your-username/ver-ocr-tool"
            ```
            This ensures your data persists across redeployments.
            """)
