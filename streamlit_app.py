"""Reading Nation Book Ecosystem — professional dashboard."""

import folium
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
import streamlit as st
from streamlit_folium import st_folium

from dashboard_data import (
    build_master_table,
    ecosystem_kpis,
    get_excel_catalog_df,
    get_excel_catalog_items,
    get_parsed_data,
    hub_ranking,
    load_all_excel_sheets,
    load_excel_sheet,
    load_source_tables,
    master_dashboard_export,
    partner_dashboard_table,
    partner_summary,
)

# ── Brand & theme ──────────────────────────────────────────────────────────────
BRAND = {
    "navy": "#1E3A5F",
    "slate": "#334155",
    "muted": "#64748B",
    "border": "#E2E8F0",
    "surface": "#F8FAFC",
    "white": "#FFFFFF",
    "accent": "#2563EB",
    "success": "#059669",
}

PARTNER_COLORS = {
    "EBCI": "#1E3A5F",
    "Lumbee": "#2563EB",
    "NC": "#0891B2",
    "SDP": "#7C3AED",
    "Yurok": "#BE185D",
}

PARTNERS = ["All", "EBCI", "Lumbee", "NC", "SDP", "Yurok"]
TYPE_ORDER = [
    "Preschool/Head Start",
    "Elementary School",
    "Public/Community Library",
    "Tribal Library",
    "Community Program",
]

PLOTLY_TEMPLATE = go.layout.Template(
    layout=go.Layout(
        font=dict(family="Inter, system-ui, sans-serif", color=BRAND["slate"], size=13),
        paper_bgcolor=BRAND["white"],
        plot_bgcolor=BRAND["white"],
        title=dict(font=dict(size=16, color=BRAND["navy"], family="Inter, system-ui, sans-serif")),
        colorway=list(PARTNER_COLORS.values()),
        xaxis=dict(gridcolor=BRAND["border"], linecolor=BRAND["border"], zerolinecolor=BRAND["border"]),
        yaxis=dict(gridcolor=BRAND["border"], linecolor=BRAND["border"], zerolinecolor=BRAND["border"]),
        legend=dict(bgcolor="rgba(255,255,255,0.9)", bordercolor=BRAND["border"], borderwidth=1),
        margin=dict(l=48, r=32, t=64, b=48),
    )
)
pio.templates["rn_professional"] = PLOTLY_TEMPLATE
pio.templates.default = "rn_professional"

PROFESSIONAL_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', system-ui, -apple-system, sans-serif;
}

.block-container {
    padding-top: 1.5rem;
    padding-bottom: 3rem;
    max-width: 1400px;
}

header[data-testid="stHeader"] {
    background: #FFFFFF;
    border-bottom: 1px solid #E2E8F0;
}

section[data-testid="stSidebar"] {
    background: #F8FAFC;
    border-right: 1px solid #E2E8F0;
}

section[data-testid="stSidebar"] .block-container {
    padding-top: 2rem;
}

.hero {
    background: linear-gradient(135deg, #FFFFFF 0%, #F8FAFC 100%);
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 2rem 2.5rem;
    margin-bottom: 1.75rem;
}

.hero-eyebrow {
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #64748B;
    margin-bottom: 0.5rem;
}

.hero-title {
    font-size: 1.85rem;
    font-weight: 700;
    color: #0F172A;
    margin: 0 0 0.5rem 0;
    line-height: 1.2;
}

.hero-subtitle {
    font-size: 0.95rem;
    color: #64748B;
    margin: 0;
    max-width: 720px;
    line-height: 1.6;
}

.section-label {
    font-size: 0.7rem;
    font-weight: 600;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: #64748B;
    margin-bottom: 0.35rem;
}

.section-title {
    font-size: 1.25rem;
    font-weight: 600;
    color: #0F172A;
    margin: 0 0 0.35rem 0;
}

.section-desc {
    font-size: 0.875rem;
    color: #64748B;
    margin: 0 0 1.25rem 0;
    line-height: 1.5;
}

.kpi-card {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 10px;
    padding: 1.1rem 1.25rem;
    height: 100%;
    box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
}

.kpi-label {
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    color: #64748B;
    margin-bottom: 0.4rem;
}

.kpi-value {
    font-size: 1.65rem;
    font-weight: 700;
    color: #1E3A5F;
    line-height: 1.1;
}

.panel {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 10px;
    padding: 1.25rem 1.5rem;
    margin-bottom: 1rem;
}

.footer {
    margin-top: 2.5rem;
    padding-top: 1.25rem;
    border-top: 1px solid #E2E8F0;
    font-size: 0.78rem;
    color: #94A3B8;
    text-align: center;
}

div[data-testid="stMetric"] {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 10px;
    padding: 0.85rem 1rem;
    box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
}

div[data-testid="stMetricLabel"] {
    font-size: 0.72rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.05em;
    text-transform: uppercase;
    color: #64748B !important;
}

div[data-testid="stMetricValue"] {
    font-size: 1.5rem !important;
    font-weight: 700 !important;
    color: #1E3A5F !important;
}

.stTabs [data-baseweb="tab-list"] {
    gap: 0;
    background: #F8FAFC;
    border: 1px solid #E2E8F0;
    border-radius: 10px;
    padding: 4px;
}

.stTabs [data-baseweb="tab"] {
    border-radius: 8px;
    padding: 0.55rem 1.1rem;
    font-weight: 500;
    color: #64748B;
}

.stTabs [aria-selected="true"] {
    background: #FFFFFF !important;
    color: #1E3A5F !important;
    box-shadow: 0 1px 3px rgba(15, 23, 42, 0.08);
}

.stDownloadButton button {
    background: #1E3A5F !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    padding: 0.5rem 1.25rem !important;
}

.stDownloadButton button:hover {
    background: #152A45 !important;
}

hr {
    border: none;
    border-top: 1px solid #E2E8F0;
    margin: 1.5rem 0;
}
</style>
"""

st.set_page_config(
    page_title="Reading Nation | Book Ecosystem",
    page_icon="📖",
    layout="wide",
    initial_sidebar_state="expanded",
)


@st.cache_data
def get_master_table():
    return build_master_table(save=False)


def inject_styles():
    st.markdown(PROFESSIONAL_CSS, unsafe_allow_html=True)


def render_hero():
    st.markdown(
        """
        <div class="hero">
            <div class="hero-eyebrow">Reading Nation Waterfall</div>
            <h1 class="hero-title">Book Ecosystem Intelligence Dashboard</h1>
            <p class="hero-subtitle">
                Location-level analytics for book distribution, children served, library circulation,
                and geographic reach across tribal partner ecosystems.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section_header(title, description=""):
    desc_html = f'<p class="section-desc">{description}</p>' if description else ""
    st.markdown(
        f"""
        <div class="section-label">Section</div>
        <h2 class="section-title">{title}</h2>
        {desc_html}
        """,
        unsafe_allow_html=True,
    )


def kpi_row(kpis):
    cols = st.columns(6)
    labels = [
        ("Locations", kpis["locations"], False),
        ("Children Served", fmt_num(kpis["children_served"]), False),
        ("Books Disseminated", fmt_num(kpis["books_distributed"]), False),
        ("LFL Boxes", fmt_num(kpis["lfl_boxes"]) if kpis["lfl_boxes"] else "—", False),
        ("In Circulation", fmt_num(kpis["books_in_circulation"]), False),
        ("Total Activity", fmt_num(kpis["total_book_activity"]), False),
    ]
    for col, (label, value, _) in zip(cols, labels):
        with col:
            st.markdown(
                f"""
                <div class="kpi-card">
                    <div class="kpi-label">{label}</div>
                    <div class="kpi-value">{value}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


def apply_chart_style(fig, height=None):
    if height:
        fig.update_layout(height=height)
    fig.update_layout(
        template="rn_professional",
        hoverlabel=dict(bgcolor="white", font_size=12, font_family="Inter"),
    )
    return fig


def filter_master(master, partner, ecosystem_types):
    filtered = master.copy()
    if partner != "All":
        filtered = filtered[filtered["partner"] == partner]
    if ecosystem_types:
        filtered = filtered[filtered["ecosystem_type"].isin(ecosystem_types)]
    return filtered


def fmt_num(value):
    if pd.isna(value):
        return "—"
    if float(value).is_integer():
        return f"{int(value):,}"
    return f"{float(value):,.1f}"


def _num(value, default=0):
    if value is None or pd.isna(value):
        return default
    return float(value)


def flow_weight(row):
    return max(
        1.0,
        _num(row.get("books_distributed")) / 600
        + _num(row.get("books_in_circulation")) / 2000
        + _num(row.get("replenishment_books_read")) / 3000,
    )


def build_density_map(mapped, show_flow, hub_only=False):
    if hub_only:
        mapped = hub_ranking(mapped, n=min(8, len(mapped)))

    center = [mapped["latitude"].mean(), mapped["longitude"].mean()]
    zoom = 5 if len(mapped["partner"].unique()) == 1 else 4
    fmap = folium.Map(location=center, zoom_start=zoom, tiles="CartoDB positron")

    if show_flow:
        for _, group in mapped.groupby("partner"):
            ordered = group.copy()
            ordered["sort_key"] = ordered["ecosystem_type"].apply(
                lambda t: TYPE_ORDER.index(t) if t in TYPE_ORDER else 99
            )
            ordered = ordered.sort_values("sort_key")
            sites = ordered.to_dict("records")
            for i in range(len(sites) - 1):
                a, b = sites[i], sites[i + 1]
                folium.PolyLine(
                    locations=[[a["latitude"], a["longitude"]], [b["latitude"], b["longitude"]]],
                    color=PARTNER_COLORS.get(b["partner"], BRAND["navy"]),
                    weight=flow_weight(b),
                    opacity=0.55,
                    tooltip=(
                        f"{a['location_name']} → {b['location_name']} | "
                        f"{_num(b.get('books_distributed')):,.0f} books"
                    ),
                ).add_to(fmap)

    for _, row in mapped.iterrows():
        books = row["books_distributed"] if pd.notna(row["books_distributed"]) else 0
        color = PARTNER_COLORS.get(row["partner"], BRAND["navy"])
        folium.CircleMarker(
            location=[row["latitude"], row["longitude"]],
            radius=max(8, min(32, books / 400)),
            popup=(
                f"<b>{row['location_name']}</b><br>"
                f"{row['partner']} · {row['ecosystem_type']}<br>"
                f"Children: {fmt_num(row['children_served'])}<br>"
                f"Books: {fmt_num(row['books_distributed'])}"
            ),
            color=color,
            fill=True,
            fill_color=color,
            fill_opacity=0.8,
            weight=1,
        ).add_to(fmap)

    return fmap


def styled_dataframe(df, height=None):
    kwargs = {"use_container_width": True, "hide_index": True}
    if height:
        kwargs["height"] = height
    st.dataframe(
        df,
        **kwargs,
        column_config={
            col: st.column_config.NumberColumn(format="%,.0f")
            for col in df.columns
            if df[col].dtype in ["int64", "float64"]
            and any(k in col.lower() for k in ["book", "children", "circulation", "lfl", "activity", "served"])
        },
    )


def render_distribution_dashboard(filtered, summary, partner):
    section_header(
        "Book Distribution",
        "Comprehensive view of dissemination, enrollment, circulation, and ecosystem composition.",
    )
    kpi_row(ecosystem_kpis(filtered))

    st.markdown("<br>", unsafe_allow_html=True)
    dash_view = master_dashboard_export(filtered)

    col_dl, _ = st.columns([1, 3])
    with col_dl:
        st.download_button(
            "Export master table (CSV)",
            dash_view.to_csv(index=False).encode("utf-8"),
            file_name="location_master_dashboard.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with st.container():
        styled_dataframe(dash_view, height=360)

    st.markdown("---")
    section_header("Distribution analytics", "Comparative metrics across locations and partners.")

    r1c1, r1c2 = st.columns(2, gap="large")

    books_df = filtered.dropna(subset=["books_distributed"]).sort_values("books_distributed")
    if not books_df.empty:
        fig = apply_chart_style(
            px.bar(
                books_df,
                x="books_distributed",
                y="location_name",
                orientation="h",
                color="partner",
                color_discrete_map=PARTNER_COLORS,
                title="Books disseminated by location",
                labels={"books_distributed": "Books", "location_name": ""},
            ),
            max(400, len(books_df) * 28),
        )
        fig.update_layout(showlegend=True, yaxis_title="")
        r1c1.plotly_chart(fig, use_container_width=True)

    kids_df = filtered.dropna(subset=["children_served"]).sort_values("children_served")
    if not kids_df.empty:
        fig = apply_chart_style(
            px.bar(
                kids_df,
                x="children_served",
                y="location_name",
                orientation="h",
                color="ecosystem_type",
                title="Children served by location",
                labels={"children_served": "Children", "location_name": ""},
                color_discrete_sequence=px.colors.sequential.Blues_r,
            ),
            max(400, len(kids_df) * 28),
        )
        fig.update_layout(showlegend=False, yaxis_title="")
        r1c2.plotly_chart(fig, use_container_width=True)

    r2c1, r2c2 = st.columns(2, gap="large")

    lfl_df = filtered.copy()
    lfl_df["lfl_boxes"] = lfl_df["lfl_boxes"].fillna(0)
    if lfl_df["lfl_boxes"].sum() > 0:
        fig = apply_chart_style(
            px.bar(
                lfl_df.sort_values("lfl_boxes"),
                x="lfl_boxes",
                y="location_name",
                orientation="h",
                color="partner",
                color_discrete_map=PARTNER_COLORS,
                title="Little Free Library boxes",
            ),
            max(340, len(lfl_df) * 26),
        )
        r2c1.plotly_chart(fig, use_container_width=True)
    else:
        r2c1.markdown(
            '<div class="panel" style="color:#64748B;font-size:0.9rem;">'
            "LFL box counts pending — update <code>data/location_registry.csv</code>.</div>",
            unsafe_allow_html=True,
        )

    circ_df = filtered.dropna(subset=["books_in_circulation"]).sort_values("books_in_circulation")
    if not circ_df.empty:
        fig = apply_chart_style(
            px.bar(
                circ_df,
                x="books_in_circulation",
                y="location_name",
                orientation="h",
                color="ecosystem_type",
                title="Library circulation by location",
                labels={"books_in_circulation": "Circulation", "location_name": ""},
            ),
            max(340, len(circ_df) * 26),
        )
        r2c2.plotly_chart(fig, use_container_width=True)

    r3c1, r3c2, r3c3 = st.columns(3, gap="medium")
    pie_source = summary if partner == "All" else summary[summary["partner"] == partner]

    if not pie_source.empty:
        fig = apply_chart_style(
            px.pie(
                pie_source,
                names="partner",
                values="books_distributed",
                color="partner",
                color_discrete_map=PARTNER_COLORS,
                title="Partner share of books",
                hole=0.55,
            ),
            380,
        )
        fig.update_traces(textposition="outside", textinfo="percent+label")
        r3c1.plotly_chart(fig, use_container_width=True)

    type_summary = (
        filtered.groupby("ecosystem_type", as_index=False)
        .agg(books_distributed=("books_distributed", "sum"))
        .sort_values("books_distributed", ascending=False)
    )
    if not type_summary.empty:
        fig = apply_chart_style(
            px.bar(
                type_summary,
                x="ecosystem_type",
                y="books_distributed",
                title="Volume by ecosystem type",
                labels={"ecosystem_type": "", "books_distributed": "Books"},
                color_discrete_sequence=[BRAND["navy"]],
            ),
            380,
        )
        fig.update_layout(showlegend=False)
        r3c2.plotly_chart(fig, use_container_width=True)

    if not pie_source.empty:
        fig = apply_chart_style(
            go.Figure(
                data=[
                    go.Bar(name="Children served", x=pie_source["partner"], y=pie_source["children_served"], marker_color=BRAND["accent"]),
                    go.Bar(name="Books disseminated", x=pie_source["partner"], y=pie_source["books_distributed"], marker_color=BRAND["navy"]),
                ]
            ).update_layout(barmode="group", title="Scale & reach by partner", xaxis_title="", yaxis_title="Count"),
            380,
        )
        r3c3.plotly_chart(fig, use_container_width=True)

    treemap_df = filtered.dropna(subset=["books_distributed"])
    if not treemap_df.empty:
        fig = apply_chart_style(
            px.treemap(
                treemap_df,
                path=["partner", "ecosystem_type", "location_name"],
                values="books_distributed",
                color="partner",
                color_discrete_map=PARTNER_COLORS,
                title="Book volume hierarchy",
            ),
            420,
        )
        st.plotly_chart(fig, use_container_width=True)

    heatmap_df = filtered.pivot_table(
        index="partner", columns="ecosystem_type", values="books_distributed", aggfunc="sum", fill_value=0
    )
    if not heatmap_df.empty:
        fig = apply_chart_style(
            px.imshow(
                heatmap_df,
                text_auto=",.0f",
                aspect="auto",
                color_continuous_scale=[[0, "#F8FAFC"], [0.5, "#93C5FD"], [1, "#1E3A5F"]],
                title="Dissemination matrix: partner × ecosystem type",
            ),
            380,
        )
        st.plotly_chart(fig, use_container_width=True)


def render_density_map(filtered, show_flow_lines, hub_only):
    section_header(
        "Ecosystem Density Map",
        "Geographic distribution of book ecosystems. Marker size reflects dissemination volume; "
        "line weight reflects book flow between connected sites.",
    )

    mapped = filtered.dropna(subset=["latitude", "longitude"]).copy()
    if mapped.empty:
        st.warning("Add coordinates in `data/location_registry.csv` to enable mapping.")
        return

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Mapped sites", len(mapped))
    c2.metric("Partners", mapped["partner"].nunique())
    c3.metric("Children served", fmt_num(mapped["children_served"].sum(skipna=True)))
    c4.metric("Books disseminated", fmt_num(mapped["books_distributed"].sum(skipna=True)))

    with st.container():
        st_folium(build_density_map(mapped, show_flow_lines, hub_only), width=1200, height=520)

    geo_df = mapped.dropna(subset=["books_distributed"])
    if not geo_df.empty:
        fig = apply_chart_style(
            px.scatter_geo(
                geo_df,
                lat="latitude",
                lon="longitude",
                size="books_distributed",
                color="partner",
                color_discrete_map=PARTNER_COLORS,
                hover_name="location_name",
                hover_data={
                    "ecosystem_type": True,
                    "children_served": True,
                    "books_distributed": True,
                    "latitude": False,
                    "longitude": False,
                },
                scope="usa",
                title="National geographic reach",
            ),
            500,
        )
        fig.update_geos(showland=True, landcolor="#F8FAFC", countrycolor="#E2E8F0", bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)

    section_header("Highest-volume hubs", "Locations ranked by combined dissemination and activity.")
    hubs = hub_ranking(filtered, n=10)[
        [
            "location_name", "partner", "ecosystem_type", "children_served",
            "books_distributed", "books_in_circulation", "replenishment_books_read", "lfl_boxes",
        ]
    ].rename(columns={
        "location_name": "Location", "partner": "Partner", "ecosystem_type": "Type",
        "children_served": "Children", "books_distributed": "Books Disseminated",
        "books_in_circulation": "In Circulation", "replenishment_books_read": "Reading Logs",
        "lfl_boxes": "LFL Boxes",
    })
    styled_dataframe(hubs)

    fig = apply_chart_style(
        px.bar(
            hub_ranking(filtered, n=10).sort_values("books_distributed"),
            x="books_distributed",
            y="location_name",
            orientation="h",
            color="partner",
            color_discrete_map=PARTNER_COLORS,
            title="Top 10 hubs by dissemination volume",
            labels={"books_distributed": "Books", "location_name": ""},
        ),
        400,
    )
    st.plotly_chart(fig, use_container_width=True)


def render_all_excel_data(partner_filter):
    section_header(
        "All Excel Data",
        "Browse every sheet in Combined Statistics_All Partners.xlsx — raw workbook layout "
        "and structured CSV extracts side by side.",
    )

    catalog = get_excel_catalog_df()
    styled_dataframe(catalog, height=320)

    st.markdown("---")

    categories = sorted({item["category"] for item in get_excel_catalog_items()})
    col_cat, col_view = st.columns([2, 1])
    with col_cat:
        category = st.selectbox("Data category", categories)
    with col_view:
        view_mode = st.radio(
            "View",
            ["Raw Excel + Parsed CSV", "Raw Excel only", "Parsed CSV only"],
            horizontal=True,
            label_visibility="collapsed",
        )

    sheets_in_category = [item for item in get_excel_catalog_items() if item["category"] == category]
    sheet_labels = [f"{item['sheet']} ({item['partner']})" for item in sheets_in_category]
    sheet_pick = st.selectbox(
        "Excel sheet",
        range(len(sheets_in_category)),
        format_func=lambda i: sheet_labels[i],
    )
    selected = sheets_in_category[sheet_pick]

    parsed_partner = partner_filter
    if selected["partner"] != "All":
        parsed_partner = selected["partner"]

    raw_df = load_excel_sheet(selected["sheet"])
    parsed_df = get_parsed_data(selected["parsed_key"], parsed_partner)

    meta1, meta2, meta3 = st.columns(3)
    meta1.metric("Raw Excel rows", len(raw_df))
    meta2.metric("Raw Excel columns", len(raw_df.columns))
    meta3.metric("Parsed CSV rows", len(parsed_df))

    st.caption(selected["description"])

    dl_col1, dl_col2, _ = st.columns([1, 1, 3])
    with dl_col1:
        st.download_button(
            "Download raw sheet (CSV)",
            raw_df.to_csv(index=False).encode("utf-8"),
            file_name=f"{selected['sheet']}.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with dl_col2:
        st.download_button(
            "Download parsed data (CSV)",
            parsed_df.to_csv(index=False).encode("utf-8"),
            file_name=f"{selected['parsed_key']}_{parsed_partner}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    if view_mode == "Raw Excel only":
        styled_dataframe(raw_df, height=480)
    elif view_mode == "Parsed CSV only":
        styled_dataframe(parsed_df, height=480)
    else:
        left, right = st.columns(2, gap="large")
        with left:
            st.markdown("**Raw Excel sheet**")
            styled_dataframe(raw_df, height=480)
        with right:
            st.markdown(f"**Parsed dataset: `{selected['parsed_key']}`**")
            styled_dataframe(parsed_df, height=480)

    with st.expander("Browse all parsed datasets from Excel", expanded=False):
        all_parsed = load_source_tables()
        parsed_keys = [
            "reading_assessments",
            "reading_scores",
            "enrollment",
            "library_usage",
            "reading_logs",
            "dissemination",
            "book_ecosystem_combined",
            "location_master",
        ]
        pick = st.selectbox("Parsed dataset", parsed_keys, key="all_parsed_pick")
        df = all_parsed[pick]
        if pick == "reading_assessments" and partner_filter != "All":
            df = df[df["partner"] == partner_filter]
        elif pick == "location_master" and partner_filter != "All":
            df = df[df["partner"] == partner_filter]
        st.caption(f"{len(df)} rows")
        styled_dataframe(df, height=360)

    with st.expander("Raw Excel — all 10 sheets (quick jump)", expanded=False):
        all_sheets = load_all_excel_sheets()
        jump = st.selectbox("Sheet", list(all_sheets.keys()), key="jump_sheet")
        styled_dataframe(all_sheets[jump], height=300)


def render_footer():
    st.markdown(
        '<div class="footer">Reading Nation Waterfall · Book Ecosystem Dashboard · Data updated from partner statistics</div>',
        unsafe_allow_html=True,
    )


def main():
    inject_styles()
    master = get_master_table()
    summary = partner_summary(master)
    ecosystem_types = sorted(master["ecosystem_type"].dropna().unique())

    with st.sidebar:
        st.markdown("### Controls")
        st.markdown('<p style="color:#64748B;font-size:0.82rem;margin-top:-0.5rem;">Filter the dashboard view</p>', unsafe_allow_html=True)
        partner = st.selectbox("Partner organization", PARTNERS, label_visibility="collapsed")
        st.markdown("**Ecosystem type**")
        selected_types = st.multiselect("Types", ecosystem_types, default=ecosystem_types, label_visibility="collapsed")
        st.divider()
        st.markdown("**Map options**")
        show_flow_lines = st.toggle("Book-flow connection lines", value=True)
        hub_only_map = st.toggle("Show top hubs only", value=False)
        st.divider()
        if st.button("Refresh data", use_container_width=True, type="primary"):
            st.cache_data.clear()
            st.rerun()
        st.markdown(
            '<p style="color:#94A3B8;font-size:0.75rem;line-height:1.4;margin-top:1rem;">'
            "Edit <b>data/location_registry.csv</b> to update LFL boxes and coordinates.</p>",
            unsafe_allow_html=True,
        )

    render_hero()
    filtered = filter_master(master, partner, selected_types)

    tab_master, tab_distribution, tab_density, tab_partner, tab_excel, tab_data = st.tabs(
        [
            "Master Table",
            "Book Distribution",
            "Density Map",
            "Partner Summary",
            "All Excel Data",
            "Data Quality",
        ]
    )

    with tab_master:
        section_header(
            "Location Master Table",
            "Foundation dataset for all visualizations and GIS work. One row per site.",
        )
        dash_view = master_dashboard_export(filtered)
        col1, col2 = st.columns([1, 4])
        with col1:
            st.download_button(
                "Download CSV",
                dash_view.to_csv(index=False).encode("utf-8"),
                file_name="location_master_dashboard.csv",
                mime="text/csv",
                use_container_width=True,
            )
        styled_dataframe(dash_view, height=420)

        with st.expander("Field definitions", expanded=False):
            st.markdown(
                """
                | Field | Description |
                |-------|-------------|
                | **Location Name** | Official site name |
                | **Ecosystem Type** | Preschool, Elementary School, Public/Tribal Library, etc. |
                | **Total Children Served** | Latest enrollment |
                | **Total Books Distributed** | LFL dissemination total |
                | **Little Free Library Boxes** | Active LFL installations |
                | **Books In Circulation** | Library circulation count |
                | **Replenishment (Reading Logs)** | Reading program activity |
                | **Latitude / Longitude** | GIS coordinates |
                """
            )

    with tab_distribution:
        render_distribution_dashboard(filtered, summary, partner)

    with tab_density:
        render_density_map(filtered, show_flow_lines, hub_only_map)

    with tab_partner:
        section_header("Partner ecosystem breakdown", "Location-level tables aligned to the program reporting format.")
        dash_partner = partner if partner != "All" else st.selectbox("Select partner", [p for p in PARTNERS if p != "All"])
        styled_dataframe(partner_dashboard_table(master, dash_partner))
        if partner == "All":
            styled_dataframe(
                summary.rename(columns={
                    "partner": "Partner", "locations": "Locations",
                    "children_served": "Children Served", "books_distributed": "Books Disseminated",
                    "lfl_boxes": "LFL Boxes", "books_read_total": "Reading Log Books",
                    "circulation": "Circulation", "pct_of_ecosystem_books": "% of All Books",
                })
            )

    with tab_excel:
        render_all_excel_data(partner)

    with tab_data:
        section_header("Data quality", "Fields missing from the location master table.")
        gaps = filtered[filtered["data_gaps"].astype(str).str.len() > 0][
            ["location_name", "partner", "ecosystem_type", "data_gaps", "notes"]
        ]
        if gaps.empty:
            st.success("All fields complete for the current filter.")
        else:
            styled_dataframe(gaps)
        st.info(
            "Partner-level book fairs, giveaways, and registry-only fields are documented "
            "in the **All Excel Data** tab."
        )

    render_footer()


if __name__ == "__main__":
    main()
