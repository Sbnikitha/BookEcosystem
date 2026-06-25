import re
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).parent
DATASET = ROOT / "dataset"
REGISTRY = ROOT / "data" / "location_registry.csv"
EXCEL = ROOT / "Combined Statistics_All Partners.xlsx"
OUTPUT = DATASET / "location_master.csv"

EXCEL_CATALOG = [
    {
        "sheet": "New Lumbee Reading Data",
        "category": "Reading assessments",
        "partner": "Lumbee",
        "parsed_key": "reading_assessments",
        "description": "NC school report cards, EOG reading, enrollment metrics",
    },
    {
        "sheet": "New EBCI Reading Data",
        "category": "Reading assessments",
        "partner": "EBCI",
        "parsed_key": "reading_assessments",
        "description": "BIE report cards, ELA proficiency bands, enrollment",
    },
    {
        "sheet": "New SDP Reading Data",
        "category": "Reading assessments",
        "partner": "SDP",
        "parsed_key": "reading_assessments",
        "description": "NM district report cards, literacy and proficiency",
    },
    {
        "sheet": "New Yurok Reading Data",
        "category": "Reading assessments",
        "partner": "Yurok",
        "parsed_key": "reading_assessments",
        "description": "CAASPP ELA results and enrollment",
    },
    {
        "sheet": "New NC Reading Data",
        "category": "Reading assessments",
        "partner": "NC",
        "parsed_key": "reading_assessments",
        "description": "ESSA report card ELA proficiency and enrollment",
    },
    {
        "sheet": "Reading Scores",
        "category": "Program reading scores",
        "partner": "All",
        "parsed_key": "reading_scores",
        "description": "K-readiness, 4th grade NAEP-style, EOG results by partner",
    },
    {
        "sheet": "Library Usage Stats",
        "category": "Library usage",
        "partner": "All",
        "parsed_key": "library_usage",
        "description": "School and public/tribal library circulation, visits, programs",
    },
    {
        "sheet": "Enrollment #s",
        "category": "Enrollment",
        "partner": "All",
        "parsed_key": "enrollment",
        "description": "Head Start and school enrollment by partner location",
    },
    {
        "sheet": "Reading Log Stats",
        "category": "Reading logs",
        "partner": "All",
        "parsed_key": "reading_logs",
        "description": "Monthly books-read logs across Head Start, school, and public libraries",
    },
    {
        "sheet": "Dissemination #s",
        "category": "Book dissemination",
        "partner": "All",
        "parsed_key": "dissemination",
        "description": "LFL books distributed by site, book fairs, and giveaways",
    },
]

LOCATION_ALIASES = {
    "EBCI: Qualla Boundary Head Start": "ebci_qualla_hs",
    "EBCI: Cherokee Elementary School": "ebci_cherokee_elem",
    "EBCI: Snowbird Community Library": "ebci_snowbird_lib",
    "Lumbee: Pembroke Head Start": "lumbee_pembroke_hs",
    "Lumbee: Pembroke Elementary School": "lumbee_pembroke_elem",
    "Lumbee: Robeson Public Library": "lumbee_robeson_lib",
    "Lumbee: Pembroke Boys & Girls Club": "lumbee_robeson_lib",
    "NC: Lame Deer Head Start": "nc_lame_deer_hs",
    "NC: Lame Deer Elementary School": "nc_lame_deer_elem",
    "NC: Woodenlegs Library": "nc_woodenlegs_lib",
    "SDP: Santo Domingo Pueblo Head Start": "sdp_santo_hs",
    "SDP: Santo Domingo Elementary School": "sdp_santo_elem",
    "SDP: Santo Domingo Pueblo Public Library": "sdp_santo_pub_lib",
    "Yurok: Ke'pel Head Start": "yurok_kepel_hs",
    "Yurok: O Me-nok Learning Center": "yurok_omenok",
    "Yurok: 'O Me-nok  Learning Center": "yurok_omenok",
    "Yurok: Del Norte Public Library": "yurok_del_norte_lib",
    "Yurok: Del Norte County Public Library?": "yurok_del_norte_lib",
    "Yurok: Boys & Girls Club LFL": "yurok_bgc_lfl",
    "Yurok: 3 Yurok Tribe Head Start Sites": "yurok_kepel_hs",
}


def parse_numeric(value):
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    match = re.search(r"[\d.]+", text)
    return float(match.group()) if match else None


def load_registry():
    registry = pd.read_csv(REGISTRY)
    registry["lfl_boxes"] = pd.to_numeric(registry["lfl_boxes"], errors="coerce")
    return registry


def load_dissemination():
    df = pd.read_csv(DATASET / "dissemination.csv")
    df = df[df["location"] != "TOTAL"].copy()
    df["location_key"] = df["location"].map(LOCATION_ALIASES)
    return df


def load_enrollment_latest():
    df = pd.read_csv(DATASET / "enrollment.csv")
    df["location_key"] = df["location"].map(LOCATION_ALIASES)
    df = df.sort_values("period").groupby("location_key", as_index=False).tail(1)
    return df.rename(columns={"enrollment_count": "children_served", "period": "enrollment_period"})


def load_library_usage_latest():
    df = pd.read_csv(DATASET / "library_usage.csv")
    df["location_key"] = df["location"].map(LOCATION_ALIASES)
    df = df.sort_values("period").groupby("location_key", as_index=False).tail(1)
    return df.rename(
        columns={
            "period": "circulation_period",
            "circulation": "circulation_raw",
            "program_attendance": "program_attendance",
            "visitation": "visitation_raw",
        }
    )


def load_reading_log_totals():
    df = pd.read_csv(DATASET / "reading_logs.csv")
    df["location_key"] = df["location"].map(LOCATION_ALIASES)
    return df.groupby("location_key", as_index=False).agg(
        books_read_total=("books_read", lambda s: pd.to_numeric(s, errors="coerce").sum()),
        participants_met_minimum=("met_minimum", lambda s: pd.to_numeric(s, errors="coerce").sum()),
    )


def summarize_gaps(row):
    gaps = []
    if pd.isna(row.get("children_served")):
        gaps.append("children_served")
    if pd.isna(row.get("books_distributed")):
        gaps.append("books_distributed")
    if pd.isna(row.get("lfl_boxes")):
        gaps.append("lfl_boxes")
    if pd.isna(row.get("latitude")) or pd.isna(row.get("longitude")):
        gaps.append("coordinates")
    if pd.isna(row.get("circulation_numeric")) and pd.isna(row.get("books_read_total")):
        gaps.append("circulation_or_reading_log")
    return "; ".join(gaps)


def build_master_table(save=False):
    registry = load_registry()
    dissemination = load_dissemination()
    enrollment = load_enrollment_latest()
    library = load_library_usage_latest()
    reading_logs = load_reading_log_totals()

    books = dissemination.groupby("location_key", as_index=False).agg(
        books_distributed=("books_distributed", "sum"),
        dissemination_categories=("category", lambda s: "; ".join(sorted(set(s)))),
    )

    master = registry.merge(books, on="location_key", how="left")
    master = master.merge(
        enrollment[["location_key", "children_served", "enrollment_period"]],
        on="location_key",
        how="left",
    )
    master = master.merge(
        library[
            [
                "location_key",
                "circulation_period",
                "circulation_raw",
                "program_attendance",
                "visitation_raw",
            ]
        ],
        on="location_key",
        how="left",
    )
    master = master.merge(reading_logs, on="location_key", how="left")

    master["circulation_numeric"] = master["circulation_raw"].map(parse_numeric)
    master["visitation_numeric"] = master["visitation_raw"].map(parse_numeric)
    master["replenishment_books_read"] = master["books_read_total"]
    master["books_in_circulation"] = master["circulation_numeric"]
    master["total_book_activity"] = master[["books_in_circulation", "replenishment_books_read"]].sum(
        axis=1, min_count=1
    )

    ecosystem_total = master["books_distributed"].sum(skipna=True)
    master["pct_of_ecosystem_books"] = master["books_distributed"].apply(
        lambda x: round(100 * x / ecosystem_total, 1) if pd.notna(x) and ecosystem_total else None
    )

    column_order = [
        "location_key",
        "location_name",
        "partner",
        "ecosystem_type",
        "children_served",
        "enrollment_period",
        "books_distributed",
        "pct_of_ecosystem_books",
        "lfl_boxes",
        "books_read_total",
        "replenishment_books_read",
        "books_in_circulation",
        "total_book_activity",
        "participants_met_minimum",
        "circulation_numeric",
        "circulation_period",
        "circulation_raw",
        "program_attendance",
        "visitation_numeric",
        "visitation_raw",
        "latitude",
        "longitude",
        "dissemination_categories",
        "notes",
    ]
    master = master[column_order].sort_values(["partner", "ecosystem_type", "location_name"])

    if save:
        OUTPUT.parent.mkdir(parents=True, exist_ok=True)
        master.to_csv(OUTPUT, index=False)
        master_dashboard_export(master).to_csv(DATASET / "location_master_dashboard.csv", index=False)

    return master


def master_dashboard_export(master):
    """GIS- and dashboard-ready view with spec column names."""
    return master[
        [
            "location_key",
            "location_name",
            "partner",
            "ecosystem_type",
            "children_served",
            "books_distributed",
            "pct_of_ecosystem_books",
            "lfl_boxes",
            "books_in_circulation",
            "replenishment_books_read",
            "total_book_activity",
            "circulation_period",
            "circulation_raw",
            "visitation_numeric",
            "latitude",
            "longitude",
        ]
    ].rename(
        columns={
            "location_name": "Location Name",
            "partner": "Partner",
            "ecosystem_type": "Ecosystem Type",
            "children_served": "Total Children Served",
            "books_distributed": "Total Books Distributed",
            "pct_of_ecosystem_books": "% of Ecosystem Books",
            "lfl_boxes": "Little Free Library Boxes",
            "books_in_circulation": "Books In Circulation (Library)",
            "replenishment_books_read": "Replenishment (Reading Logs)",
            "total_book_activity": "Total Book Activity",
            "circulation_period": "Circulation Period",
            "circulation_raw": "Circulation Notes",
            "visitation_numeric": "Visitation (numeric)",
            "latitude": "Latitude",
            "longitude": "Longitude",
        }
    )


def ecosystem_kpis(master):
    return {
        "locations": len(master),
        "children_served": master["children_served"].sum(skipna=True),
        "books_distributed": master["books_distributed"].sum(skipna=True),
        "lfl_boxes": master["lfl_boxes"].sum(skipna=True),
        "books_in_circulation": master["books_in_circulation"].sum(skipna=True),
        "replenishment_books_read": master["replenishment_books_read"].sum(skipna=True),
        "total_book_activity": master["total_book_activity"].sum(skipna=True),
    }


def hub_ranking(master, n=10):
    ranked = master.copy()
    ranked["hub_score"] = ranked["books_distributed"].fillna(0) + ranked["total_book_activity"].fillna(0) * 0.25
    return ranked.sort_values("hub_score", ascending=False).head(n)


def partner_summary(master):
    summary = master.groupby("partner", as_index=False).agg(
        locations=("location_key", "count"),
        children_served=("children_served", "sum"),
        books_distributed=("books_distributed", "sum"),
        lfl_boxes=("lfl_boxes", "sum"),
        books_read_total=("books_read_total", "sum"),
        circulation=("circulation_numeric", "sum"),
    )
    total_books = summary["books_distributed"].sum()
    summary["pct_of_ecosystem_books"] = (100 * summary["books_distributed"] / total_books).round(1)
    return summary


def partner_dashboard_table(master, partner):
    subset = master[master["partner"] == partner].copy()
    display = subset[
        [
            "location_name",
            "ecosystem_type",
            "children_served",
            "books_distributed",
            "pct_of_ecosystem_books",
            "lfl_boxes",
        ]
    ].rename(
        columns={
            "location_name": "Location",
            "ecosystem_type": "Type",
            "children_served": "Children Served",
            "books_distributed": "Books Disseminated",
            "pct_of_ecosystem_books": "% of Ecosystem Books",
            "lfl_boxes": "Little Free Library Boxes",
        }
    )
    total = pd.DataFrame(
        [
            {
                "Location": "Total",
                "Type": "Book Ecosystem",
                "Children Served": subset["children_served"].sum(skipna=True),
                "Books Disseminated": subset["books_distributed"].sum(skipna=True),
                "% of Ecosystem Books": round(subset["pct_of_ecosystem_books"].sum(skipna=True), 1),
                "Little Free Library Boxes": subset["lfl_boxes"].sum(skipna=True),
            }
        ]
    )
    return pd.concat([display, total], ignore_index=True)


def load_source_tables():
    return {
        "dissemination": pd.read_csv(DATASET / "dissemination.csv"),
        "enrollment": pd.read_csv(DATASET / "enrollment.csv"),
        "library_usage": pd.read_csv(DATASET / "library_usage.csv"),
        "reading_logs": pd.read_csv(DATASET / "reading_logs.csv"),
        "reading_assessments": pd.read_csv(DATASET / "reading_assessments.csv"),
        "reading_scores": pd.read_csv(DATASET / "reading_scores.csv"),
        "location_registry": pd.read_csv(REGISTRY),
        "location_master": pd.read_csv(DATASET / "location_master.csv"),
        "book_ecosystem_combined": pd.read_csv(DATASET / "book_ecosystem_combined.csv"),
    }


def load_excel_sheet(sheet_name: str) -> pd.DataFrame:
    raw = pd.read_excel(EXCEL, sheet_name=sheet_name, header=None)
    raw.columns = [f"Col {i}" for i in range(len(raw.columns))]
    return raw


def load_all_excel_sheets() -> dict[str, pd.DataFrame]:
    xl = pd.ExcelFile(EXCEL)
    return {name: load_excel_sheet(name) for name in xl.sheet_names}


def get_excel_catalog_items():
    return EXCEL_CATALOG


def get_excel_catalog_df() -> pd.DataFrame:
    rows = []
    parsed = load_source_tables()
    for item in EXCEL_CATALOG:
        key = item["parsed_key"]
        parsed_df = parsed.get(key)
        partner_rows = len(parsed_df)
        if key == "reading_assessments" and item["partner"] != "All":
            partner_rows = len(parsed_df[parsed_df["partner"] == item["partner"]])
        rows.append(
            {
                "Excel sheet": item["sheet"],
                "Category": item["category"],
                "Partner": item["partner"],
                "Parsed dataset": key,
                "Parsed rows (relevant)": partner_rows,
                "Description": item["description"],
            }
        )
    return pd.DataFrame(rows)


def get_parsed_data(parsed_key: str, partner: str | None = None) -> pd.DataFrame:
    tables = load_source_tables()
    df = tables[parsed_key].copy()
    if parsed_key == "reading_assessments" and partner and partner != "All":
        df = df[df["partner"] == partner]
    elif parsed_key == "location_master" and partner and partner != "All":
        df = df[df["partner"] == partner]
    return df
