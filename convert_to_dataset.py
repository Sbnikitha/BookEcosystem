#!/usr/bin/env python3
"""Convert Combined Statistics Excel workbook into organized CSV datasets."""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

XLSX = Path("Combined Statistics_All Partners.xlsx")
OUTPUT_DIR = Path("dataset")

PARTNER_MAP = {
    "EBCI": "EBCI",
    "Lumbee": "Lumbee",
    "NC": "NC",
    "SDP": "SDP",
    "Yurok": "Yurok",
}

PARTNER_SHEETS = {
    "New Lumbee Reading Data": "Lumbee",
    "New EBCI Reading Data": "EBCI",
    "New SDP Reading Data": "SDP",
    "New Yurok Reading Data": "Yurok",
    "New NC Reading Data": "NC",
}


def clean_text(value) -> str | None:
    if pd.isna(value):
        return None
    text = str(value).strip()
    return text if text else None


def normalize_year(value) -> str | None:
    if pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return f"{value.year}-{value.month:02d}"
    text = str(value).strip()
    if not text or text.lower() in {"nan", "totals", "total"}:
        return None
    text = text.replace(" - ", "-").replace(" ", "")
    match = re.match(r"^Fall(\d{4})$", text, re.I)
    if match:
        return f"Fall {match.group(1)}"
    match = re.match(r"(\d{4})[-/](\d{2,4})", text)
    if match:
        start, end = match.groups()
        end = end if len(end) == 2 else end[-2:]
        return f"{start}-{end}"
    if re.match(r"^\d{4}-\d{2}$", text):
        return text
    if re.match(r"^\d{4}$", text):
        return text
    if text.lower().startswith("fall "):
        return text
    if text.lower().startswith("summer"):
        return text
    return text


def parse_label(raw: str) -> tuple[str | None, str | None, str | None]:
    if not raw:
        return None, None, None
    parts = [p.strip() for p in raw.split("\n") if p.strip()]
    if len(parts) == 1:
        return parts[0], parts[0], None
    if len(parts) == 2:
        return parts[0], parts[1], None
    return parts[0], "\n".join(parts[1:-1]), parts[-1]


def is_year_header(row: pd.Series) -> list[str | None]:
    years: list[str | None] = []
    for value in row:
        if pd.isna(value):
            years.append(None)
            continue
        text = str(value).strip()
        if re.match(r"^\d{4}-\d{2}$", text):
            years.append(text)
        elif re.match(r"^\d{4}-\d{2,4}$", text.replace(" ", "")):
            years.append(normalize_year(text))
        elif re.match(r"^\d{4}$", text):
            years.append(text)
        else:
            years.append(None)
    return years


def melt_year_table(
    df: pd.DataFrame,
    partner: str,
    source_sheet: str,
    section: str | None = None,
) -> list[dict]:
    rows: list[dict] = []
    current_section = section
    year_cols: list[str | None] = []

    for _, row in df.iterrows():
        label = clean_text(row.iloc[0])
        year_values = is_year_header(row)

        if any(year_values[1:]):
            year_cols = year_values
            if label and not any(y for y in year_cols if y):
                current_section = label
            continue

        if not label:
            continue

        if label.endswith("Report Cards") or "Report Card" in label:
            current_section = label
            continue

        if "Reading Assessment" in label and not year_cols:
            current_section = label
            continue

        entity, metric, student_group = parse_label(label)
        if not year_cols:
            continue

        notes = clean_text(row.iloc[-1]) if len(row) > 1 else None
        for col_idx, school_year in enumerate(year_cols):
            if not school_year or col_idx == 0:
                continue
            value = row.iloc[col_idx]
            if pd.isna(value):
                continue
            rows.append(
                {
                    "partner": partner,
                    "source_sheet": source_sheet,
                    "section": current_section,
                    "entity": entity,
                    "metric": metric,
                    "student_group": student_group,
                    "school_year": school_year,
                    "value": value,
                    "notes": notes if col_idx == 1 else None,
                }
            )
    return rows


def parse_partner_reading_sheets(xlsx: str) -> pd.DataFrame:
    records: list[dict] = []
    for sheet, partner in PARTNER_SHEETS.items():
        df = pd.read_excel(xlsx, sheet_name=sheet, header=None)
        records.extend(melt_year_table(df, partner, sheet))
    return pd.DataFrame(records)


def parse_reading_scores(xlsx: str) -> pd.DataFrame:
    df = pd.read_excel(xlsx, sheet_name="Reading Scores", header=None)
    records: list[dict] = []
    current_section: str | None = None
    metric_names: list[str | None] = []

    for _, row in df.iterrows():
        label = clean_text(row.iloc[0])

        if label and ("Scores" in label or "Test Results" in label):
            current_section = label
            metric_names = []
            continue

        if label == "Year":
            metric_names = []
            for col in range(1, len(row), 2):
                metric_names.append(clean_text(row.iloc[col + 1]) if col + 1 < len(row) else None)
            continue

        if not label or label.startswith("*"):
            continue

        school_year = normalize_year(label)
        if not school_year:
            continue

        for idx, partner in enumerate(["EBCI", "Lumbee", "NC", "SDP", "Yurok"]):
            val_col = 1 + idx * 2
            count_col = val_col + 1
            if val_col >= len(row):
                break
            value = row.iloc[val_col]
            count = row.iloc[count_col] if count_col < len(row) else None
            if pd.isna(value) and pd.isna(count):
                continue
            records.append(
                {
                    "section": current_section,
                    "school_year": school_year,
                    "partner": partner,
                    "metric": metric_names[idx] if idx < len(metric_names) else None,
                    "value": value,
                    "student_count": count,
                }
            )
    return pd.DataFrame(records)


def parse_multi_partner_blocks(
    df: pd.DataFrame,
    block_starts: list[tuple[int, str]],
    value_cols: list[str],
) -> pd.DataFrame:
    records: list[dict] = []

    for start_row, category in block_starts:
        header_row = df.iloc[start_row + 1]
        data_start = start_row + 3
        partner_specs: list[tuple[int, str | None, str]] = []

        col = 1
        while col < len(header_row):
            location = clean_text(header_row.iloc[col])
            if not location:
                col += len(value_cols)
                continue
            partner = next((p for key, p in PARTNER_MAP.items() if key in location), None)
            partner_specs.append((col, partner, location))
            col += len(value_cols)

        for row_idx in range(data_start, len(df)):
            row = df.iloc[row_idx]
            period = clean_text(row.iloc[0])
            if not period:
                continue
            if period.lower().startswith("notes"):
                break
            if period.lower() in {"year", "totals", "total"}:
                continue
            if "Enrollment Size" in period or "Usage Stats" in period:
                break

            for col_start, partner, location in partner_specs:
                values = {}
                for offset, field in enumerate(value_cols):
                    idx = col_start + offset
                    values[field] = row.iloc[idx] if idx < len(row) else None
                if all(pd.isna(v) or v == "" for v in values.values()):
                    continue
                records.append(
                    {
                        "category": category,
                        "partner": partner,
                        "location": location,
                        "period": normalize_year(period) or period,
                        **values,
                    }
                )
    return pd.DataFrame(records)


def parse_library_usage(xlsx: str) -> pd.DataFrame:
    df = pd.read_excel(xlsx, sheet_name="Library Usage Stats", header=None)
    return parse_multi_partner_blocks(
        df,
        [(0, "School Library"), (13, "Public/Tribal Library")],
        ["circulation", "program_attendance", "visitation"],
    )


def parse_enrollment(xlsx: str) -> pd.DataFrame:
    df = pd.read_excel(xlsx, sheet_name="Enrollment #s", header=None)
    return parse_multi_partner_blocks(
        df,
        [(0, "Head Start"), (12, "School")],
        ["enrollment_count", "notes"],
    )


def parse_reading_log_section(
    df: pd.DataFrame,
    category: str,
    location_row_idx: int,
    header_row_idx: int,
    data_start: int,
    data_end: int,
    block_width: int = 4,
) -> list[dict]:
    records: list[dict] = []
    location_row = df.iloc[location_row_idx]
    blocks: list[tuple[int, str | None, str | None]] = []

    for col in range(0, len(location_row), block_width):
        location = clean_text(location_row.iloc[col])
        if not location or location in {"Month/Year", "Year", "Mo/Yr"}:
            continue
        partner = next((p for key, p in PARTNER_MAP.items() if key in location), None)
        blocks.append((col, partner, location))

    for row_idx in range(data_start, data_end):
        row = df.iloc[row_idx]
        if all(pd.isna(row)):
            continue

        for col_start, partner, location in blocks:
            period_raw = clean_text(row.iloc[col_start]) if col_start < len(row) else None
            if not period_raw or period_raw.lower() in {"totals", "total"}:
                continue

            period = normalize_year(period_raw) or period_raw
            if not re.search(r"\d", period):
                continue

            books = row.iloc[col_start + 1] if col_start + 1 < len(row) else None
            minimum = row.iloc[col_start + 2] if col_start + 2 < len(row) else None
            notes = row.iloc[col_start + 3] if col_start + 3 < len(row) else None
            if pd.isna(books) and pd.isna(minimum):
                continue
            records.append(
                {
                    "category": category,
                    "partner": partner,
                    "location": location,
                    "period": period,
                    "books_read": books,
                    "met_minimum": minimum,
                    "notes": notes,
                }
            )
    return records


def parse_reading_logs(xlsx: str) -> pd.DataFrame:
    df = pd.read_excel(xlsx, sheet_name="Reading Log Stats", header=None)
    sections = [
        ("Head Start", 1, 2, 3, 10),
        ("School Library", 13, 14, 15, 32),
        ("Public/Tribal Library", 36, 37, 38, 44),
    ]
    records: list[dict] = []
    for category, loc_row, header_row, start, end in sections:
        records.extend(parse_reading_log_section(df, category, loc_row, header_row, start, end))
    return pd.DataFrame(records)


def parse_dissemination(xlsx: str) -> pd.DataFrame:
    df = pd.read_excel(xlsx, sheet_name="Dissemination #s", header=None)
    records: list[dict] = []
    current_category: str | None = None

    for _, row in df.iterrows():
        label = clean_text(row.iloc[0])
        count = row.iloc[1]
        if not label:
            continue
        if pd.isna(count) and label not in PARTNER_MAP and label != "TOTAL":
            current_category = label
            continue
        if label == "Location":
            continue
        if label == "TOTAL":
            records.append(
                {
                    "category": current_category,
                    "partner": "ALL",
                    "location": "TOTAL",
                    "books_distributed": count,
                }
            )
            continue

        partner = next((p for key, p in PARTNER_MAP.items() if key == label or key in label), None)
        records.append(
            {
                "category": current_category,
                "partner": partner,
                "location": label,
                "books_distributed": count,
            }
        )
    return pd.DataFrame(records)


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    xlsx = str(XLSX)

    datasets = {
        "reading_assessments": parse_partner_reading_sheets(xlsx),
        "reading_scores": parse_reading_scores(xlsx),
        "library_usage": parse_library_usage(xlsx),
        "enrollment": parse_enrollment(xlsx),
        "reading_logs": parse_reading_logs(xlsx),
        "dissemination": parse_dissemination(xlsx),
    }

    for name, frame in datasets.items():
        path = OUTPUT_DIR / f"{name}.csv"
        frame.to_csv(path, index=False)
        print(f"Wrote {path} ({len(frame)} rows)")

    combined = pd.concat(
        [
            datasets["reading_assessments"].assign(record_type="reading_assessment"),
            datasets["reading_scores"].assign(record_type="reading_score"),
            datasets["library_usage"].assign(record_type="library_usage"),
            datasets["enrollment"].assign(record_type="enrollment"),
            datasets["reading_logs"].assign(record_type="reading_log"),
            datasets["dissemination"].assign(record_type="dissemination"),
        ],
        ignore_index=True,
        sort=False,
    )
    combined_path = OUTPUT_DIR / "book_ecosystem_combined.csv"
    combined.to_csv(combined_path, index=False)
    print(f"Wrote {combined_path} ({len(combined)} rows)")

    print("\nSummary by partner (reading_assessments):")
    print(datasets["reading_assessments"].groupby("partner").size())


if __name__ == "__main__":
    main()
