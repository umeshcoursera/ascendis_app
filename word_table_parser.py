from __future__ import annotations

import io
import re
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET

import pandas as pd


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{WORD_NS}}}"


@dataclass(frozen=True)
class RawTable:
    index: int
    rows: list[list[str]]


def normalize_space(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _cell_text(cell: ET.Element) -> str:
    paragraphs: list[str] = []
    for paragraph in cell.findall(f".//{W}p"):
        text = "".join(node.text or "" for node in paragraph.findall(f".//{W}t"))
        text = normalize_space(text)
        if text:
            paragraphs.append(text)
    return normalize_space(" ".join(paragraphs))


def extract_docx_tables(content: bytes) -> list[RawTable]:
    """Read every Word table directly from a DOCX, without Microsoft Word."""
    with zipfile.ZipFile(io.BytesIO(content)) as archive:
        try:
            document_xml = archive.read("word/document.xml")
        except KeyError as exc:
            raise ValueError("The file is not a valid Word DOCX document.") from exc

    root = ET.fromstring(document_xml)
    output: list[RawTable] = []

    for table_index, table in enumerate(root.findall(f".//{W}tbl"), start=1):
        rows: list[list[str]] = []
        vertical_values: dict[int, str] = {}

        for row_node in table.findall(f"./{W}tr"):
            row: list[str] = []
            column_index = 0

            for cell in row_node.findall(f"./{W}tc"):
                properties = cell.find(f"./{W}tcPr")
                span = 1
                vertical_merge = None

                if properties is not None:
                    span_node = properties.find(f"./{W}gridSpan")
                    if span_node is not None:
                        span = int(span_node.get(f"{W}val", "1"))
                    vertical_merge = properties.find(f"./{W}vMerge")

                value = _cell_text(cell)
                if vertical_merge is not None:
                    merge_type = vertical_merge.get(f"{W}val", "continue")
                    if merge_type == "continue" and not value:
                        value = vertical_values.get(column_index, "")

                for offset in range(span):
                    row.append(value)
                    if vertical_merge is not None:
                        vertical_values[column_index + offset] = value
                column_index += span

            rows.append(row)

        width = max((len(row) for row in rows), default=0)
        rows = [row + [""] * (width - len(row)) for row in rows]
        output.append(RawTable(table_index, rows))

    return output


def extract_legacy_word_tables(content: bytes, suffix: str) -> list[RawTable]:
    """Use Word automation for legacy DOC/RTF files on Windows."""
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise RuntimeError(
            "Legacy .doc and .rtf files require pywin32 and Microsoft Word. "
            "DOCX files do not have this requirement."
        ) from exc

    with tempfile.TemporaryDirectory(prefix="word-table-app-") as temp_dir:
        input_path = Path(temp_dir, f"upload{suffix}")
        input_path.write_bytes(content)

        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = None
        try:
            document = word.Documents.Open(str(input_path), ReadOnly=True)
            output: list[RawTable] = []
            for table_index in range(1, document.Tables.Count + 1):
                table = document.Tables(table_index)
                rows: list[list[str]] = []
                for row_index in range(1, table.Rows.Count + 1):
                    row = table.Rows(row_index)
                    values = [
                        normalize_space(cell.Range.Text.rstrip("\r\x07"))
                        for cell in row.Cells
                    ]
                    rows.append(values)
                width = max((len(row) for row in rows), default=0)
                rows = [row + [""] * (width - len(row)) for row in rows]
                output.append(RawTable(table_index, rows))
            return output
        finally:
            if document is not None:
                document.Close(False)
            word.Quit()


def extract_word_tables(content: bytes, filename: str) -> list[RawTable]:
    suffix = Path(filename).suffix.lower()
    if suffix == ".docx":
        return extract_docx_tables(content)
    if suffix in {".doc", ".rtf"}:
        return extract_legacy_word_tables(content, suffix)
    raise ValueError("Upload a .docx, .doc, or .rtf file.")


def raw_table_frame(table: RawTable) -> pd.DataFrame:
    width = max((len(row) for row in table.rows), default=0)
    rows = [row + [""] * (width - len(row)) for row in table.rows]
    return pd.DataFrame(rows, columns=[f"Column {i}" for i in range(1, width + 1)])


def _metric_name(header: str) -> str | None:
    header = normalize_space(header)
    lower = header.lower()
    if not lower or len(header) > 60 or "p-value" in lower or "p value" in lower:
        return None
    if "%" in lower or "percent" in lower or "proportion" in lower:
        return "Percent"
    if "geometric mean" in lower or re.search(r"\bgeo(?:metric)?\.? mean\b", lower):
        return "Geometric Mean"
    if "least squares mean" in lower or re.search(r"\bls mean\b", lower):
        return "LS Mean"
    if "adjusted mean" in lower:
        return "Adjusted Mean"
    if re.search(r"\bmean\b", lower):
        return "Mean"
    if re.search(r"\bmedian\b", lower):
        return "Median"
    if re.fullmatch(r"se|standard error(?: \([^)]*\))?", lower):
        return "SE"
    if re.fullmatch(r"sd|standard deviation(?: \([^)]*\))?", lower):
        return "SD"
    if re.fullmatch(r"(?:event |incidence )?rate(?: \([^)]*\))?", lower):
        return "Rate"
    return None


def _column_role(header: str) -> str:
    label = re.sub(r"[^a-z0-9]+", " ", normalize_space(header).lower()).strip()
    if re.fullmatch(
        r"visit|time|time point|timepoint|assessment|scheduled visit|study visit|week|day|month",
        label,
    ):
        return "visit"
    if re.fullmatch(r"statistics?|summary statistics?|descriptive statistics?|stats?", label):
        return "statistic"
    if re.search(r"\b(variable|parameter|outcome|endpoint|measure|test|analyte)\b", label):
        return "variable"
    return "descriptor"


def _row_layout_header(rows: list[list[str]]) -> tuple[int, int, int, list[int]] | None:
    """Find a Visit/Statistic table with one numeric column per treatment group."""
    for header_index in range(min(len(rows), 10)):
        roles = [_column_role(value) for value in rows[header_index]]
        if "visit" not in roles or "statistic" not in roles:
            continue

        visit_column = roles.index("visit")
        statistic_column = roles.index("statistic")
        group_columns: list[int] = []
        for column in range(len(rows[header_index])):
            if column in {visit_column, statistic_column} or not rows[header_index][column]:
                continue
            numeric_cells = sum(
                bool(_parse_numbers(row[column]))
                for row in rows[header_index + 1 :]
                if column < len(row) and normalize_space(row[column])
            )
            if numeric_cells and column > statistic_column:
                group_columns.append(column)

        if group_columns:
            return header_index, visit_column, statistic_column, group_columns
    return None


def _header_pair(rows: list[list[str]]) -> tuple[int, list[int]] | None:
    best: tuple[int, list[int]] | None = None
    for index in range(min(len(rows) - 1, 10)):
        metric_columns = [
            column
            for column, value in enumerate(rows[index + 1])
            if _metric_name(value) is not None
        ]
        if len(metric_columns) >= 1 and (best is None or len(metric_columns) > len(best[1])):
            best = (index, metric_columns)
    return best


def _strip_sample_size(label: str) -> str:
    label = re.sub(r"\s*\(\s*N\s*=.*?\)\s*$", "", label, flags=re.IGNORECASE)
    return normalize_space(label)


def _parse_numbers(value: object) -> list[float]:
    text = normalize_space(value).replace(",", "")
    return [float(item) for item in re.findall(r"(?<![A-Za-z])[-+]?\d*\.?\d+", text)]


def _parse_value(value: object, metric: str) -> tuple[float | None, float | None]:
    numbers = _parse_numbers(value)
    if not numbers:
        return None, None
    if metric == "Percent" and len(numbers) >= 2 and "(" in str(value):
        return numbers[1], None
    spread = numbers[1] if "Mean" in metric and len(numbers) >= 2 else None
    return numbers[0], spread


def _is_count_statistic(value: str) -> bool:
    label = re.sub(r"[^a-z0-9]+", "", normalize_space(value).lower())
    return label in {"n", "no", "number", "count"}


def _find_n_column(subheaders: list[str], group_headers: list[str], value_column: int) -> int | None:
    group = group_headers[value_column]
    candidates = range(max(1, value_column - 4), value_column)
    for column in reversed(list(candidates)):
        if group_headers[column] == group and re.fullmatch(
            r"n|no\.?|number", normalize_space(subheaders[column]), flags=re.IGNORECASE
        ):
            return column
    return None


def _section_parts(section: str, fallback: str) -> tuple[str, str]:
    section = normalize_space(section)
    if not section:
        return fallback, fallback
    for separator in (" - ", " -- ", ": "):
        if separator in section:
            measure, parameter = section.rsplit(separator, 1)
            if measure.strip() and parameter.strip():
                return normalize_space(measure), normalize_space(parameter)
    return section, fallback


def _visit_order(visit: str, fallback: int) -> float:
    lower = visit.lower()
    if "baseline" in lower or "screen" in lower:
        return 0.0
    number = re.search(r"[-+]?\d*\.?\d+", lower)
    if number:
        return float(number.group())
    if "end" in lower:
        return 1_000_000.0 + fallback
    return 500_000.0 + fallback


def _parse_column_statistic_table(table: RawTable) -> pd.DataFrame:
    """Parse tables where each group owns Mean/Median/etc. subcolumns."""
    rows = table.rows
    header = _header_pair(rows)
    if header is None:
        return pd.DataFrame()

    header_index, value_columns = header
    group_headers = rows[header_index]
    subheaders = rows[header_index + 1]
    records: list[dict[str, object]] = []
    section = ""

    for row_number, row in enumerate(rows[header_index + 2 :], start=header_index + 3):
        nonempty = list(dict.fromkeys(value for value in map(normalize_space, row) if value))
        numeric_values = {
            column: _parse_value(row[column], _metric_name(subheaders[column]) or "")
            for column in value_columns
        }
        has_value = any(value is not None for value, _ in numeric_values.values())

        if len(nonempty) == 1:
            section = nonempty[0]
            continue

        visit = normalize_space(row[0])
        if not visit or not has_value:
            continue

        measure, parameter = _section_parts(section, f"Table {table.index}")
        for column in value_columns:
            metric = _metric_name(subheaders[column])
            value, spread = numeric_values[column]
            if metric is None or value is None:
                continue

            group = _strip_sample_size(group_headers[column])
            if not group:
                group = f"Group {column}"
            n_column = _find_n_column(subheaders, group_headers, column)
            n_values = _parse_numbers(row[n_column]) if n_column is not None else []

            records.append(
                {
                    "source_table": table.index,
                    "measure": measure,
                    "parameter": parameter,
                    "visit": visit,
                    "visit_order": _visit_order(visit, row_number),
                    "group": group,
                    "metric": metric,
                    "value": value,
                    "spread": spread,
                    "n": int(n_values[0]) if n_values else pd.NA,
                }
            )

    return pd.DataFrame.from_records(records)


def _parse_row_statistic_table(table: RawTable) -> pd.DataFrame:
    """Parse tables with Visit, Variable, Statistic, and group value columns."""
    layout = _row_layout_header(table.rows)
    if layout is None:
        return pd.DataFrame()

    header_index, visit_column, statistic_column, group_columns = layout
    headers = table.rows[header_index]
    descriptor_columns = [
        column
        for column in range(len(headers))
        if column not in group_columns and column != statistic_column
    ]
    variable_columns = [
        column for column in descriptor_columns if _column_role(headers[column]) == "variable"
    ]
    context = {column: "" for column in descriptor_columns}
    sample_sizes: dict[tuple[str, str, str, str], int] = {}
    records: list[dict[str, object]] = []

    for row_number, row in enumerate(table.rows[header_index + 1 :], start=header_index + 2):
        for column in descriptor_columns:
            value = normalize_space(row[column])
            if value:
                context[column] = value

        visit = context.get(visit_column, "")
        statistic = normalize_space(row[statistic_column])
        if not visit or not statistic:
            continue

        measure_values = [context[column] for column in variable_columns if context.get(column)]
        measure = " | ".join(dict.fromkeys(measure_values)) or f"Table {table.index}"

        parameter_parts: list[str] = []
        for column in descriptor_columns:
            if column == visit_column or column in variable_columns or not context.get(column):
                continue
            parameter_parts.append(f"{normalize_space(headers[column])}: {context[column]}")
        parameter = " | ".join(parameter_parts) or measure

        for column in group_columns:
            group = _strip_sample_size(headers[column]) or f"Group {column}"
            key = (visit, measure, parameter, group)
            numbers = _parse_numbers(row[column])
            if _is_count_statistic(statistic):
                if numbers:
                    sample_sizes[key] = int(numbers[0])
                continue

            metric = _metric_name(statistic)
            value, spread = _parse_value(row[column], metric or "")
            if metric is None or value is None:
                continue

            records.append(
                {
                    "source_table": table.index,
                    "measure": measure,
                    "parameter": parameter,
                    "visit": visit,
                    "visit_order": _visit_order(visit, row_number),
                    "group": group,
                    "metric": metric,
                    "value": value,
                    "spread": spread,
                    "n": sample_sizes.get(key, pd.NA),
                }
            )

    return pd.DataFrame.from_records(records)


def parse_summary_table(table: RawTable) -> pd.DataFrame:
    """Convert a supported trend-table layout to tidy plotting records."""
    row_layout = _parse_row_statistic_table(table)
    if not row_layout.empty:
        return row_layout
    return _parse_column_statistic_table(table)


def combine_summary_tables(tables: Iterable[RawTable]) -> pd.DataFrame:
    """Stack continuation pages and retain distinct parameters for faceting."""
    frames = [parse_summary_table(table) for table in tables]
    frames = [frame for frame in frames if not frame.empty]
    if not frames:
        return pd.DataFrame()

    combined = pd.concat(frames, ignore_index=True)
    key = ["measure", "parameter", "visit", "group", "metric"]
    combined = combined.drop_duplicates(key, keep="first")

    parameter_order = {value: index for index, value in enumerate(combined["parameter"].unique())}
    group_order = {value: index for index, value in enumerate(combined["group"].unique())}
    combined = combined.assign(
        _parameter_order=combined["parameter"].map(parameter_order),
        _group_order=combined["group"].map(group_order),
    )
    combined = combined.sort_values(
        ["_parameter_order", "_group_order", "visit_order", "source_table"], kind="stable"
    )
    combined = combined.drop(columns=["_parameter_order", "_group_order"]).reset_index(drop=True)
    return combined


def ordered_visits(data: pd.DataFrame) -> list[str]:
    visits = data[["visit", "visit_order"]].drop_duplicates()
    visits = visits.sort_values(["visit_order", "visit"])
    return visits["visit"].tolist()
