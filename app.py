from __future__ import annotations

from pathlib import Path
import sys

import pandas as pd
import plotly.express as px
import streamlit as st
from streamlit.runtime.scriptrunner import get_script_run_ctx

from word_table_parser import (
    RawTable,
    combine_summary_tables,
    extract_word_tables,
    ordered_visits,
    raw_table_frame,
)


APP_DIR = Path(__file__).resolve().parent
EXAMPLE_DOCUMENT = APP_DIR / "Table_14.310.8_eGFR_Change_from_Baseline.docx"

# Make PyCharm's regular Run button behave like `streamlit run app.py`.
if __name__ == "__main__" and get_script_run_ctx(suppress_warning=True) is None:
    from streamlit.web import cli as streamlit_cli

    sys.argv = ["streamlit", "run", str(Path(__file__).resolve()), *sys.argv[1:]]
    raise SystemExit(streamlit_cli.main())

st.set_page_config(page_title="Longitudinal Word Table Explorer", page_icon="DT", layout="wide")
st.markdown(
    """
    <style>
    :root { --ink: #18302b; --accent: #d46035; --paper: #f5f0e6; }
    .stApp { background: linear-gradient(145deg, #f8f4ea 0%, #edf3ee 55%, #f7eee8 100%); }
    h1, h2, h3 { color: var(--ink); font-family: Georgia, 'Times New Roman', serif; }
    [data-testid="stMetric"] { background: rgba(255,255,255,.65); border: 1px solid #d8d2c5;
      border-radius: 4px; padding: .8rem; }
    div[data-testid="stPlotlyChart"] { background: rgba(255,255,255,.72); border: 1px solid #d8d2c5;
      padding: .5rem; }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def load_tables(content: bytes, filename: str) -> list[RawTable]:
    return extract_word_tables(content, filename)


@st.cache_data(show_spinner=False)
def tidy_tables(tables: list[RawTable]) -> pd.DataFrame:
    return combine_summary_tables(tables)


def build_plot(data: pd.DataFrame, plot_type: str, show_error: bool):
    visit_levels = ordered_visits(data)
    parameters = data["parameter"].drop_duplicates().tolist()
    facet_options = {
        "facet_col": "parameter" if len(parameters) > 1 else None,
        "facet_col_wrap": 2,
    }
    common = {
        "data_frame": data,
        "x": "visit",
        "y": "value",
        "color": "group",
        "category_orders": {"visit": visit_levels},
        "custom_data": ["n", "source_table", "metric"],
        "labels": {"visit": "Visit", "value": data["metric"].iloc[0], "group": "Group"},
        **facet_options,
    }

    error_column = "spread" if show_error and data["spread"].notna().any() else None
    if plot_type == "Line":
        figure = px.line(**common, markers=True, error_y=error_column)
    elif plot_type == "Bar":
        figure = px.bar(**common, barmode="group", error_y=error_column)
    else:
        figure = px.scatter(**common, error_y=error_column)

    figure.update_traces(
        hovertemplate=(
            "Visit: %{x}<br>Value: %{y:.2f}<br>n: %{customdata[0]}"
            "<br>Table: %{customdata[1]}<extra>%{fullData.name}</extra>"
        )
    )
    figure.update_layout(
        template="plotly_white",
        height=520 if len(parameters) <= 2 else 760,
        margin=dict(l=45, r=20, t=55, b=55),
        legend=dict(orientation="h", yanchor="bottom", y=1.04, xanchor="left", x=0),
        font=dict(family="Trebuchet MS", color="#18302b"),
        colorway=["#1f6b5c", "#d46035", "#486b9b", "#a4863d", "#724c3b"],
    )
    figure.update_xaxes(tickangle=-35)
    figure.update_yaxes(zeroline=True, zerolinecolor="#8d948f")
    figure.for_each_annotation(lambda item: item.update(text=item.text.split("=")[-1]))
    return figure


st.title("Longitudinal Word Table Explorer")
st.caption(
    "Combines report tables across pages, discovers subgroup parameters, and plots visits over time. "
    "Total columns are retained."
)

uploaded = st.file_uploader("Word report", type=["docx", "doc", "rtf"])
if uploaded is not None:
    document_content = uploaded.getvalue()
    document_name = uploaded.name
elif EXAMPLE_DOCUMENT.exists():
    document_content = EXAMPLE_DOCUMENT.read_bytes()
    document_name = EXAMPLE_DOCUMENT.name
    st.info(f"Using the workspace example: {document_name}")
else:
    st.stop()

try:
    with st.spinner("Reading every table in the report..."):
        all_tables = load_tables(document_content, document_name)
except Exception as exc:
    st.error(f"Could not read the document: {exc}")
    st.stop()

if not all_tables:
    st.warning("No Word tables were found in this document.")
    st.stop()

table_options = [table.index for table in all_tables]
selected_ids = st.sidebar.multiselect("Tables/pages", table_options, default=table_options)
selected_tables = [table for table in all_tables if table.index in selected_ids]
combined = tidy_tables(selected_tables)

if combined.empty:
    st.warning("The selected tables do not contain a recognizable visit-by-group summary layout.")
else:
    parsed_table_ids = set(combined["source_table"].astype(int))
    unparsed_table_ids = [table_id for table_id in selected_ids if table_id not in parsed_table_ids]
    if unparsed_table_ids:
        labels = ", ".join(f"Table {table_id}" for table_id in unparsed_table_ids)
        st.info(f"No trend layout was detected in: {labels}. These tables remain available below.")

    measures = combined["measure"].drop_duplicates().tolist()
    selected_measure = st.sidebar.selectbox("Measure", measures)
    measure_data = combined[combined["measure"] == selected_measure]

    metrics = measure_data["metric"].drop_duplicates().tolist()
    default_metric = metrics.index("Mean") if "Mean" in metrics else 0
    selected_metric = st.sidebar.selectbox("Statistic", metrics, index=default_metric)
    metric_data = measure_data[measure_data["metric"] == selected_metric]

    parameters = metric_data["parameter"].drop_duplicates().tolist()
    selected_parameters = st.sidebar.multiselect("Parameters", parameters, default=parameters)

    groups = metric_data["group"].drop_duplicates().tolist()
    selected_groups = st.sidebar.multiselect("Groups", groups, default=groups)

    plot_data = metric_data[
        metric_data["parameter"].isin(selected_parameters) & metric_data["group"].isin(selected_groups)
    ].copy()

    auto_plot = "Line" if plot_data["visit"].nunique() > 1 else "Bar"
    plot_choice = st.sidebar.selectbox("Figure", ["Auto", "Line", "Bar", "Scatter"])
    plot_type = auto_plot if plot_choice == "Auto" else plot_choice
    show_error = st.sidebar.checkbox("Show variability/error", value=True)

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Tables", len(selected_tables))
    col2.metric("Parameters", plot_data["parameter"].nunique())
    col3.metric("Visits", plot_data["visit"].nunique())
    col4.metric("Groups", plot_data["group"].nunique())

    if plot_data.empty:
        st.warning("Choose at least one parameter and group to draw the figure.")
    else:
        title_measure = plot_data["measure"].mode().iloc[0]
        st.subheader(title_measure)
        figure = build_plot(plot_data, plot_type, show_error)
        st.plotly_chart(figure, width="stretch")

        export = plot_data.sort_values(["parameter", "visit_order", "group"])
        st.download_button(
            "Download combined data (CSV)",
            export.to_csv(index=False).encode("utf-8"),
            file_name=f"{Path(document_name).stem}_longitudinal.csv",
            mime="text/csv",
        )
        st.download_button(
            "Download interactive figure (HTML)",
            figure.to_html(include_plotlyjs="cdn").encode("utf-8"),
            file_name=f"{Path(document_name).stem}_figure.html",
            mime="text/html",
        )

    with st.expander("Combined longitudinal data"):
        st.dataframe(combined, width="stretch", hide_index=True)

with st.expander("Original tables"):
    for table in all_tables:
        st.markdown(f"**Table {table.index}**")
        st.dataframe(raw_table_frame(table), width="stretch", hide_index=True)
