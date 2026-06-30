import io
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

# ------------------------------------------------------------
# Import HanParal core functions
# ------------------------------------------------------------

ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"

if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from hanparal_core import (
    parse_items_text,
    detect_target_columns,
    configure_target_columns,
    validate_base_columns,
    search_source_term,
    search_multiple_source_terms,
    collapse_repeated_source_segments,
    expand_search_results_to_token_hits,
    create_term_summary,
    prepare_annotation_table,
    export_annotation_workbook_to_bytes,
)


# ------------------------------------------------------------
# Page setup
# ------------------------------------------------------------

st.set_page_config(
    page_title="HanParal",
    page_icon="🔎",
    layout="wide",
)

st.title("HanParal")
st.caption("A no-code interface for multilingual aligned corpus concordance and annotation.")

st.markdown(
    """
    HanParal helps you search source-text terms in multilingual aligned corpora,
    generate token-level concordance results, and export an Excel annotation workbook.
    """
)

st.warning(
    "Do not upload copyrighted, private, or sensitive corpora to a public deployment "
    "unless you are allowed to process them there. For sensitive research data, run HanParal locally."
)


# ------------------------------------------------------------
# Sidebar
# ------------------------------------------------------------

st.sidebar.title("Workflow")

workflow = st.sidebar.radio(
    "Choose what you want to do",
    [
        "Create annotation file",
        "Analyze completed annotation file",
    ],
)

if workflow == "Analyze completed annotation file":
    st.info("Phase 2 analysis will be added after the Phase 1 interface is stable.")
    st.stop()


# ------------------------------------------------------------
# Helper: load uploaded file
# ------------------------------------------------------------

@st.cache_data
def load_uploaded_file(file_name, file_bytes):
    if file_name.lower().endswith(".csv"):
        return pd.read_csv(io.BytesIO(file_bytes))
    return pd.read_excel(io.BytesIO(file_bytes))


# ------------------------------------------------------------
# Step 1 — Upload corpus
# ------------------------------------------------------------

st.header("1. Upload your aligned corpus")

uploaded_file = st.file_uploader(
    "Upload an .xlsx or .csv file",
    type=["xlsx", "csv"],
)

if uploaded_file is None:
    st.info("Upload a file to begin.")
    st.stop()

try:
    corpus = load_uploaded_file(
        uploaded_file.name,
        uploaded_file.getvalue(),
    )
except Exception as error:
    st.error(f"HanParal could not read this file: {error}")
    st.stop()

corpus.columns = [str(col).strip() for col in corpus.columns]

try:
    target_columns = detect_target_columns(corpus)

    # Configure HanParal dynamically according to the uploaded corpus.
    configure_target_columns(
        n_targets=len(target_columns),
        n_strategy_slots=3,
    )

    validate_base_columns(corpus)

except Exception as error:
    st.error(str(error))
    st.markdown(
        """
        Your file should contain at least:

        `id | section | source_text | target_1`

        For multiple target texts, add:

        `target_2 | target_3 | target_4`
        """
    )
    st.stop()

st.success("Corpus loaded successfully.")

col1, col2, col3 = st.columns(3)

col1.metric("Segments", len(corpus))
col2.metric("Detected target texts", len(target_columns))
col3.metric("Columns", len(corpus.columns))

with st.expander("Preview corpus", expanded=False):
    st.dataframe(corpus.head(20), use_container_width=True)


# ------------------------------------------------------------
# Step 2 — Search settings
# ------------------------------------------------------------

st.header("2. Search source-text terms")

search_mode = st.radio(
    "Search mode",
    ["Single term", "Multiple terms"],
    horizontal=True,
)

if search_mode == "Single term":
    term = st.text_input(
        "Search term",
        placeholder="Example: 아리랑",
    )
    terms = [term.strip()] if term.strip() else []

else:
    terms_text = st.text_area(
        "Paste terms separated by commas or line breaks",
        placeholder="아리랑, 애국가, 태극기",
        height=120,
    )
    terms = parse_items_text(terms_text)


with st.expander("Advanced settings", expanded=False):
    case_sensitive = st.checkbox(
        "Treat uppercase/lowercase as different",
        value=False,
    )

    collapse_repeated = st.checkbox(
        "Merge repeated source segments before token expansion",
        value=True,
    )

    context_chars = st.slider(
        "Characters of context around each hit",
        min_value=10,
        max_value=150,
        value=35,
        step=5,
    )

    strategy_slots = st.slider(
        "Strategy slots per target text",
        min_value=1,
        max_value=5,
        value=3,
        step=1,
    )


# Reconfigure if the user changes the number of strategy slots.
configure_target_columns(
    n_targets=len(target_columns),
    n_strategy_slots=strategy_slots,
)


# ------------------------------------------------------------
# Step 3 — Annotation options
# ------------------------------------------------------------

st.header("3. Category and strategy options")

default_categories = "\n".join(
    [
        "Category 1",
        "Category 2",
        "Category 3",
        "Category 4",
        "Category 5",
    ]
)

category_options_text = st.text_area(
    "Category options",
    value=default_categories,
    height=130,
    help="Write one category per line.",
)

default_strategies = "\n".join(
    [
        "Repetition",
        "Orthographic adaptation",
        "Linguistic translation",
        "Extratextual gloss",
        "Intratextual gloss",
        "Synonymy",
        "Limited universalization",
        "Absolute universalization",
        "Naturalization",
        "Deletion",
        "Autonomous creation",
        "Compensation",
        "Displacement",
        "Attenuation",
    ]
)

strategy_options_text = st.text_area(
    "Strategy options",
    value=default_strategies,
    height=220,
    help="Write one strategy per line.",
)

category_options = parse_items_text(category_options_text)
strategy_options = parse_items_text(strategy_options_text)


# ------------------------------------------------------------
# Step 4 — Run search
# ------------------------------------------------------------

st.header("4. Run search")

run_search = st.button("Run search", type="primary")

if not run_search:
    st.stop()

if not terms:
    st.error("Please enter at least one search term.")
    st.stop()

with st.spinner("Searching corpus..."):
    if search_mode == "Single term":
        passage_results = search_source_term(
            corpus,
            terms[0],
            search_in=("source_text",),
            case_sensitive=case_sensitive,
        )
    else:
        passage_results = search_multiple_source_terms(
            corpus,
            terms,
            search_in=("source_text",),
            case_sensitive=case_sensitive,
        )

    raw_matching_rows = len(passage_results)

    if collapse_repeated:
        passage_results = collapse_repeated_source_segments(passage_results)

    search_results = expand_search_results_to_token_hits(
        passage_results,
        context_chars=context_chars,
        case_sensitive=case_sensitive,
    )

    term_summary = create_term_summary(search_results)

    annotation_table = prepare_annotation_table(search_results)


# ------------------------------------------------------------
# Step 5 — Results preview
# ------------------------------------------------------------

st.header("5. Results preview")

col1, col2, col3 = st.columns(3)

col1.metric("Raw matching rows", raw_matching_rows)
col2.metric("Token-hit rows", len(search_results))
col3.metric("Terms searched", len(terms))

if search_results.empty:
    st.warning("No hits found.")
    st.stop()

with st.expander("Term summary", expanded=True):
    st.dataframe(term_summary, use_container_width=True)

preview_columns = [
    col
    for col in (
        [
            "term_source",
            "category",
            "token_hit_id",
            "kwic_source",
            "source_text",
        ]
        + target_columns
    )
    if col in search_results.columns
]

st.dataframe(
    search_results[preview_columns].head(100),
    use_container_width=True,
)


# ------------------------------------------------------------
# Step 6 — Download annotation workbook
# ------------------------------------------------------------

st.header("6. Download annotation workbook")

with st.spinner("Preparing Excel workbook..."):
    excel_bytes = export_annotation_workbook_to_bytes(
        search_results=search_results,
        annotation_table=annotation_table,
        term_summary=term_summary,
        category_options=category_options,
        strategy_options=strategy_options,
        strategy_slots=strategy_slots,
    )

st.download_button(
    label="Download annotation workbook",
    data=excel_bytes,
    file_name="hanparal_annotation_workbook.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)