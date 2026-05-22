"""
Excel Automator — Streamlit app
Three tools in one:
  1. Sheet / File Merger
  2. Auto-Formatter & Exception Highlighter
  3. Budget Variance Reporter
"""

import io
import os
import streamlit as st
import pandas as pd

from tools.merger    import merge_files, merge_sheets
from tools.formatter import format_and_highlight
from tools.variance  import generate_variance_report


# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Excel Automator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Shared styles ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .tool-header { font-size: 1.6rem; font-weight: 700; margin-bottom: 0.2rem; }
    .tool-sub    { color: #888; margin-bottom: 1.2rem; font-size: 0.95rem; }
    .legend-box  { background:var(--secondary-background-color); color:var(--text-color);
                   border:1px solid rgba(128,128,128,0.2); border-radius:8px;
                   padding:12px 16px; margin-bottom:1rem; }
    .legend-item { display:inline-block; margin-right:1.2rem; font-size:0.85rem; color:var(--text-color); }
    .dot-red    { display:inline-block;width:14px;height:14px;border-radius:50%;background:#FF4C4C;margin-right:5px;vertical-align:middle;flex-shrink:0; }
    .dot-orange { display:inline-block;width:14px;height:14px;border-radius:50%;background:#FF9900;margin-right:5px;vertical-align:middle;flex-shrink:0; }
    .dot-yellow { display:inline-block;width:14px;height:14px;border-radius:50%;background:#F4C430;margin-right:5px;vertical-align:middle;flex-shrink:0; }
    .dot-green  { display:inline-block;width:14px;height:14px;border-radius:50%;background:#3CB371;margin-right:5px;vertical-align:middle;flex-shrink:0; }
</style>
""", unsafe_allow_html=True)


# ── Sidebar navigation ─────────────────────────────────────────────────────────
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/microsoft-excel-2019.png", width=56)
    st.title("Excel Automator")
    st.caption("Upload → Process → Download")
    st.divider()
    tool = st.radio(
        "Choose a tool",
        ["📂 Sheet / File Merger", "🎨 Auto-Formatter", "📊 Variance Reporter"],
        label_visibility="collapsed",
    )
    st.divider()
    st.caption("🔒 Files are processed in memory and never stored.")
    st.caption("Built with Python · pandas · openpyxl")


# ── File-like wrapper so stored bytes work with existing tool functions ────────
class _NamedBytes:
    def __init__(self, name: str, data: bytes):
        self.name = name
        self._data = data
    def read(self):
        return self._data


# ── on_change callbacks: read bytes the moment files land in the uploader ──────
def _save_multi():
    files = st.session_state.get("uploader_multi", [])
    if files:
        st.session_state["stored_multi"] = [
            {"name": f.name, "data": f.read()} for f in files
        ]
    else:
        st.session_state["stored_multi"] = []

def _save_single():
    f = st.session_state.get("uploader_single")
    if f:
        st.session_state["stored_single"] = {"name": f.name, "data": f.read()}
    else:
        st.session_state["stored_single"] = None

def _save_formatter():
    f = st.session_state.get("uploader_formatter")
    if f:
        st.session_state["stored_formatter"] = {"name": f.name, "data": f.read()}
    else:
        st.session_state["stored_formatter"] = None

def _save_variance():
    f = st.session_state.get("uploader_variance")
    if f:
        st.session_state["stored_variance"] = {"name": f.name, "data": f.read()}
    else:
        st.session_state["stored_variance"] = None


# ── Shared result display ──────────────────────────────────────────────────────
def show_result(file_bytes: bytes, filename: str, success_msg: str):
    st.success(f"✅ {success_msg}")
    st.download_button(
        label=f"⬇ Download {filename}",
        data=file_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    st.divider()
    st.caption("Preview (first 50 rows):")
    try:
        preview_df = pd.read_excel(io.BytesIO(file_bytes), nrows=50)
        st.dataframe(preview_df, use_container_width=True)
    except Exception:
        st.info("Preview unavailable for multi-sheet files — open the downloaded file to see all sheets.")


# ══════════════════════════════════════════════════════════════════════════════
# TOOL 1 — Sheet / File Merger
# ══════════════════════════════════════════════════════════════════════════════
if tool == "📂 Sheet / File Merger":
    st.markdown('<div class="tool-header">📂 Sheet / File Merger</div>', unsafe_allow_html=True)
    st.markdown('<div class="tool-sub">Combine multiple Excel files or sheets into one clean master file.</div>', unsafe_allow_html=True)

    mode = st.radio(
        "What do you want to merge?",
        ["Multiple files → one master", "Multiple sheets inside one file → one sheet"],
        horizontal=True,
    )
    st.divider()

    # ── Multi-file mode ────────────────────────────────────────────────────────
    if mode == "Multiple files → one master":

        with st.expander("💡 Need sample files to try?"):
            c1, c2, c3 = st.columns(3)
            with c1:
                with open("sample_data/SAMPLE~1.xlsx", "rb") as f:
                    st.download_button("⬇ SAMPLE~1.xlsx", f, "SAMPLE~1.xlsx", key="dl_a")
            with c2:
                with open("sample_data/SAMPLE~2.xlsx", "rb") as f:
                    st.download_button("⬇ SAMPLE~2.xlsx", f, "SAMPLE~2.xlsx", key="dl_b")
            with c3:
                with open("sample_data/SAMPLE~3.xlsx", "rb") as f:
                    st.download_button("⬇ SAMPLE~3.xlsx", f, "SAMPLE~3.xlsx", key="dl_c")

        st.file_uploader(
            "Upload two or more Excel files",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            key="uploader_multi",
            on_change=_save_multi,
            help="Each file's data will be combined. A _source_file column tracks origin.",
        )

        stored = st.session_state.get("stored_multi", [])
        if stored:
            names = ', '.join(e['name'].replace('~', '\~') for e in stored)
            st.success(f"📎 {len(stored)} file(s) ready: {names}")

        merge_mode = st.radio(
            "Merge style",
            ["Stack rows vertically (one big sheet)", "Keep each file as its own sheet"],
            horizontal=True,
        )


        if st.button("Merge files", type="primary"):
            if not stored:
                st.warning("Please upload at least two Excel files first.")
            else:
                with st.spinner("Merging..."):
                    try:
                        file_objs = [_NamedBytes(e["name"], e["data"]) for e in stored]
                        style = "stack" if "Stack" in merge_mode else "sheets"
                        result_bytes = merge_files(file_objs, mode=style)
                        show_result(result_bytes, "merged_master.xlsx", "Merged file ready!")
                    except Exception as e:
                        st.error(f"Error: {e}")

    # ── Single-file / multi-sheet mode ────────────────────────────────────────
    else:
        st.file_uploader(
            "Upload one Excel file with multiple sheets",
            type=["xlsx", "xls"],
            accept_multiple_files=False,
            key="uploader_single",
            on_change=_save_single,
        )

        stored = st.session_state.get("stored_single")
        if stored:
            st.success(f"📎 Ready: {stored['name']}")

        if st.button("Merge sheets", type="primary"):
            if not stored:
                st.warning("Please upload a file first.")
            else:
                with st.spinner("Merging sheets..."):
                    try:
                        result_bytes = merge_sheets(_NamedBytes(stored["name"], stored["data"]))
                        show_result(result_bytes, "merged_sheets.xlsx", "Sheets merged!")
                    except Exception as e:
                        st.error(f"Error: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# TOOL 2 — Auto-Formatter
# ══════════════════════════════════════════════════════════════════════════════
elif tool == "🎨 Auto-Formatter":
    st.markdown('<div class="tool-header">🎨 Auto-Formatter & Exception Highlighter</div>', unsafe_allow_html=True)
    st.markdown('<div class="tool-sub">Instantly highlight overdue dates, budget overruns, and negative values.</div>', unsafe_allow_html=True)

    st.markdown("""
    <div class="legend-box">
        <span class="legend-item"><span class="dot-red"></span>Overdue date</span>
        <span class="legend-item"><span class="dot-orange"></span>Budget overrun</span>
        <span class="legend-item"><span class="dot-yellow"></span>Negative value</span>
    </div>
    """, unsafe_allow_html=True)

    with st.expander("💡 Need a sample file to try?"):
        with open("sample_data/sample_format.xlsx", "rb") as f:
            st.download_button("⬇ sample_format.xlsx", f, "sample_format.xlsx", key="dl_fmt")

    st.file_uploader(
        "Upload an Excel file",
        type=["xlsx", "xls"],
        key="uploader_formatter",
        on_change=_save_formatter,
    )

    stored = st.session_state.get("stored_formatter")
    if stored:
        st.success(f"📎 Ready: {stored['name']}")

    with st.expander("⚙️ Highlight settings", expanded=True):
        col1, col2, col3 = st.columns(3)
        flag_overdue   = col1.checkbox("Flag overdue dates",   value=True)
        flag_negatives = col2.checkbox("Flag negative values", value=True)
        enable_overrun = col3.checkbox("Flag budget overruns", value=False)
        overrun_threshold = 0.0
        if enable_overrun:
            overrun_threshold = st.number_input(
                "Flag budget/cost/spend columns above this value ($)",
                min_value=0.0, value=10000.0, step=500.0,
            )

    if st.button("Format & highlight", type="primary"):
        if not stored:
            st.warning("Please upload a file first.")
        else:
            with st.spinner("Formatting..."):
                try:
                    result_bytes = format_and_highlight(
                        _NamedBytes(stored["name"], stored["data"]),
                        overrun_threshold=overrun_threshold,
                        flag_overdue=flag_overdue,
                        flag_negatives=flag_negatives,
                    )
                    show_result(result_bytes, "formatted_output.xlsx", "Formatting complete!")
                except Exception as e:
                    st.error(f"Error: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# TOOL 3 — Variance Reporter
# ══════════════════════════════════════════════════════════════════════════════
elif tool == "📊 Variance Reporter":
    st.markdown('<div class="tool-header">📊 Budget Variance Reporter</div>', unsafe_allow_html=True)
    st.markdown('<div class="tool-sub">Compare actuals vs. forecast, flag overruns, generate a summary + chart.</div>', unsafe_allow_html=True)

    st.markdown("""
    <div class="legend-box">
        <span class="legend-item"><span class="dot-green"></span>Under budget</span>
        <span class="legend-item"><span class="dot-yellow"></span>0 – threshold %</span>
        <span class="legend-item"><span class="dot-red"></span>Over threshold (flagged)</span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("**Your file needs at least:** a *Category* column · a *Budget* or *Forecast* column · an *Actual* column")

    with st.expander("💡 Need a sample file to try?"):
        with open("sample_data/sample_variance.xlsx", "rb") as f:
            st.download_button("⬇ sample_variance.xlsx", f, "sample_variance.xlsx", key="dl_var")

    st.file_uploader(
        "Upload your budget file",
        type=["xlsx", "xls"],
        key="uploader_variance",
        on_change=_save_variance,
    )

    stored = st.session_state.get("stored_variance")
    if stored:
        st.success(f"📎 Ready: {stored['name']}")

    threshold = st.slider(
        "Flag variances above (%)",
        min_value=1, max_value=50, value=10,
        help="Rows where Actual exceeds Budget by more than this % are highlighted red.",
    )

    if st.button("Generate report", type="primary"):
        if not stored:
            st.warning("Please upload a file first.")
        else:
            with st.spinner("Generating variance report..."):
                try:
                    result_bytes = generate_variance_report(
                        _NamedBytes(stored["name"], stored["data"]),
                        variance_threshold=float(threshold),
                    )
                    show_result(result_bytes, "variance_report.xlsx", "Report ready! Two sheets: Variance Report + Summary with chart.")
                except ValueError as e:
                    st.error(str(e))
                except Exception as e:
                    st.error(f"Unexpected error: {e}")
