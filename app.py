"""Streamlit presentation skin for the deterministic RCM-EC engine."""

from __future__ import annotations

from html import escape
from pathlib import Path

import pandas as pd
import streamlit as st

from engine.io_excel import read_config, register_to_bytes
from engine.library import run_library
from engine.reconcile import ReconciliationResult


APP_ROOT = Path(__file__).resolve().parent
LIBRARY_ROOT = APP_ROOT / "synthetic_library"
CONFIG_PATH = LIBRARY_ROOT / "config" / "config_field_map.xlsx"


st.set_page_config(
    page_title="RCM–EC Reconciliation Demo",
    page_icon="🔎",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.4rem; padding-bottom: 2rem; max-width: 1500px;}
    .provenance-banner {
        border-radius: 8px; color: #FFFFFF; font-size: 0.95rem; font-weight: 800;
        letter-spacing: 0.05em; margin-bottom: 1rem; padding: 0.8rem 1rem;
        text-align: center;
    }
    .synthetic-banner {background: #FF4F59;}
    .uploaded-banner {background: #B56A00;}
    div[data-testid="stMetric"] {
        background: #282A27; border: 1px solid #3A3D39; border-radius: 8px;
        min-height: 118px; padding: 0.8rem 1rem;
    }
    div[data-testid="stMetric"] label {color: #D2D5D9;}
    .email-card {
        background: #FFFFFF; border-radius: 8px; color: #181C23;
        margin-top: 0.75rem; padding: 1.25rem 1.5rem;
    }
    .email-subject {border-bottom: 1px solid #D8D8D8; font-weight: 800; margin-bottom: 1rem; padding-bottom: 0.8rem;}
    .email-card table {border-collapse: collapse; font-size: 0.88rem; width: 100%;}
    .email-card th {background: #F1F2F4; text-align: left;}
    .email-card th, .email-card td {border-bottom: 1px solid #D8D8D8; padding: 0.55rem; vertical-align: top;}
    .resolved-heading {color: #43C98A; font-weight: 800; margin-top: 1.2rem;}
    .footer {color: #A8ADB5; font-size: 0.78rem; margin-top: 2.5rem; text-align: center;}
    </style>
    """,
    unsafe_allow_html=True,
)


def render_banner(mode: str) -> None:
    if mode == "Library demo":
        css_class = "synthetic-banner"
        text = "SYNTHETIC DATA — DEMO ONLY"
    else:
        css_class = "uploaded-banner"
        text = "UPLOADED DATA — data-handling gate applies"
    st.markdown(
        f'<div class="provenance-banner {css_class}">{escape(text)}</div>',
        unsafe_allow_html=True,
    )


def render_footer() -> None:
    st.markdown(
        '<div class="footer">Deterministic engine — no AI components. Logic mirrors Office Script '
        'RunReconciliation (M365-only build, verified 27 Aug 2026).</div>',
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def load_library_results(apply_normalisers: bool) -> dict[str, ReconciliationResult]:
    return run_library(LIBRARY_ROOT, apply_normalisers=apply_normalisers)


@st.cache_data(show_spinner=False)
def load_runtime_config():
    return read_config(CONFIG_PATH)


def findings_frame(result: ReconciliationResult) -> pd.DataFrame:
    return pd.DataFrame(result.findings)


def filter_findings(
    frame: pd.DataFrame,
    severities: list[str],
    statuses: list[str],
    recruiters: list[str],
    search_text: str,
) -> pd.DataFrame:
    if frame.empty:
        return frame
    filtered = frame[
        frame["Severity"].isin(severities)
        & frame["Status"].isin(statuses)
        & frame["Recruiter"].isin(recruiters)
    ]
    search = search_text.strip()
    if search:
        matches = filtered.astype(str).apply(
            lambda row: row.str.contains(search, case=False, regex=False).any(),
            axis=1,
        )
        filtered = filtered[matches]
    return filtered


def style_register(frame: pd.DataFrame):
    def severity_style(value: str) -> str:
        return "background-color:#FF4F59;color:#FFFFFF;font-weight:700" if value == "HIGH" else ""

    def status_style(value: str) -> str:
        if value == "New":
            return "background-color:#FFAD28;color:#181C23;font-weight:700"
        if value == "Recurring":
            return "background-color:#6558D3;color:#FFFFFF;font-weight:700"
        return ""

    return frame.style.map(severity_style, subset=["Severity"]).map(status_style, subset=["Status"])


if "app_mode" not in st.session_state:
    st.session_state.app_mode = "Library demo"
if "normalisers_on" not in st.session_state:
    st.session_state.normalisers_on = True

render_banner(st.session_state.app_mode)
st.title("RCM–EC requisition ↔ position reconciliation")
st.caption("Presentation skin for the verified M365-only reconciliation build.")

st.radio(
    "Mode",
    options=("Library demo", "Real-time Excel uploads"),
    horizontal=True,
    key="app_mode",
)

if st.session_state.app_mode == "Real-time Excel uploads":
    st.info("Upload mode is introduced in Phase 4. Library mode is complete and available above.")
    render_footer()
    st.stop()

try:
    results = load_library_results(st.session_state.normalisers_on)
    config = load_runtime_config()
except Exception as error:
    st.error(f"Reconciliation halted: {error}")
    st.exception(error)
    st.stop()

run_dates = tuple(results)
if "selected_run_date" not in st.session_state or st.session_state.selected_run_date not in results:
    st.session_state.selected_run_date = run_dates[0]

st.subheader("Day-by-day chain")
rail = st.columns(len(run_dates))
for column, run_date in zip(rail, run_dates):
    summary = results[run_date].summary
    label = (
        f"{run_date}\n\n{summary['exceptions']} exceptions\n"
        f"{summary['new']} new · {summary['recurring']} recurring · ✓{summary['resolved']}"
    )
    if column.button(
        label,
        key=f"run_date_{run_date}",
        use_container_width=True,
        type="primary" if run_date == st.session_state.selected_run_date else "secondary",
    ):
        st.session_state.selected_run_date = run_date
        st.rerun()

st.caption(
    "Each day's register starts as a copy of the previous day's — "
    "New/Recurring/Resolved needs no database."
)

selected_date = st.session_state.selected_run_date
selected = results[selected_date]
summary = selected.summary

kpis = st.columns(5)
kpis[0].metric("Active requisitions", summary["active_count"])
kpis[1].metric("Field pairs checked", summary["pairs_checked"])
kpis[2].metric("Exceptions", summary["exceptions"], delta=f"{summary['high']} HIGH", delta_color="inverse")
kpis[3].metric("New / Recurring", f"{summary['new']} / {summary['recurring']}")
kpis[4].metric("Resolved this run", summary["resolved"])

register_tab, digests_tab, config_tab = st.tabs(
    ("Mismatch register", "Recruiter digests", "Config & rules")
)

with register_tab:
    frame = findings_frame(selected)
    severity_options = [severity for severity in ("HIGH", "MEDIUM", "LOW") if severity in set(frame.get("Severity", []))]
    status_options = [status for status in ("New", "Recurring") if status in set(frame.get("Status", []))]
    recruiter_options = sorted(set(frame.get("Recruiter", [])))

    filter_columns = st.columns((1, 1, 1.3, 2))
    severities = filter_columns[0].multiselect("Severity", severity_options, default=severity_options)
    statuses = filter_columns[1].multiselect("Status", status_options, default=status_options)
    recruiters = filter_columns[2].multiselect("Recruiter", recruiter_options, default=recruiter_options)
    search_text = filter_columns[3].text_input("Search", placeholder="Requisition, field, value…")

    filtered = filter_findings(frame, severities, statuses, recruiters, search_text)
    if filtered.empty:
        st.info("No mismatch findings match the selected filters.")
    else:
        display = filtered.copy()
        display["Values"] = display["Requisition Value"] + " ⟷ " + display["Position Value"]
        display = display[
            [
                "Requisition No",
                "Position Number",
                "Severity",
                "Status",
                "Recruiter",
                "Recruitment Stage",
                "Field",
                "Values",
                "Requisition Title",
                "Hiring Manager",
                "Exception Type",
            ]
        ]
        st.dataframe(style_register(display), use_container_width=True, hide_index=True)

    st.download_button(
        "Download this register (.xlsx)",
        data=register_to_bytes(selected, data_mode="library"),
        file_name=f"RCM_EC_Register_{selected_date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=False,
    )

    st.markdown('<div class="resolved-heading">✓ Resolved since previous run</div>', unsafe_allow_html=True)
    if selected.resolved:
        st.dataframe(pd.DataFrame(selected.resolved), use_container_width=True, hide_index=True)
    else:
        st.success("No issues were resolved since the previous run.")

with digests_tab:
    st.caption("These digests are what the scheduled M365 flow will email; the loop is not yet wired.")
    if not selected.digests:
        st.info("No recruiter digests for this run because there are no mismatch findings.")
    else:
        digest_names = tuple(selected.digests)
        digest_labels = {
            name: f"{name} — {selected.digests[name].count} items ({selected.digests[name].high_count} HIGH)"
            for name in digest_names
        }
        selected_recruiter = st.selectbox(
            "Recruiter",
            options=digest_names,
            format_func=digest_labels.get,
        )
        digest = selected.digests[selected_recruiter]
        subject = f"RCM–EC reconciliation: {digest.count} items need your review ({selected_date})"
        st.markdown(
            f'<div class="email-card"><div class="email-subject">Subject: {escape(subject)}</div>'
            f"{digest.html}</div>",
            unsafe_allow_html=True,
        )

with config_tab:
    st.subheader("Runtime field map")
    st.caption("Only rows with status = confirmed are compared. Pending rows are visible but inactive.")
    field_map = pd.DataFrame([pair.__dict__ for pair in config.field_map])

    def pending_style(row: pd.Series) -> list[str]:
        style = "color:#8B9098;background-color:#303238" if row["status"].lower() == "pending" else ""
        return [style] * len(row)

    st.dataframe(field_map.style.apply(pending_style, axis=1), use_container_width=True, hide_index=True)

    st.subheader("Value normalisers")
    normalisers = pd.DataFrame([normaliser.__dict__ for normaliser in config.normalisers])
    if normalisers.empty:
        st.info("No ValueNormalisers are configured.")
    else:
        st.dataframe(normalisers, use_container_width=True, hide_index=True)

    st.toggle("Apply ValueNormalisers", key="normalisers_on")
    country_false_positives = sum(
        finding["Field"] == "Country <> Country" for finding in selected.findings
    )
    if st.session_state.normalisers_on:
        st.success(f"Normaliser ON — {country_false_positives} Country false positives in this run.")
    else:
        st.warning(
            f"Normaliser OFF — {country_false_positives} Country false positives in this run; "
            "this is why ValueNormalisers exists."
        )

render_footer()
