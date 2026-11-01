import streamlit as st
import tempfile
from pathlib import Path
from datetime import date, timedelta
import traceback

from report_processor import run_daily_report

st.set_page_config(
    page_title="BSNL Warangal OA | Daily Provisioning Report",
    page_icon="📊",
    layout="wide",
)

st.markdown(
    """
    <style>
      .main-title {font-size: 1.7rem; font-weight: 750; color: #17365D;}
      .sub-title {color: #5B6573; font-size: 0.95rem;}
      div[data-testid="stMetric"] {
        background: #F4F7FB; border: 1px solid #DCE4EF;
        padding: 12px 14px; border-radius: 10px;
      }
      .stDownloadButton button {width: 100%;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="main-title">BSNL – WARANGAL OA</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">BBM-wise Daily Provisioning Report | Upload three Excel source files to generate the consolidated report.</div>',
    unsafe_allow_html=True,
)
st.divider()

# The processor reads these two mapping workbooks from the same directory as
# report_processor.py. Keep both files in the deployed GitHub repository.
processor_dir = Path(__file__).resolve().parent
required_masters = ["OLT_BBC_MAP.xlsx", "BBC_Master.xlsx"]
missing_masters = [name for name in required_masters if not (processor_dir / name).exists()]

with st.sidebar:
    st.header("Master file status")
    for name in required_masters:
        if (processor_dir / name).exists():
            st.success(f"{name} – available")
        else:
            st.error(f"{name} – missing")
    st.caption(
        "Both mapping workbooks must be placed in the same GitHub repository folder "
        "as app.py and report_processor.py."
    )
    st.divider()
    st.caption("Report layout: DAILY PROVISIONS REPORT.xlsx (27-column BBM-wise format)")

st.subheader("1. Upload source Excel files")
c1, c2, c3 = st.columns(3)

with c1:
    provisioning_file = st.file_uploader(
        "OLT Wise Provisioning",
        type=["xlsx", "xls"],
        help="Upload the OLT Wise Provisioning export.",
        key="provisioning_upload",
    )
with c2:
    uc_file = st.file_uploader(
        "UC Approved / BU Connection Status",
        type=["xlsx", "xls"],
        help="Upload the BU Connection Status Report containing UC approval details.",
        key="uc_upload",
    )
with c3:
    pending_file = st.file_uploader(
        "Pending OBS",
        type=["xlsx", "xls"],
        help="Upload the Pending OBS export.",
        key="pending_upload",
    )

st.subheader("2. Reporting date")
use_manual_date = st.checkbox(
    "Select reporting date manually",
    value=False,
    help="By default, the processor derives the reporting date from the OLT Wise Provisioning export timestamp.",
)
manual_report_date = None
if use_manual_date:
    manual_report_date = st.date_input(
        "Report as on",
        value=date.today() - timedelta(days=1),
        format="DD-MM-YYYY",
    )

st.caption(
    "If the manual date option is not selected, the report date is derived from the "
    "provisioning export timestamp (normally the previous day)."
)

st.subheader("3. Generate report")
ready = all([provisioning_file, uc_file, pending_file]) and not missing_masters

if missing_masters:
    st.warning(
        "The report cannot be generated until these mapping files are added to the "
        "repository: " + ", ".join(missing_masters)
    )

generate = st.button(
    "Generate Daily Provisioning Report",
    type="primary",
    use_container_width=True,
    disabled=not ready,
)

if generate:
    try:
        with st.spinner("Reading Excel files, mapping BBMs and preparing the report..."):
            with tempfile.TemporaryDirectory(prefix="bsnl_wgl_daily_") as temp_dir:
                temp_path = Path(temp_dir)

                def save_upload(uploaded, filename):
                    target = temp_path / filename
                    target.write_bytes(uploaded.getvalue())
                    return target

                provisioning_path = save_upload(
                    provisioning_file, "OLT_Wise_Provisioning.xlsx"
                )
                uc_path = save_upload(
                    uc_file, "UC_Connection_Status.xlsx"
                )
                pending_path = save_upload(
                    pending_file, "Pending_OBS.xlsx"
                )

                output_xlsx = temp_path / "BBM_Wise_Daily_Provisioning_Report.xlsx"
                output_html = temp_path / "BBM_Wise_Daily_Provisioning_Report.html"

                report_date_arg = manual_report_date if use_manual_date else None

                xlsx_path, html_path, metadata = run_daily_report(
                    provisioning_file=provisioning_path,
                    uc_file=uc_path,
                    pending_file=pending_path,
                    output_xlsx=output_xlsx,
                    output_html=output_html,
                    report_date=report_date_arg,
                )

                # Read the generated artifacts before the temporary directory closes.
                xlsx_bytes = Path(xlsx_path).read_bytes()
                html_bytes = Path(html_path).read_bytes() if html_path else None

        st.success("Daily Provisioning Report generated successfully.")

        st.subheader("Report summary")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Reporting date", metadata.get("report_date", "—"))
        m2.metric("BBMs covered", metadata.get("bbm_count", 0))
        m3.metric("Gross provisions", f'{metadata.get("gross_provision", 0):,}')
        m4.metric("Net", f'{metadata.get("net", 0):,}')

        st.download_button(
            "Download Excel Report",
            data=xlsx_bytes,
            file_name="BBM_Wise_Daily_Provisioning_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

        if html_bytes:
            st.download_button(
                "Download HTML Report",
                data=html_bytes,
                file_name="BBM_Wise_Daily_Provisioning_Report.html",
                mime="text/html",
                use_container_width=True,
            )

        if metadata.get("unmapped_uc_count", 0) or metadata.get("unmapped_pending_count", 0):
            st.warning(
                "Some UC or Pending OBS records could not be mapped to a BBM. "
                "Please review the Audit_Notes and raw input sheets in the Excel report."
            )

        st.info(
            "The Excel report includes the BBM-wise summary, TOTAL row, source-data sheets "
            "and Audit_Notes for reconciliation. Review the audit notes before circulating "
            "the report."
        )

    except Exception as exc:
        st.error(f"Unable to generate the report: {exc}")
        with st.expander("Technical details"):
            st.code(traceback.format_exc())

st.divider()
st.caption(
    "Internal reporting utility – Warangal OA. Upload files for the same reporting cycle "
    "to avoid date mismatches in consolidated totals."
)
