"""
PromptPath Executive Summary Report — Streamlit UI

Run locally:
  pip install -r requirements.txt
  streamlit run streamlit_app.py
"""

from __future__ import annotations

import io
import os
import re
import secrets as secrets_stdlib
import tempfile
import zipfile
from datetime import date
from pathlib import Path
from typing import Optional

import streamlit as st

try:
    import docx  # noqa: F401  # package name: python-docx
    import openpyxl  # noqa: F401
except ModuleNotFoundError as e:
    st.set_page_config(page_title="PromptPath Executive Report", layout="centered")
    st.error(
        f"Missing dependency: {e}. Streamlit Cloud installs packages from **requirements.txt** "
        "at the repository root (needs `python-docx` and `openpyxl`). Commit that file, reboot the app, "
        "and check **Manage app → Logs** for pip errors."
    )
    st.stop()

from promptpath_exec_report_v1 import (
    LISTENED_LINE_TYPE_OPTIONS,
    ReportConfig,
    generate_dealer_reports,
    generate_report,
    normalize_store_filters,
    validate_dept_csv,
    validate_inbound_csv,
    validate_outbound_csv,
)

_APP_DIR = Path(__file__).resolve().parent


def _default_logo_path() -> Optional[Path]:
    """Linux (Streamlit Cloud) is case-sensitive; repo file is PromptPath_Logo.png."""
    for name in ("PromptPath_Logo.png", "Promptpath_Logo.png"):
        p = _APP_DIR / name
        if p.is_file():
            return p
    return None


def _collect_allowed_passwords() -> tuple[str, ...]:
    """Plain shared secret(s) from env or Streamlit secrets. Empty tuple = no login gate."""
    raw: list[str] = []
    env_one = os.environ.get("PROMPTPATH_APP_PASSWORD", "").strip()
    if env_one:
        raw.append(env_one)
    env_many = os.environ.get("PROMPTPATH_APP_PASSWORDS", "").strip()
    if env_many:
        raw.extend(p.strip() for p in env_many.split(",") if p.strip())
    try:
        sec = st.secrets
        one = str(sec.get("app_password", "") or "").strip()
        if one:
            raw.append(one)
        many = sec.get("app_passwords")
        if isinstance(many, str):
            raw.extend(p.strip() for p in many.split(",") if p.strip())
        elif isinstance(many, (list, tuple)):
            raw.extend(str(x).strip() for x in many if str(x).strip())
    except Exception:
        pass
    seen: dict[str, None] = {}
    for p in raw:
        seen.setdefault(p, None)
    return tuple(seen.keys())


def _password_matches(candidate: str, stored: str) -> bool:
    if len(candidate) != len(stored):
        return False
    try:
        return secrets_stdlib.compare_digest(candidate.encode("utf-8"), stored.encode("utf-8"))
    except Exception:
        return False


def _ensure_password_gate() -> None:
    allowed = _collect_allowed_passwords()
    if not allowed:
        return
    if st.session_state.get("_pp_authenticated"):
        return
    st.title("PromptPath Executive Summary Report")
    st.info(
        "This deployment requires the shared password. Ask your admin if you do not have it."
    )
    pw = st.text_input("Password", type="password", autocomplete="current-password")
    if st.button("Sign in", type="primary"):
        if any(_password_matches(pw, a) for a in allowed):
            st.session_state._pp_authenticated = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    st.stop()


def _maybe_sign_out_sidebar() -> None:
    if not _collect_allowed_passwords():
        return
    if not st.session_state.get("_pp_authenticated"):
        return
    with st.sidebar:
        st.caption("Session")
        if st.button("Sign out"):
            st.session_state._pp_authenticated = False
            st.rerun()


def _safe_docx_name(name: str, default: str = "PromptPath_Report.docx") -> str:
    base = (name or "").strip() or default
    if not base.lower().endswith(".docx"):
        base += ".docx"
    # Avoid path traversal and odd characters
    base = os.path.basename(base)
    base = re.sub(r'[^A-Za-z0-9._\- ]', "_", base).strip()
    return base or default


def _logo_path_from_upload(uploaded: Optional[object], tmpdir: str) -> str:
    if uploaded is not None:
        p = os.path.join(tmpdir, "logo.png")
        with open(p, "wb") as f:
            f.write(uploaded.getvalue())
        return p
    plogo = _default_logo_path()
    if plogo is not None:
        return str(plogo)
    raise FileNotFoundError(
        "No logo uploaded and PromptPath_Logo.png was not found beside this app. "
        "Place PromptPath_Logo.png in the project folder."
    )


def main() -> None:
    st.set_page_config(page_title="PromptPath Executive Report", layout="centered")
    _ensure_password_gate()
    _maybe_sign_out_sidebar()
    st.title("PromptPath Executive Summary Report")
    st.caption(
        "Upload leaderboard exports and department call volumes. Each run saves the recap .docx plus "
        "a companion number audit Excel file (_audit.xlsx) with raw CSV figures (All Dealers vs sum of dealers)."
    )

    st.session_state.setdefault("docx_bytes", None)
    st.session_state.setdefault("docx_name", None)
    st.session_state.setdefault("audit_bytes", None)
    st.session_state.setdefault("audit_name", None)
    st.session_state.setdefault("dealer_zip_bytes", None)
    st.session_state.setdefault("dealer_zip_name", None)

    with st.form("report_form"):
        group_name = st.text_input("Dealer group name", value="Example Automotive Group")
        c1, c2 = st.columns(2)
        with c1:
            period_start = st.date_input("Period start", value=date(2025, 4, 1))
        with c2:
            period_end = st.date_input("Period end", value=date(2025, 4, 28))

        ib_only = st.checkbox(
            "Inbound-only report (no outbound metrics; hides Sales Dials column)",
            value=False,
        )

        store_filter_raw = st.text_area(
            "Optional store filters (substring match; leave empty for all stores)",
            value="",
            height=110,
            placeholder="All Star\nGenesis Baton Rouge",
            help=(
                "One substring per line, or separate with commas or semicolons. "
                "A store is included if its name matches any of these (OR)."
            ),
        )

        out_name = st.text_input("Output filename", value="PromptPath_Report.docx")

        ib_csv = st.file_uploader("Inbound leaderboard CSV (required)", type=["csv"])
        ob_csv = st.file_uploader(
            "Outbound leaderboard CSV (required unless inbound-only is checked)",
            type=["csv"],
            disabled=ib_only,
            help="Ignored when inbound-only is enabled.",
        )

        dept_csv = st.file_uploader(
            "Department calls CSV or TSV (columns: dealer_name, category, calls)",
            type=["csv", "tsv", "txt"],
        )

        listened_lines = st.multiselect(
            "Lines PromptPath currently listens to",
            options=list(LISTENED_LINE_TYPE_OPTIONS),
            default=list(LISTENED_LINE_TYPE_OPTIONS),
            help="Select all line types that apply. All selected = \"all your lines\" in the report.",
        )

        submitted = st.form_submit_button("Generate DOCX")

    if not submitted:
        if st.session_state.docx_bytes:
            st.subheader("Download")
            st.download_button(
                label="Download last generated report",
                data=st.session_state.docx_bytes,
                file_name=st.session_state.docx_name or "PromptPath_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        if st.session_state.audit_bytes:
            st.download_button(
                label="Download last number audit (Excel)",
                data=st.session_state.audit_bytes,
                file_name=st.session_state.audit_name or "PromptPath_Report_audit.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        if st.session_state.dealer_zip_bytes:
            st.download_button(
                label="Download last individual dealer reports (ZIP)",
                data=st.session_state.dealer_zip_bytes,
                file_name=st.session_state.dealer_zip_name or "PromptPath_Report_individual_dealers.zip",
                mime="application/zip",
            )
        return

    errors: list[str] = []
    if not group_name.strip():
        errors.append("Enter a dealer group name.")
    if period_end < period_start:
        errors.append("Period end must be on or after period start.")
    if ib_csv is None:
        errors.append("Upload the inbound leaderboard CSV.")
    if dept_csv is None:
        errors.append("Upload the department calls file.")
    if not ib_only and ob_csv is None:
        errors.append("Upload the outbound CSV or enable inbound-only.")

    if errors:
        for e in errors:
            st.error(e)
        return

    store_filter = normalize_store_filters(store_filter_raw)
    out_fn = _safe_docx_name(out_name)

    tmpdir = tempfile.mkdtemp(prefix="pp_report_")
    ib_path = os.path.join(tmpdir, "inbound.csv")
    dept_path = os.path.join(tmpdir, "dept.csv")

    with open(ib_path, "wb") as f:
        f.write(ib_csv.getvalue())
    with open(dept_path, "wb") as f:
        f.write(dept_csv.getvalue())

    try:
        validate_inbound_csv(ib_path)
        validate_dept_csv(dept_path)
    except ValueError as ve:
        st.error(str(ve))
        return

    ob_path_opt: Optional[str] = None
    if not ib_only:
        ob_path_opt = os.path.join(tmpdir, "outbound.csv")
        with open(ob_path_opt, "wb") as f:
            f.write(ob_csv.getvalue())
        try:
            validate_outbound_csv(ob_path_opt)
        except ValueError as ve:
            st.error(str(ve))
            return

    try:
        logo_path = _logo_path_from_upload(None, tmpdir)
    except FileNotFoundError as e:
        st.error(str(e))
        return

    out_path = os.path.join(tmpdir, out_fn)

    cfg = ReportConfig(
        group_name=group_name.strip(),
        period_start=period_start.strftime("%Y-%m-%d"),
        period_end=period_end.strftime("%Y-%m-%d"),
        ib_csv_path=ib_path,
        ob_csv_path=ob_path_opt,
        dept_csv_path=dept_path,
        logo_path=logo_path,
        output_path=out_path,
        ib_only=ib_only,
        store_filter=store_filter,
        listened_lines=listened_lines,
    )

    st.session_state.dealer_zip_bytes = None
    st.session_state.dealer_zip_name = None

    try:
        generate_report(cfg)
    except FileNotFoundError as e:
        st.error(f"File not found: {e}")
        return
    except Exception as e:
        st.exception(e)
        return

    dealer_dir = os.path.join(tmpdir, "dealers")
    dealer_paths: list[str] = []
    try:
        dealer_paths = generate_dealer_reports(cfg, dealer_dir)
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for p in dealer_paths:
                zf.write(p, arcname=os.path.basename(p))
        zip_buf.seek(0)
        st.session_state.dealer_zip_bytes = zip_buf.read()
        st.session_state.dealer_zip_name = f"{Path(out_fn).stem}_individual_dealers.zip"
    except Exception as e:
        st.warning(f"Individual dealer reports could not be generated: {e}")

    audit_name = Path(out_fn).stem + "_audit.xlsx"
    audit_path = os.path.join(tmpdir, audit_name)

    with open(out_path, "rb") as f:
        st.session_state.docx_bytes = f.read()
    st.session_state.docx_name = out_fn
    if os.path.isfile(audit_path):
        with open(audit_path, "rb") as f:
            st.session_state.audit_bytes = f.read()
        st.session_state.audit_name = audit_name
    else:
        st.session_state.audit_bytes = None
        st.session_state.audit_name = None

    if st.session_state.audit_bytes:
        st.success(f"Generated: {out_fn} and {audit_name}")
    else:
        st.success(f"Generated: {out_fn}")
        st.warning("Number audit workbook was not created (check that openpyxl is installed).")
    nzip = bool(st.session_state.dealer_zip_bytes)
    if nzip:
        st.info(
            f"Individual dealer reports: ZIP with {len(dealer_paths)} store file(s) "
            f"({st.session_state.dealer_zip_name})."
        )
    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button(
            label="Download DOCX",
            data=st.session_state.docx_bytes,
            file_name=st.session_state.docx_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="dl_docx",
        )
    with c2:
        if st.session_state.audit_bytes:
            st.download_button(
                label="Download number audit (Excel)",
                data=st.session_state.audit_bytes,
                file_name=st.session_state.audit_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_audit",
            )
    with c3:
        if st.session_state.dealer_zip_bytes:
            st.download_button(
                label="Download individual dealers (ZIP)",
                data=st.session_state.dealer_zip_bytes,
                file_name=st.session_state.dealer_zip_name,
                mime="application/zip",
                key="dl_dealers",
            )


if __name__ == "__main__":
    main()
