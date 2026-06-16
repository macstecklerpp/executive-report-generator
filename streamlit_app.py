"""
PromptPath Executive Summary Report — Streamlit UI

Run locally:
  pip install -r requirements.txt
  streamlit run streamlit_app.py
"""

from __future__ import annotations

import base64
import csv
import html as html_lib
import io
import os
import re
import secrets as secrets_stdlib
import tempfile
import zipfile
from datetime import date
from pathlib import Path
from typing import Any, Optional

import streamlit as st

try:
    import docx  # noqa: F401  # package name: python-docx
    import openpyxl  # noqa: F401
    import resend
except ModuleNotFoundError as e:
    st.set_page_config(page_title="PromptPath Executive Report", layout="centered")
    st.error(f"Missing dependency: {e}")
    st.markdown(
        "Install dependencies into your active Python environment:\n\n"
        "```bash\n"
        "cd /path/to/CS-BiWeeklyReports\n"
        "python3 -m venv .venv\n"
        "source .venv/bin/activate\n"
        "pip install -r requirements.txt\n"
        "streamlit run streamlit_app.py\n"
        "```\n\n"
        "If you already have a `.venv` in this project, activate it first (`source .venv/bin/activate`) "
        "or run directly with `.venv/bin/streamlit run streamlit_app.py`.\n\n"
        "On Streamlit Cloud, ensure `requirements.txt` is committed at the repo root, then reboot the app "
        "and check **Manage app → Logs** for pip errors."
    )
    st.stop()

from promptpath_exec_report_v1 import (
    LISTENED_LINE_TYPE_OPTIONS,
    ReportConfig,
    _derive_report_strings,
    generate_dealer_reports,
    generate_report,
    normalize_store_filters,
    report_output_basename,
    validate_dept_csv,
    validate_inbound_csv,
    validate_outbound_csv,
)

_APP_DIR = Path(__file__).resolve().parent
_USERS_CSV_PATH = _APP_DIR / "Admin - User List - Sheet1.csv"
_TO_ADDRESS = "macalister@promptpath.ai"
_DEFAULT_FROM = "onboarding@resend.dev"
_LOGO_CID = "promptpath-logo"


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


def _maybe_sidebar_actions() -> None:
    with st.sidebar:
        if _collect_allowed_passwords() and st.session_state.get("_pp_authenticated"):
            st.caption("Session")
            if st.button("Sign out"):
                st.session_state._pp_authenticated = False
                st.rerun()

        if st.session_state.get("pp_reports"):
            st.caption("Results")
            if st.button("Clear results"):
                st.session_state.pp_reports = []
                st.rerun()


def _safe_docx_name(name: str, default: str = "PromptPath_Report.docx") -> str:
    base = (name or "").strip() or default
    if not base.lower().endswith(".docx"):
        base += ".docx"
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


def _secret_or_env(secret_key: str, env_var: str) -> str:
    val = os.environ.get(env_var, "").strip()
    if val:
        return val
    try:
        return str(st.secrets.get(secret_key, "") or "").strip()
    except Exception:
        return ""


def _get_resend_api_key() -> Optional[str]:
    key = _secret_or_env("resend_api_key", "RESEND_API_KEY")
    return key or None


def _get_resend_from() -> str:
    return _secret_or_env("resend_from", "RESEND_FROM") or _DEFAULT_FROM


def _configure_resend() -> bool:
    key = _get_resend_api_key()
    if not key:
        return False
    resend.api_key = key
    return True


def _load_logo_bytes() -> Optional[bytes]:
    logo_path = _default_logo_path()
    if logo_path is None:
        return None
    return logo_path.read_bytes()


def _logo_inline_attachment() -> Optional[dict[str, Any]]:
    raw = _load_logo_bytes()
    if raw is None:
        return None
    return {
        "filename": "PromptPath_Logo.png",
        "content": base64.b64encode(raw).decode("ascii"),
        "content_type": "image/png",
        "content_id": _LOGO_CID,
    }


def _build_email_attachments(*file_attachments: dict[str, Any]) -> list[dict[str, Any]]:
    attachments: list[dict[str, Any]] = []
    logo = _logo_inline_attachment()
    if logo:
        attachments.append(logo)
    attachments.extend(file_attachments)
    return attachments


def _safe_store_name(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9_\- ]", "_", name.strip()).strip() or "dealer"


def _first_name(full_name: str) -> str:
    parts = (full_name or "").strip().split()
    return parts[0] if parts else "there"


def _load_user_csv() -> tuple[dict[str, list[tuple[str, str]]], dict[str, str]]:
    """Return store_users and store_originals keyed by safe store filename stem."""
    store_users: dict[str, list[tuple[str, str]]] = {}
    store_originals: dict[str, str] = {}

    if not _USERS_CSV_PATH.is_file():
        return store_users, store_originals

    with open(_USERS_CSV_PATH, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            dealer_raw = (row.get("Dealer Name") or "").strip()
            if not dealer_raw:
                continue
            full_name = (row.get("Full Name") or "").strip()
            email = (row.get("Email") or "").strip()
            if not full_name:
                continue

            dealers = [d.strip() for d in dealer_raw.split(",") if d.strip()]
            for dealer in dealers:
                safe = _safe_store_name(dealer)
                store_originals.setdefault(safe, dealer)
                users = store_users.setdefault(safe, [])
                entry = (_first_name(full_name), email)
                if entry not in users:
                    users.append(entry)

    return store_users, store_originals


def _build_email_html(
    greeting: str,
    body_paragraphs: list[str],
    bullet_items: Optional[list[str]] = None,
    include_logo: bool = False,
    include_questions_line: bool = True,
    footer_lines: Optional[list[str]] = None,
    sign_off: str = "Thanks, Mac",
) -> str:
    parts = ['<div style="font-family: Arial, sans-serif; max-width: 600px; color: #1D2D44;">']
    if include_logo:
        parts.append(
            f'<p><img src="cid:{_LOGO_CID}" alt="PromptPath" width="200" '
            'style="display:block; margin-bottom: 16px;"></p>'
        )
    parts.append(f"<p>{html_lib.escape(greeting)}</p>")
    for paragraph in body_paragraphs:
        parts.append(f"<p>{paragraph}</p>")
    if bullet_items:
        parts.append("<p><strong>What's inside this report:</strong></p><ul>")
        for item in bullet_items:
            parts.append(f"<li>{html_lib.escape(item)}</li>")
        parts.append("</ul>")
    if include_questions_line:
        parts.append("<p>Let me know if you have any questions!</p>")
    parts.append(f"<p>{html_lib.escape(sign_off)}</p>")
    if footer_lines:
        email_lines = "<br>".join(html_lib.escape(line) for line in footer_lines if line.strip())
        if email_lines:
            parts.append(
                '<p style="margin-top: 24px; font-size: 13px; color: #666;">'
                f"<strong>Send to:</strong><br>{email_lines}</p>"
            )
    parts.append("</div>")
    return "".join(parts)


def _send_group_report_email(result_entry: dict[str, Any]) -> None:
    if not result_entry.get("docx_bytes"):
        raise ValueError("Group report DOCX is missing.")
    if not _configure_resend():
        raise ValueError("Add resend_api_key to Streamlit secrets or RESEND_API_KEY env var.")

    group_name = result_entry["group_name"]
    period_label = result_entry.get("period_label") or "the reporting period"
    include_logo = _logo_inline_attachment() is not None
    html_body = _build_email_html(
        greeting=f"Hi {group_name} team,",
        body_paragraphs=[
            f"Attached is your bi-weekly Executive Performance Report covering {html_lib.escape(period_label)}."
        ],
        bullet_items=[
            "Call volume and trends vs. the prior period",
            "Appointment set rates (hard and soft) by calls and unique customers",
            "Customer sentiment breakdown (Delighted vs. Disappointed)",
        ],
        include_logo=include_logo,
    )
    attachment_name = result_entry.get("docx_name") or "PromptPath_Report.docx"
    resend.Emails.send(
        {
            "from": _get_resend_from(),
            "to": [_TO_ADDRESS],
            "subject": f"PromptPath Executive Performance Report — {group_name} ({period_label})",
            "html": html_body,
            "attachments": _build_email_attachments(
                {
                    "filename": attachment_name,
                    "content": base64.b64encode(result_entry["docx_bytes"]).decode("ascii"),
                }
            ),
        }
    )


def _store_key_for_docx_stem(
    stem: str,
    store_users: dict[str, list[tuple[str, str]]],
    store_originals: dict[str, str],
    period_start: Optional[str],
    period_end: Optional[str],
) -> Optional[str]:
    """Map a ZIP .docx stem to the user CSV store key (supports new dated filenames)."""
    if period_start and period_end:
        for safe, original in store_originals.items():
            if report_output_basename(original, period_start, period_end) == stem:
                return safe
    if stem in store_users:
        return stem
    return None


def _send_dealer_emails(result_entry: dict[str, Any]) -> tuple[int, int, list[str]]:
    if not result_entry.get("zip_bytes"):
        raise ValueError("Individual dealer ZIP is missing.")
    if not _configure_resend():
        raise ValueError("Add resend_api_key to Streamlit secrets or RESEND_API_KEY env var.")

    store_users, store_originals = _load_user_csv()
    period_label = result_entry.get("period_label") or "the reporting period"
    period_start = result_entry.get("period_start")
    period_end = result_entry.get("period_end")
    include_logo = _logo_inline_attachment() is not None
    sent = 0
    skipped = 0
    skipped_messages: list[str] = []

    with zipfile.ZipFile(io.BytesIO(result_entry["zip_bytes"])) as zf:
        docx_names = sorted(n for n in zf.namelist() if n.lower().endswith(".docx"))
        for arcname in docx_names:
            safe_name = _store_key_for_docx_stem(
                Path(arcname).stem,
                store_users,
                store_originals,
                period_start,
                period_end,
            )
            users = store_users.get(safe_name or "", [])
            if not users:
                skipped += 1
                display = store_originals.get(safe_name or "", Path(arcname).stem)
                skipped_messages.append(f"No users found for store: {display} ({arcname})")
                continue

            store_display = store_originals.get(safe_name or "", Path(arcname).stem.replace("_", " "))
            if len(users) > 1:
                greeting = "Hi Team,"
            else:
                greeting = f"Hi {users[0][0]},"
            recipient_emails = [email for _name, email in users if email.strip()]
            html_body = _build_email_html(
                greeting=greeting,
                body_paragraphs=[
                    (
                        f"Attached is your PromptPath report for "
                        f"<strong>{html_lib.escape(store_display)}</strong> "
                        f"covering {html_lib.escape(period_label)}."
                    )
                ],
                include_logo=include_logo,
                footer_lines=recipient_emails,
            )
            docx_bytes = zf.read(arcname)
            resend.Emails.send(
                {
                    "from": _get_resend_from(),
                    "to": [_TO_ADDRESS],
                    "subject": f"PromptPath Report — {store_display} ({period_label})",
                    "html": html_body,
                    "attachments": _build_email_attachments(
                        {
                            "filename": Path(arcname).name,
                            "content": base64.b64encode(docx_bytes).decode("ascii"),
                        }
                    ),
                }
            )
            sent += 1

    return sent, skipped, skipped_messages


def _snapshot_upload(upload: Optional[object]) -> bytes:
    if upload is None:
        return b""
    try:
        upload.seek(0)
        return upload.read()
    except Exception:
        return upload.getvalue() if hasattr(upload, "getvalue") else b""


def _next_group_id() -> int:
    n = int(st.session_state.get("pp_next_group_id", 0)) + 1
    st.session_state.pp_next_group_id = n
    return n


def _empty_group(gid: int) -> dict[str, Any]:
    return {
        "id": gid,
        "group_name": "",
        "period_start": date(2026, 5, 1),
        "period_end": date(2026, 5, 15),
        "ib_only": False,
        "store_filter": "",
        "listened_lines": list(LISTENED_LINE_TYPE_OPTIONS),
        "ib_blob": b"",
        "ob_blob": b"",
        "dept_blob": b"",
    }


def _init_batch_session_state() -> None:
    st.session_state.setdefault("pp_reports", [])
    st.session_state.setdefault("pp_next_group_id", 0)
    if "pp_groups" not in st.session_state or not st.session_state.pp_groups:
        gid = _next_group_id()
        st.session_state.pp_groups = [_empty_group(gid)]


def _group_idx(gid: int) -> Optional[int]:
    for i, g in enumerate(st.session_state.pp_groups):
        if g["id"] == gid:
            return i
    return None


def _init_group_widgets(g: dict[str, Any]) -> None:
    gid = g["id"]
    defaults: dict[str, Any] = {
        f"pp_g{gid}_group_name": g.get("group_name", ""),
        f"pp_g{gid}_period_start": g.get("period_start", date(2026, 5, 1)),
        f"pp_g{gid}_period_end": g.get("period_end", date(2026, 5, 15)),
        f"pp_g{gid}_ib_only": g.get("ib_only", False),
        f"pp_g{gid}_store_filter": g.get("store_filter", ""),
        f"pp_g{gid}_listened_lines": g.get("listened_lines", list(LISTENED_LINE_TYPE_OPTIONS)),
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val


def _on_upload_change(gid: int, blob_field: str, widget_key: str) -> None:
    idx = _group_idx(gid)
    if idx is None:
        return
    upload = st.session_state.get(widget_key)
    st.session_state.pp_groups[idx][blob_field] = _snapshot_upload(upload)


def _read_group_inputs(g: dict[str, Any]) -> dict[str, Any]:
    gid = g["id"]
    ib_upload = st.session_state.get(f"pp_g{gid}_ib_csv")
    ob_upload = st.session_state.get(f"pp_g{gid}_ob_csv")
    dept_upload = st.session_state.get(f"pp_g{gid}_dept_csv")
    ib_only = bool(st.session_state.get(f"pp_g{gid}_ib_only", False))

    ib_blob = _snapshot_upload(ib_upload) if ib_upload is not None else g.get("ib_blob", b"")
    dept_blob = _snapshot_upload(dept_upload) if dept_upload is not None else g.get("dept_blob", b"")
    ob_blob = b""
    if not ib_only:
        ob_blob = _snapshot_upload(ob_upload) if ob_upload is not None else g.get("ob_blob", b"")

    idx = _group_idx(gid)
    if idx is not None:
        st.session_state.pp_groups[idx]["ib_blob"] = ib_blob
        st.session_state.pp_groups[idx]["dept_blob"] = dept_blob
        st.session_state.pp_groups[idx]["ob_blob"] = ob_blob

    return {
        "id": gid,
        "group_name": str(st.session_state.get(f"pp_g{gid}_group_name", "")).strip(),
        "period_start": st.session_state.get(f"pp_g{gid}_period_start"),
        "period_end": st.session_state.get(f"pp_g{gid}_period_end"),
        "ib_only": ib_only,
        "store_filter_raw": str(st.session_state.get(f"pp_g{gid}_store_filter", "")),
        "listened_lines": list(st.session_state.get(f"pp_g{gid}_listened_lines", [])),
        "ib_blob": ib_blob,
        "ob_blob": ob_blob,
        "dept_blob": dept_blob,
    }


def _validate_group_inputs(cfg: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    label = cfg["group_name"] or f"Group {cfg['id']}"
    if not cfg["group_name"]:
        errors.append(f"{label}: enter a dealer group name.")
    if cfg["period_end"] < cfg["period_start"]:
        errors.append(f"{label}: period end must be on or after period start.")
    if not cfg["ib_blob"]:
        errors.append(f"{label}: upload the inbound leaderboard CSV.")
    if not cfg["dept_blob"]:
        errors.append(f"{label}: upload the department calls file.")
    if not cfg["ib_only"] and not cfg["ob_blob"]:
        errors.append(f"{label}: upload the outbound CSV or enable inbound-only.")
    return errors


def _error_result(group_name: str, message: str) -> dict[str, Any]:
    return {
        "group_name": group_name or "Unnamed group",
        "docx_bytes": None,
        "docx_name": None,
        "audit_bytes": None,
        "audit_name": None,
        "zip_bytes": None,
        "zip_name": None,
        "period_label": None,
        "period_start": None,
        "period_end": None,
        "error": message,
    }


def _generate_one_group(cfg: dict[str, Any]) -> dict[str, Any]:
    group_name = cfg["group_name"]
    display_name = group_name or f"Group {cfg['id']}"
    store_filter = normalize_store_filters(cfg["store_filter_raw"])
    period_start = cfg["period_start"].strftime("%Y-%m-%d")
    period_end = cfg["period_end"].strftime("%Y-%m-%d")
    out_fn = _safe_docx_name(f"{report_output_basename(group_name, period_start, period_end)}.docx")

    tmpdir = tempfile.mkdtemp(prefix="pp_report_")
    ib_path = os.path.join(tmpdir, "inbound.csv")
    dept_path = os.path.join(tmpdir, "dept.csv")

    with open(ib_path, "wb") as f:
        f.write(cfg["ib_blob"])
    with open(dept_path, "wb") as f:
        f.write(cfg["dept_blob"])

    validate_inbound_csv(ib_path)
    validate_dept_csv(dept_path)

    ob_path_opt: Optional[str] = None
    if not cfg["ib_only"]:
        ob_path_opt = os.path.join(tmpdir, "outbound.csv")
        with open(ob_path_opt, "wb") as f:
            f.write(cfg["ob_blob"])
        validate_outbound_csv(ob_path_opt)

    logo_path = _logo_path_from_upload(None, tmpdir)
    out_path = os.path.join(tmpdir, out_fn)

    report_cfg = ReportConfig(
        group_name=group_name,
        period_start=period_start,
        period_end=period_end,
        ib_csv_path=ib_path,
        ob_csv_path=ob_path_opt,
        dept_csv_path=dept_path,
        logo_path=logo_path,
        output_path=out_path,
        ib_only=cfg["ib_only"],
        store_filter=store_filter,
        listened_lines=cfg["listened_lines"],
    )

    generate_report(report_cfg)

    audit_name = Path(out_fn).stem + "_audit.xlsx"
    audit_path = os.path.join(tmpdir, audit_name)

    with open(out_path, "rb") as f:
        docx_bytes = f.read()

    audit_bytes: Optional[bytes] = None
    if os.path.isfile(audit_path):
        with open(audit_path, "rb") as f:
            audit_bytes = f.read()

    zip_bytes: Optional[bytes] = None
    zip_name = f"{Path(out_fn).stem}_individual_dealers.zip"
    dealer_dir = os.path.join(tmpdir, "dealers")
    try:
        dealer_paths = generate_dealer_reports(report_cfg, dealer_dir)
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for p in dealer_paths:
                zf.write(p, arcname=os.path.basename(p))
        zip_buf.seek(0)
        zip_bytes = zip_buf.read()
    except Exception:
        zip_bytes = None
        zip_name = None

    period_label, _gen_date = _derive_report_strings(period_start, period_end)

    return {
        "group_name": display_name,
        "docx_bytes": docx_bytes,
        "docx_name": out_fn,
        "audit_bytes": audit_bytes,
        "audit_name": audit_name if audit_bytes else None,
        "zip_bytes": zip_bytes,
        "zip_name": zip_name,
        "period_label": period_label,
        "period_start": period_start,
        "period_end": period_end,
        "error": None,
    }


def _run_batch() -> None:
    groups = st.session_state.pp_groups
    if not groups:
        st.error("Add at least one dealer group.")
        return

    st.session_state.pp_reports = []
    progress = st.progress(0.0, text="Starting batch generation…")
    total = len(groups)

    for i, g in enumerate(groups):
        cfg = _read_group_inputs(g)
        display_name = cfg["group_name"] or f"Group {cfg['id']}"
        progress.progress(i / total, text=f"Processing {display_name} ({i + 1} of {total})…")

        validation_errors = _validate_group_inputs(cfg)
        if validation_errors:
            st.session_state.pp_reports.append(
                _error_result(display_name, "\n".join(validation_errors))
            )
            continue

        try:
            result = _generate_one_group(cfg)
            st.session_state.pp_reports.append(result)
        except ValueError as ve:
            st.session_state.pp_reports.append(_error_result(display_name, str(ve)))
        except FileNotFoundError as e:
            st.session_state.pp_reports.append(_error_result(display_name, f"File not found: {e}"))
        except Exception as e:
            st.session_state.pp_reports.append(
                _error_result(display_name, f"Unexpected error: {e}")
            )

    progress.progress(1.0, text="Batch generation complete.")


def _render_group_card(g: dict[str, Any], idx: int, can_remove: bool) -> None:
    gid = g["id"]
    _init_group_widgets(g)

    title = str(st.session_state.get(f"pp_g{gid}_group_name", "")).strip() or f"Group {idx + 1}"
    with st.expander(title, expanded=True):
        if can_remove:
            if st.button("Remove group", key=f"pp_g{gid}_remove"):
                st.session_state.pp_groups.pop(idx)
                st.rerun()

        st.text_input("Dealer group name", key=f"pp_g{gid}_group_name")
        c1, c2 = st.columns(2)
        with c1:
            st.date_input("Period start", key=f"pp_g{gid}_period_start")
        with c2:
            st.date_input("Period end", key=f"pp_g{gid}_period_end")

        st.checkbox(
            "Inbound-only report (no outbound metrics; hides Sales Dials column)",
            key=f"pp_g{gid}_ib_only",
        )

        st.text_area(
            "Optional store filters (substring match; leave empty for all stores)",
            key=f"pp_g{gid}_store_filter",
            height=110,
            placeholder="All Star\nGenesis Baton Rouge",
            help=(
                "One substring per line, or separate with commas or semicolons. "
                "A store is included if its name matches any of these (OR)."
            ),
        )

        ib_key = f"pp_g{gid}_ib_csv"
        ob_key = f"pp_g{gid}_ob_csv"
        dept_key = f"pp_g{gid}_dept_csv"

        st.file_uploader(
            "Inbound leaderboard CSV (required)",
            type=["csv"],
            key=ib_key,
            on_change=_on_upload_change,
            args=(gid, "ib_blob", ib_key),
        )
        st.file_uploader(
            "Outbound leaderboard CSV (required unless inbound-only is checked)",
            type=["csv"],
            disabled=bool(st.session_state.get(f"pp_g{gid}_ib_only", False)),
            help="Ignored when inbound-only is enabled.",
            key=ob_key,
            on_change=_on_upload_change,
            args=(gid, "ob_blob", ob_key),
        )
        st.file_uploader(
            "Department calls CSV or TSV",
            type=["csv", "tsv", "txt"],
            key=dept_key,
            help=(
                "Long format: dealer_name, category, calls. Or wide leaderboard export with "
                "Dealerships, Period (Current/Previous), and Service/Parts/Finance/Other Inbound Calls. "
                "Current + Previous rows enable month-over-month % under each call type."
            ),
            on_change=_on_upload_change,
            args=(gid, "dept_blob", dept_key),
        )

        st.multiselect(
            "Lines PromptPath currently listens to",
            options=list(LISTENED_LINE_TYPE_OPTIONS),
            key=f"pp_g{gid}_listened_lines",
            help='Select all line types that apply. All selected = "all your lines" in the report.',
        )

        if g.get("ib_blob"):
            st.caption("Inbound CSV loaded.")
        if g.get("dept_blob"):
            st.caption("Department CSV loaded.")
        if g.get("ob_blob"):
            st.caption("Outbound CSV loaded.")


def _render_results() -> None:
    reports = st.session_state.get("pp_reports") or []
    if not reports:
        return

    st.divider()
    ok = sum(1 for r in reports if not r.get("error"))
    failed = len(reports) - ok
    st.subheader(f"Results — {len(reports)} group(s)")
    if failed:
        st.warning(f"{ok} succeeded, {failed} failed.")
    else:
        st.success(f"All {ok} report(s) generated successfully.")

    for i, entry in enumerate(reversed(reports)):
        with st.expander(entry["group_name"], expanded=True):
            if entry.get("error"):
                st.error(entry["error"])
                continue

            c1, c2, c3 = st.columns(3)
            with c1:
                st.download_button(
                    label="Download report (.docx)",
                    data=entry["docx_bytes"],
                    file_name=entry["docx_name"] or "PromptPath_Report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"dl_docx_{i}_{entry['group_name']}",
                )
            with c2:
                if entry.get("audit_bytes"):
                    st.download_button(
                        label="Download audit (.xlsx)",
                        data=entry["audit_bytes"],
                        file_name=entry["audit_name"] or "PromptPath_Report_audit.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_audit_{i}_{entry['group_name']}",
                    )
            with c3:
                if entry.get("zip_bytes"):
                    st.download_button(
                        label="Download individual dealers (.zip)",
                        data=entry["zip_bytes"],
                        file_name=entry["zip_name"] or "PromptPath_Report_individual_dealers.zip",
                        mime="application/zip",
                        key=f"dl_zip_{i}_{entry['group_name']}",
                    )

            st.markdown("**Email delivery**")
            email_key = f"{i}_{entry['group_name']}"
            if entry.get("docx_bytes"):
                if st.button("Send Group Report", key=f"send_group_{email_key}"):
                    if not _get_resend_api_key():
                        st.error("Add resend_api_key to Streamlit secrets or RESEND_API_KEY env var.")
                    else:
                        try:
                            with st.spinner("Sending group report..."):
                                _send_group_report_email(entry)
                            st.success(f"Group report sent to {_TO_ADDRESS}.")
                        except Exception as e:
                            st.error(f"Group report send failed: {e}")

            if entry.get("zip_bytes"):
                if st.button("Send Individual Emails", key=f"send_indiv_{email_key}"):
                    if not _get_resend_api_key():
                        st.error("Add resend_api_key to Streamlit secrets or RESEND_API_KEY env var.")
                    else:
                        try:
                            with st.spinner("Sending individual emails..."):
                                sent, skipped, skipped_msgs = _send_dealer_emails(entry)
                            st.success(f"Sent {sent} individual email(s) to {_TO_ADDRESS}.")
                            if skipped_msgs:
                                st.warning(
                                    f"{skipped} store(s) skipped:\n" + "\n".join(f"- {m}" for m in skipped_msgs)
                                )
                        except Exception as e:
                            st.error(f"Individual email send failed: {e}")


def main() -> None:
    st.set_page_config(page_title="PromptPath Executive Report", layout="centered")
    _ensure_password_gate()
    _maybe_sidebar_actions()
    st.title("PromptPath Executive Summary Report")
    st.caption(
        "Configure one or more dealer groups below, upload each group's CSVs, then click "
        "**Generate All Reports**. Each group produces a recap .docx, a number audit Excel file "
        "(_audit.xlsx), and a ZIP of individual dealer reports."
    )

    _init_batch_session_state()

    groups = st.session_state.pp_groups
    for idx, g in enumerate(groups):
        _render_group_card(g, idx, can_remove=len(groups) > 1)

    col_add, col_gen = st.columns([1, 2])
    with col_add:
        if st.button("+ Add Another Group"):
            st.session_state.pp_groups.append(_empty_group(_next_group_id()))
            st.rerun()
    with col_gen:
        if st.button("Generate All Reports", type="primary"):
            _run_batch()
            st.rerun()

    _render_results()


if __name__ == "__main__":
    main()
