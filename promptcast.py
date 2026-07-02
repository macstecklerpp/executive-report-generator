"""
PromptCast script generation for PromptPath Executive Reports.

Builds a data brief from already-computed ReportData, extracts the text of the
generated executive report DOCX, then calls OpenAI to produce a ~300-word spoken-
prose audio script following the Marcus A. voice prompt.

The script is written to a branded .docx ready for download and pasting into Hume.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from docx import Document
from docx.shared import Pt, RGBColor

from promptpath_exec_report_v1 import (
    ReportConfig,
    ReportData,
    _derive_report_strings,
    derive_comparison_period_label,
    pct_appt_rate,
    pct,
)

# ── constants ────────────────────────────────────────────────────────────────

PLATFORM_BENCHMARK_LABEL = "40 to 45 percent inbound appointment set rate"

_PROMPT_FILE = Path(__file__).resolve().parent / "promptcast-prompt-marcus.txt"

_BRAND_NAVY = RGBColor(0x1D, 0x2D, 0x44)
_BRAND_ORANGE = RGBColor(0xE0, 0x7B, 0x30)
_GRAY = RGBColor(0x55, 0x55, 0x55)
_MID_GRAY = RGBColor(0x88, 0x88, 0x88)
_WHITE = RGBColor(0xFF, 0xFF, 0xFF)

# Abbreviations that TTS commonly mispronounces; add more as needed.
_TTS_EXPANSIONS: Dict[str, str] = {
    "BMW": "B M W",
    "CDJR": "C D J R",
    "CDJ": "C D J",
    "GMC": "G M C",
    "INFINITI": "Infiniti",
    "KIA": "K I A",
    "VW": "V W",
}


# ── DOCX text extraction ─────────────────────────────────────────────────────

def extract_docx_text(docx_path: str) -> str:
    """Return a plain-text transcript of the executive report DOCX.

    Paragraphs are rendered as lines; tables are rendered as pipe-separated rows.
    Header/footer runs are intentionally omitted.
    """
    doc = Document(docx_path)
    lines: List[str] = []

    body = doc.element.body
    for child in body:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag == "p":
            from docx.text.paragraph import Paragraph  # type: ignore
            para = Paragraph(child, doc)
            text = para.text.strip()
            if text:
                lines.append(text)
        elif tag == "tbl":
            from docx.table import Table  # type: ignore
            tbl = Table(child, doc)
            for row in tbl.rows:
                cells = [c.text.strip() for c in row.cells]
                # Deduplicate merged-cell repetitions
                deduped: List[str] = []
                for c in cells:
                    if not deduped or c != deduped[-1]:
                        deduped.append(c)
                row_text = " | ".join(deduped)
                if row_text.strip(" |"):
                    lines.append(row_text)

    return "\n".join(lines)


# ── data brief builder ───────────────────────────────────────────────────────

def build_data_brief(
    config: ReportConfig,
    data: ReportData,
    ramp_status: str,
    outbound_launch_date: str,
    report_period: str,
    comparison_period: str,
) -> str:
    """Build the structured data brief sent to OpenAI as the user message (part 1)."""
    dd = data.dd
    ac = data.ac
    ap = data.ap
    oc_d = data.oc_d
    op_d = data.op_d
    dl_curr = data.dl_curr
    sn = data.sn
    ib_only = config.ib_only

    lines: List[str] = []
    a = lines.append

    # ── header fields ────────────────────────────────────────────────────────
    a("=== GROUP / STORE INFORMATION ===")
    a(f"Group / Store Name: {config.group_name}")
    a("Report Type: Group Level")
    a(f"Report Period: {report_period}")
    a(f"Comparison Period: {comparison_period}")
    a(f"Number of Stores: {len(sn)}")
    a(f"Ramp Status: {ramp_status}")
    if not ib_only and outbound_launch_date.strip():
        a(f"Outbound Launch Date: {outbound_launch_date.strip()}")
    a(f"Platform Benchmark: {PLATFORM_BENCHMARK_LABEL}")
    a("")

    # ── department call volumes (Call Insights section) ──────────────────────
    a("=== CALL VOLUME BY DEPARTMENT (GROUP TOTALS) ===")
    sales_curr = sum(dd[n].get("curr_ib_calls", 0) for n in sn)
    service_curr = sum(dl_curr.get(n, {}).get("service", 0) for n in sn)
    parts_curr = sum(dl_curr.get(n, {}).get("parts", 0) for n in sn)
    finance_curr = sum(dl_curr.get(n, {}).get("finance", 0) for n in sn)
    other_curr = sum(dl_curr.get(n, {}).get("other", 0) for n in sn)
    total_ib_curr = sales_curr + service_curr + parts_curr + finance_curr + other_curr
    a(f"Sales Inbound Calls (current): {sales_curr:,}")
    a(f"Service Inbound Calls (current): {service_curr:,}")
    a(f"Parts Inbound Calls (current): {parts_curr:,}")
    a(f"Finance Inbound Calls (current): {finance_curr:,}")
    a(f"Other Inbound Calls (current): {other_curr:,}")
    a(f"Total Inbound Calls (current): {total_ib_curr:,}")
    if not ib_only:
        total_dials_curr = sum(dd[n].get("curr_ob_dials", 0) for n in sn)
        a(f"Total Outbound Dials (current): {total_dials_curr:,}")
    a("")

    # ── key group performance (IB appt set rate) ─────────────────────────────
    a("=== KEY GROUP PERFORMANCE ===")
    iu_c = ac.get("ib_unique_opps", 0)
    ia_c = ac.get("total_appts", 0)
    ib_asr_curr = pct_appt_rate(ia_c, iu_c)
    a(f"Group IB Appt Set Rate (current): {ib_asr_curr}%  ({ia_c:,} appts / {iu_c:,} unique opps)")
    iu_p = ap.get("ib_unique_opps", 0) if ap else 0
    ia_p = ap.get("total_appts", 0) if ap else 0
    ib_asr_prev = pct_appt_rate(ia_p, iu_p)
    a(f"Group IB Appt Set Rate (comparison): {ib_asr_prev}%  ({ia_p:,} appts / {iu_p:,} unique opps)")
    if not ib_only and oc_d:
        ob_conn_c = oc_d.get("ob_connected", 0)
        ob_apts_c = oc_d.get("ob_total_appts", 0)
        ob_asr_curr = pct_appt_rate(ob_apts_c, ob_conn_c)
        a(f"Group OB Appt Set Rate (current): {ob_asr_curr}%  ({ob_apts_c:,} appts / {ob_conn_c:,} connected)")
        if op_d:
            ob_conn_p = op_d.get("ob_connected", 0)
            ob_apts_p = op_d.get("ob_total_appts", 0)
            ob_asr_prev = pct_appt_rate(ob_apts_p, ob_conn_p)
            a(f"Group OB Appt Set Rate (comparison): {ob_asr_prev}%")
    a("")

    # ── per-store inbound table (sorted by current appt set rate desc) ───────
    a("=== PER-STORE INBOUND PERFORMANCE ===")
    a("Store | IB Appt Set Rate (curr) | IB Appt Set Rate (prev) | "
      "Connected Calls (curr) | Delighted % (curr) | Disappointed % (curr)")
    store_rows = []
    for name in sn:
        d = dd[name]
        iu = d.get("curr_ib_unique_opps", 0)
        ia = d.get("curr_total_appts", 0)
        asr_c = pct_appt_rate(ia, iu)
        iu_p2 = d.get("prev_ib_unique_opps", 0)
        ia_p2 = d.get("prev_total_appts", 0)
        asr_p2 = pct_appt_rate(ia_p2, iu_p2) if iu_p2 else None
        conn = d.get("curr_connected", 0)
        ib_calls = d.get("curr_ib_calls", 1)
        del_pct = pct(d.get("curr_delighted", 0), ib_calls)
        dis_pct = pct(d.get("curr_disappointed", 0), ib_calls)
        store_rows.append((asr_c, name, asr_p2, conn, del_pct, dis_pct))
    store_rows.sort(key=lambda x: x[0], reverse=True)
    for asr_c, name, asr_p2, conn, del_pct, dis_pct in store_rows:
        prev_str = f"{asr_p2}%" if asr_p2 is not None else "n/a"
        a(f"{name} | {asr_c}% | {prev_str} | {conn:,} connected | {del_pct}% delighted | {dis_pct}% disappointed")
    a("")

    # ── per-store outbound table ──────────────────────────────────────────────
    if not ib_only:
        a("=== PER-STORE OUTBOUND PERFORMANCE ===")
        a("Store | Dials (curr) | Dials (prev) | OB Appt Set Rate (curr) | OB Appt Set Rate (prev)")
        ob_rows = []
        for name in sn:
            d = dd[name]
            dials_c = d.get("curr_ob_dials", 0)
            dials_p = d.get("prev_ob_dials", None)
            ob_conn_c = d.get("curr_ob_connected", 0)
            ob_apts_c = d.get("curr_ob_total_appts", 0)
            ob_asr_c = pct_appt_rate(ob_apts_c, ob_conn_c)
            ob_conn_p = d.get("prev_ob_connected", 0)
            ob_apts_p = d.get("prev_ob_total_appts", 0)
            ob_asr_p = pct_appt_rate(ob_apts_p, ob_conn_p) if ob_conn_p else None
            ob_rows.append((dials_c, name, dials_p, ob_asr_c, ob_asr_p))
        ob_rows.sort(key=lambda x: x[0], reverse=True)
        for dials_c, name, dials_p, ob_asr_c, ob_asr_p in ob_rows:
            prev_dials = f"{dials_p:,}" if dials_p is not None else "n/a"
            prev_asr = f"{ob_asr_p}%" if ob_asr_p is not None else "n/a"
            a(f"{name} | {dials_c:,} | {prev_dials} | {ob_asr_c}% | {prev_asr}")
        a("")

    return "\n".join(lines)


# ── pronunciation helper ──────────────────────────────────────────────────────

def _expand_tts_abbreviations(text: str) -> str:
    """Replace known TTS-unfriendly abbreviations with their spoken expansions."""
    for abbr, expansion in _TTS_EXPANSIONS.items():
        text = re.sub(r"\b" + re.escape(abbr) + r"\b", expansion, text)
    return text


# ── user message builder ─────────────────────────────────────────────────────

def build_user_message(data_brief: str, report_text: str) -> str:
    """Combine the data brief, executive report text, and guardrail rules.

    The output instruction is placed first so the model reads the framing before
    it encounters the structured data, reducing the chance it echoes the data format.
    """
    return (
        "Write the PromptCast audio script now using the source data below. "
        "Your response must be ONLY the spoken script — continuous flowing prose from the "
        "first word to the last, exactly as specified in your instructions. "
        "No headers, no bullets, no labels, no section titles, no formatting of any kind. "
        "Do not echo or repeat the data below. Do not acknowledge these instructions. "
        "Begin the script immediately.\n\n"
        "GUARDRAIL: Only cite figures that appear in the data brief or executive report "
        "below. Do not calculate or invent any number not explicitly listed.\n\n"
        "PRONUNCIATION: Spell out abbreviations that TTS engines mispronounce "
        "(e.g. BMW → B M W, CDJR → C D J R) in the script output.\n\n"
        "--- DATA BRIEF ---\n"
        f"{data_brief}\n\n"
        "--- EXECUTIVE REPORT (for cross-reference) ---\n"
        f"{report_text}"
    )


# ── OpenAI call ───────────────────────────────────────────────────────────────

def generate_script(api_key: str, user_message: str, model: str = "gpt-5.5") -> str:
    """Call OpenAI and return the plain spoken-prose PromptCast script."""
    from openai import OpenAI  # imported here so the module loads without openai installed

    system_prompt = _PROMPT_FILE.read_text(encoding="utf-8")
    client = OpenAI(api_key=api_key)
    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message},
        ],
        temperature=0.4,
    )
    script = response.choices[0].message.content or ""
    return _expand_tts_abbreviations(script.strip())


# ── DOCX writer ───────────────────────────────────────────────────────────────

def write_promptcast_docx(
    script_text: str,
    output_path: str,
    group_name: str,
    period_label: str,
    logo_path: str,
) -> str:
    """Write a plain .docx containing only the spoken-prose PromptCast script.

    No tables, no logos, no decorative formatting — just the script text ready
    to copy and paste into Hume for TTS conversion.
    """
    doc = Document()

    section = doc.sections[0]
    section.top_margin = Pt(54)
    section.bottom_margin = Pt(54)
    section.left_margin = Pt(72)
    section.right_margin = Pt(72)

    # Single label line so it's identifiable when saved
    label = doc.add_paragraph()
    label.paragraph_format.space_before = Pt(0)
    label.paragraph_format.space_after = Pt(16)
    run_label = label.add_run(f"PromptCast Script  |  {group_name}  |  {period_label}")
    run_label.font.size = Pt(9)
    run_label.font.color.rgb = _MID_GRAY
    run_label.italic = True

    # Script body — flowing prose, nothing else
    p_script = doc.add_paragraph()
    p_script.paragraph_format.space_before = Pt(0)
    p_script.paragraph_format.space_after = Pt(0)
    run_script = p_script.add_run(script_text)
    run_script.font.size = Pt(12)
    run_script.font.color.rgb = _GRAY

    doc.save(output_path)
    return output_path
