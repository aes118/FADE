# Grant Application Portal – Logframe + Workplan + Budget
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
import uuid
import base64
import os
import html, re
from html import escape
from datetime import datetime, date
import hashlib
import streamlit as st, requests, tempfile, os
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from docx.enum.section import WD_ORIENT, WD_SECTION_START


# ---------------- Page config ----------------
st.set_page_config(page_title="Falcon Awards Application Portal", layout="wide")
st.sidebar.image("glide_logo.png", width="stretch")

# ---------------- Helpers & State ----------------
def _s(v):
    try:
        import pandas as pd
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v).strip()

def generate_id():
    return str(uuid.uuid4())[:8]

footer_note = "Note: All monetary values are in USD and will be converted to AED in the official project agreement."

# app state
for key in ["impacts", "outcomes", "outputs", "kpis", "workplan", "budget", "disbursement"]:
    if key not in st.session_state:
        st.session_state[key] = []

# edit-state for inline editing
for key in ["edit_goal", "edit_outcome", "edit_output", "edit_kpi", "edit_activity", "edit_budget_row"]:
    if key not in st.session_state:
        st.session_state[key] = None

def _find_by_id(lst, _id):
    for i, x in enumerate(lst):
        if x.get("id") == _id:
            return i
    return None

def delete_cascade(*, goal_id=None, outcome_id=None, output_id=None):
    """Delete an item and its children in the Goal->Outcome->Output->KPI hierarchy."""
    if goal_id:
        for oc in [o for o in st.session_state.outcomes if o.get("parent_id") == goal_id]:
            delete_cascade(outcome_id=oc["id"])
        st.session_state.impacts = [g for g in st.session_state.impacts if g["id"] != goal_id]

    if outcome_id:
        for out in [o for o in st.session_state.outputs if o.get("parent_id") == outcome_id]:
            delete_cascade(output_id=out["id"])
        st.session_state.outcomes = [o for o in st.session_state.outcomes if o["id"] != outcome_id]

    if output_id:
        # delete KPIs under output
        st.session_state.kpis = [k for k in st.session_state.kpis if k.get("parent_id") != output_id]
        st.session_state.outputs = [o for o in st.session_state.outputs if o["id"] != output_id]

def render_editable_item(
    *,
    item: dict,
    list_name: str,
    edit_flag_key: str,
    view_md_func,
    fields=None,
    default_label="Name",
    on_delete=None,
    key_prefix="logframe"   # <— NEW
):
    c1, c2, c3 = st.columns([0.85, 0.07, 0.08])

    # unique widget id base
    wid = f"{key_prefix}_{list_name}_{item['id']}"

    if st.session_state.get(edit_flag_key) == item["id"]:
        new_values = {}
        if fields:
            for fkey, widget_func, label in fields:
                new_values[fkey] = widget_func(label, value=item.get(fkey, ""), key=f"{wid}_{fkey}")
        else:
            new_values["name"] = c1.text_input(default_label, value=item.get("name", ""), key=f"{wid}_name")

        if c2.button("💾", key=f"{wid}_save"):
            idx = _find_by_id(st.session_state[list_name], item["id"])
            if idx is not None:
                # Defensive: prevent users from saving names that include labels
                if "name" in new_values:
                    if list_name == "activities":
                        new_values["name"] = strip_label_prefix(new_values["name"], "Activity")
                    elif list_name == "kpis":
                        new_values["name"] = strip_label_prefix(new_values["name"], "KPI")

                for k, v in new_values.items():
                    st.session_state[list_name][idx][k] = v.strip() if isinstance(v, str) else v
            st.session_state[edit_flag_key] = None
            st.rerun()

        if c3.button("✖️", key=f"{wid}_cancel"):
            st.session_state[edit_flag_key] = None
            st.rerun()
    else:
        c1.markdown(view_md_func(item), unsafe_allow_html=True)
        if c2.button("✏️", key=f"{wid}_edit"):
            st.session_state[edit_flag_key] = item["id"]
            st.rerun()
        if c3.button("🗑️", key=f"{wid}_del"):
            if on_delete:
                on_delete()

from datetime import date

def compute_numbers(include_activities: bool = False):
    """
    Preserve user/excel order.
    Outputs: numbered per Outcome in list order.
    KPIs:    numbered per Output in list order (as stored in st.session_state.kpis).
    Activities (if requested): numbered per Output in list order.
    """
    out_num, kpi_num = {}, {}
    outcomes = st.session_state.get("outcomes", [])
    outputs  = st.session_state.get("outputs", [])
    kpis     = st.session_state.get("kpis", [])

    # Outputs numbered per Outcome (list order)
    for oc in outcomes:
        oc_outs = [o for o in outputs if o.get("parent_id") == oc["id"]]  # list order preserved
        for i, out in enumerate(oc_outs, start=1):
            out_num[out["id"]] = f"{i}"

    # KPIs numbered per Output (list order)
    for out_id, n in out_num.items():
        p = 1
        for k in kpis:                       # iterate as-is → preserves entry/import order
            if k.get("parent_id") == out_id:
                kpi_num[k["id"]] = f"{n}.{p}"
                p += 1

    if not include_activities:
        return out_num, kpi_num

    # Activities numbered per Output (list order)
    act_num = {}
    workplan = st.session_state.get("workplan", [])
    for out_id, n in out_num.items():
        q = 1
        for a in workplan:
            if a.get("output_id") == out_id:
                act_num[a["id"]] = f"{n}.{q}"
                q += 1
    return out_num, kpi_num, act_num

def _workplan_df():
    """Return a tidy DF with Activity, Output, Start, End (rows that have both dates)."""
    import pandas as pd
    outs = {o["id"]: (o.get("name") or "Output") for o in st.session_state.get("outputs", [])}
    rows = []
    for a in st.session_state.get("workplan", []):
        s, e = a.get("start"), a.get("end")
        if s and e:
            rows.append({
                "Activity": a.get("name",""),
                "Output": outs.get(a.get("output_id"), "(unassigned)"),
                "Start": s,
                "End": e,
            })
    if not rows:
        return pd.DataFrame(columns=["Activity","Output","Start","End"])
    return (
        pd.DataFrame(rows)
        .sort_values(["Output","Start","Activity"], kind="stable")
        .reset_index(drop=True)
    )

def _draw_gantt(ax, df, *, show_today=False):
    """
    Clean, readable Gantt:
    - Bars colored by Output with subtle edges
    - Group separators
    - Major ticks by quarter, minor ticks monthly; labels MMM-YYYY
    - Optional 'today' line
    - Bottom-centered legend
    """
    import numpy as np
    import matplotlib.pyplot as plt
    import matplotlib.dates as mdates
    from matplotlib.patches import Patch
    from datetime import date, timedelta

    if df.empty:
        ax.text(0.5, 0.5, "No activities with both start & end dates",
                ha="center", va="center")
        ax.axis("off")
        return

    # ---- Order outputs by earliest start for a logical reading order
    first_start = df.groupby("Output")["Start"].min().sort_values()
    ordered_outputs = first_start.index.tolist()
    df = (
        df.assign(_out_order=df["Output"].map({o:i for i, o in enumerate(ordered_outputs)}))
          .sort_values(["_out_order", "Start", "Activity"], kind="stable")
          .drop(columns="_out_order")
          .reset_index(drop=True)
    )

    # ---- Stable color mapping (tab20_cycle)
    cmap = plt.get_cmap("tab20")
    color_map = {o: cmap(i % 20) for i, o in enumerate(ordered_outputs)}

    # ---- Build bar geometry
    y      = np.arange(len(df))
    widths = [(e - s).days or 0.5 for s, e in zip(df["Start"], df["End"])]
    lefts  = df["Start"].tolist()
    colors = [color_map[o] for o in df["Output"]]

    # ---- Bars (with gentle edge for contrast)
    ax.barh(
        y, widths, left=lefts, color=colors,
        edgecolor="white", linewidth=0.8, alpha=0.95
    )

    # Optional: label durations at the end of each bar (uncomment to enable)
    # for yi, s, w in zip(y, lefts, widths):
    #     ax.text(s + timedelta(days=w) , yi, f" {int(w)}d",
    #             va="center", ha="left", fontsize=8, color="#444")

    # ---- Y axis: activities
    ax.set_yticks(y)
    ax.set_yticklabels(df["Activity"], fontsize=9)
    ax.invert_yaxis()
    ax.set_ylabel("Activity")

    # ---- X axis: quarterly major ticks + monthly minors; labels MMM-YYYY
    ax.xaxis.set_major_locator(mdates.MonthLocator(bymonth=(1,4,7,10)))  # Jan/Apr/Jul/Oct
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b-%Y"))
    ax.xaxis.set_minor_locator(mdates.MonthLocator())  # every month
    for lab in ax.get_xticklabels():
        lab.set_rotation(30); lab.set_ha("right")

    # ---- Range padding (7% of span on both sides)
    x_min = min(df["Start"]); x_max = max(df["End"])
    span  = (x_max - x_min).days or 1
    pad   = max(int(span * 0.07), 7)
    ax.set_xlim(x_min - timedelta(days=pad), x_max + timedelta(days=pad))

    # ---- Grid (minor vertical dots)
    ax.grid(axis="x", which="minor", linestyle=":", linewidth=0.6, alpha=0.35)
    ax.grid(axis="x", which="major", linestyle="-", linewidth=0.6, alpha=0.35)
    ax.set_xlabel("Date")

    # ---- Group separators (between Outputs)
    sep_rows = [i for i in range(1, len(df)) if df.loc[i, "Output"] != df.loc[i-1, "Output"]]
    xmin, xmax = ax.get_xlim()
    for i in sep_rows:
        ax.hlines(i - 0.5, xmin, xmax, colors="#BFC5CC", linestyles="-", linewidth=0.8, alpha=0.7)
    ax.set_xlim(xmin, xmax)  # keep full-width lines

    # ---- Optional 'today' line
    if show_today:
        today = date.today()
        ax.axvline(today, color="#666", linewidth=1.1, alpha=0.8)
        ax.text(today, ax.get_ylim()[0] - 0.25, "Today", rotation=90,
                va="top", ha="center", fontsize=8, color="#666")

    # ---- Legend (bottom center), tidy spacing
    handles = [Patch(facecolor=color_map[o], label=o) for o in ordered_outputs]
    ax.legend(
        handles=handles,
        title="Outputs",
        loc="upper center",
        bbox_to_anchor=(0.2, -0.30),
        ncol=2,  # wrap if many outputs
        # ncol=min(len(handles), 4),  # wrap if many outputs
        frameon=False,
        handlelength=1.8,
        columnspacing=1.4,
        handletextpad=0.6,
        borderaxespad=0.0,
    )

def select_output_id(label, current_id, key):
    """Select an Output and return its ID (options are IDs; stable + clean)."""
    outs = st.session_state.get("outputs", [])
    out_nums, _ = compute_numbers()
    id_to_name = {o["id"]: (o.get("name") or "Output") for o in outs}
    options = [o["id"] for o in outs]
    idx = options.index(current_id) if current_id in options else 0
    return st.selectbox(
        label,
        options,
        index=idx if options else 0,
        format_func=lambda oid: f"Output {out_nums.get(oid,'?')} — {id_to_name.get(oid,'Output')}",
        key=key,
    )

def strip_label_prefix(text: str, kind: str) -> str:
    """
    Remove labels like 'Activity 1.2 — ' or 'KPI 1.2.3: ' from a string.
    Accepts separators '—', ':', or '-'.
    """
    if not isinstance(text, str):
        return text
    pat = rf'^\s*{kind}\s+\d+(?:\.\d+)*\s*[—:\-]\s*'
    return re.sub(pat, '', text).strip()

def parse_date_like(v):
    """Return a datetime.date or None from common date formats or existing date/datetime/pandas types."""
    if v is None:
        return None

    # Handle pandas NaT / NaN early
    try:
        import pandas as pd
        if pd.isna(v):
            return None
        if isinstance(v, pd.Timestamp):
            return v.date()
    except Exception:
        pass

    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v

    # Strings
    s = str(v).strip()
    if not s or s.lower() in ("none", "nan", "nat"):
        return None

    # Try a bunch of common formats (with and without HH:MM:SS)
    fmts = (
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%d/%m/%Y",
        "%d/%m/%Y %H:%M:%S",
        "%m/%d/%Y",
        "%m/%d/%Y %H:%M:%S",
        "%Y/%m/%d",
        "%Y/%m/%d %H:%M:%S",
        "%d/%b/%Y",            # 03/Sep/2025
        "%d/%b/%Y %H:%M:%S",
        "%d-%b-%Y",            # 03-Sep-2025
        "%d-%b-%Y %H:%M:%S",
    )
    for fmt in fmts:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue

    return None

def fmt_dd_mmm_yyyy(v):
    """Return 'DD/MMM/YYYY' (e.g., 03/Sep/2025) or '' if not set/parsable."""
    try:
        import pandas as pd
        if pd.isna(v):                   # handles NaT / NaN
            return ""
        if isinstance(v, pd.Timestamp):  # valid timestamp -> format directly
            return v.strftime("%d/%b/%Y")
    except Exception:
        pass

    d = parse_date_like(v)
    return d.strftime("%d/%b/%Y") if d else ""

def fmt_money(val) -> str:
    """Return number with thousands dot and 2 decimals, e.g., 1.234.567,89."""
    try:
        x = float(val)
    except (TypeError, ValueError):
        return ""
    # First format in US style, then swap separators
    s = f"{x:,.2f}"                  # -> 1,234,567.89
    return s.replace(",", "␟").replace(".", ",").replace("␟", ".")

def view_logframe_element(inner_html: str, kind: str = "output") -> str:
    """Wrap inner HTML in a styled card. kind: 'output' | 'kpi' (or others later)."""
    return f"<div class='lf-card lf-card--{kind}'>{inner_html}</div>"

def sync_disbursement_from_kpis():
    """
    Keep st.session_state.disbursement in sync with KPIs:
    - rows only for KPIs with linked_payment == True
    - preserve user-entered anticipated_date and amount_usd
    - refresh output_id and kpi_name from KPI
    - if date is missing, prefill with latest linked Activity end, else KPI end, else KPI start
    """
    pay_kpis = [k for k in st.session_state.kpis if bool(k.get("linked_payment"))]
    by_id = {d.get("kpi_id"): d for d in st.session_state.disbursement}
    keep_ids = set()

    # helper: latest activity end for a KPI
    def latest_activity_end(kpi_id):
        dates = []
        for a in st.session_state.get("workplan", []):
            if kpi_id in (a.get("kpi_ids") or []):
                if a.get("end"):
                    dates.append(a["end"])
        return max(dates) if dates else None

    for k in pay_kpis:
        kid = k["id"]
        row = by_id.get(kid)
        lae = latest_activity_end(kid)
        if row is None:
            st.session_state.disbursement.append({
                "kpi_id": kid,
                "output_id": k.get("parent_id"),
                "kpi_name": k.get("name", ""),
                "anticipated_date": lae or k.get("end_date") or k.get("start_date") or None,
                "deliverable": k.get("name", ""),
                "amount_usd": 0.0,
            })
        else:
            # refresh mirror fields
            row["output_id"] = k.get("parent_id")
            row["kpi_name"]  = k.get("name", "")
            # if user hasn't chosen a date yet, backfill one
            if not row.get("anticipated_date"):
                row["anticipated_date"] = lae or k.get("end_date") or k.get("start_date") or None
        keep_ids.add(kid)

    # remove rows for KPIs no longer payment-linked
    st.session_state.disbursement = [d for d in st.session_state.disbursement if d.get("kpi_id") in keep_ids]

def build_logframe_docx():
    # Lazy import so app loads even if package is missing
    try:
        from docx import Document
        from docx.shared import Cm, RGBColor, Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.enum.section import WD_ORIENT
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import Pt
    except Exception:
        st.error("`python-docx` is required. In your venv run:\n  pip uninstall -y docx\n  pip install -U python-docx")
        raise

    PRIMARY_SHADE = "0A2F41"

    def _shade(cell, hex_fill):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), hex_fill)
        tcPr.append(shd)

    def _repeat_header(row):
        trPr = row._tr.get_or_add_trPr()
        tblHeader = OxmlElement("w:tblHeader")
        trPr.append(tblHeader)

    def _set_cell_text(cell, text, *, bold=False, white=False, align_left=True):
        cell.text = ""
        p = cell.paragraphs[0]
        if align_left:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(text or "")
        run.bold = bool(bold)
        run.font.name = "Calibri"
        run.font.size = Pt(11)
        try:
            rpr = run._element.rPr
            rpr.rFonts.set(qn("w:ascii"), "Calibri")
            rpr.rFonts.set(qn("w:hAnsi"), "Calibri")
            rpr.rFonts.set(qn("w:cs"), "Calibri")
        except Exception:
            pass
        if white:
            run.font.color.rgb = RGBColor(255, 255, 255)

    def _add_run(p, text, bold=False):
        run = p.add_run(text or "")
        run.bold = bold
        run.font.name = "Calibri"
        run.font.size = Pt(11)
        return run

    def _new_landscape_section(title):
        sec = doc.add_section(WD_SECTION_START.NEW_PAGE)
        sec.orientation = WD_ORIENT.LANDSCAPE
        # swap page size
        sec.page_width, sec.page_height = sec.page_height, sec.page_width
        _h1(title)
        return sec

    def _ensure_portrait_section(title):
        sec = doc.add_section(WD_SECTION_START.NEW_PAGE)
        sec.orientation = WD_ORIENT.PORTRAIT
        # ensure portrait proportions
        if sec.page_width > sec.page_height:
            sec.page_width, sec.page_height = sec.page_height, sec.page_width
        _h1(title)
        return sec

    def _content_width_cm(sec):
        # use .cm to avoid manual EMU conversion
        return sec.page_width.cm - sec.left_margin.cm - sec.right_margin.cm

    def _h1(text):
        p = doc.add_paragraph(text)
        p.style = doc.styles['Heading 1']
        return p

    def _new_section(title):
        # page/section break + H1
        doc.add_section(WD_SECTION_START.NEW_PAGE)
        _h1(title)

    from io import BytesIO

    def _gantt_png_buf():
        """Render the workplan Gantt to a PNG buffer for docx embedding."""
        import matplotlib.pyplot as plt
        df = _workplan_df()
        if df.empty:
            return None
        h = min(1.2 + 0.35 * len(df), 12)  # scale height with number of rows
        fig, ax = plt.subplots(figsize=(11, h))
        _draw_gantt(ax, df, show_today=False)
        # extra space for legend + rotated ticks
        fig.subplots_adjust(left=0.30, right=0.995, top=0.96, bottom=0.36)
        buf = BytesIO()
        # 'tight' ensures nothing gets cut off (legend, labels)
        fig.savefig(buf, format="png", dpi=220, bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        return buf

    # ---- Document setup: Portrait + standard margins
    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.PORTRAIT
    for side in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        setattr(sec, side, Cm(2.54))

    _h1("LOGFRAME")

    # ---- Data & numbering
    goal_text = (st.session_state.impacts[0]["name"] if st.session_state.get("impacts") else "")
    outcome_text = (st.session_state.outcomes[0]["name"] if st.session_state.get("outcomes") else "")
    out_nums, kpi_nums = compute_numbers()

    def _sort_by_num(label):
        if not label:
            return (9999,)
        try:
            return tuple(int(x) for x in str(label).split("."))
        except Exception:
            return (9999,)

    # ---- Helper: add a 2-row banner (label row shaded, content row unshaded)
    def _add_banner_block(label_text, content_text):
        t = doc.add_table(rows=2, cols=4)
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.LEFT
        # Row 0: label
        r0 = t.rows[0]
        c0 = r0.cells[0]
        c0.merge(r0.cells[3])
        _shade(c0, PRIMARY_SHADE)
        _set_cell_text(c0, label_text.upper(), bold=True, white=True)
        # Row 1: content
        r1 = t.rows[1]
        c1 = r1.cells[0]
        c1.merge(r1.cells[3])
        _set_cell_text(c1, content_text or "")
        doc.add_paragraph("")

    # ---- GOAL & OUTCOME banners (keep these)
    if goal_text:
        _add_banner_block("GOAL", goal_text)

    # ==== Outcome-level KPI table (separate) ====
    outcome_kpis = [k for k in st.session_state.get("kpis", []) if k.get("parent_level") == "Outcome"]

    if outcome_kpis:
        k = outcome_kpis[0]  # only one allowed

        tbl_outcome = doc.add_table(rows=1, cols=4)
        tbl_outcome.style = "Table Grid"
        tbl_outcome.alignment = WD_TABLE_ALIGNMENT.LEFT
        tbl_outcome.autofit = True

        # same headers as the main table to keep visual consistency
        hdr = tbl_outcome.rows[0]
        labels = ("Outcome", "KPI", "Means of Verification", "Key Assumptions")
        for i, lab in enumerate(labels):
            _set_cell_text(hdr.cells[i], lab, bold=True, white=True)
            _shade(hdr.cells[i], PRIMARY_SHADE)
        _repeat_header(hdr)

        # set stable column widths so Word won't squash them
        from docx.shared import Cm
        col_widths = (Cm(6.0), Cm(9.0), Cm(6.0), Cm(6.0))
        for i, w in enumerate(col_widths):
            for r in tbl_outcome.rows:
                r.cells[i].width = w
            tbl_outcome.columns[i].width = w

        r = tbl_outcome.add_row()
        for i, w in enumerate(col_widths):
            r.cells[i].width = w

        # Col 0: Outcome text
        _set_cell_text(r.cells[0], outcome_text or "")

        # Col 1: KPI details (same style as other KPIs)
        kcell = r.cells[1]; kcell.text = ""
        p = kcell.paragraphs[0]
        _add_run(p, f"Outcome KPI — {k.get('name','')}")
        p.add_run("\n")
        bp = (k.get("baseline","") or "").strip()
        if bp: _add_run(p, "Baseline: ", True); _add_run(p, bp); p.add_run("\n")
        tg = (k.get("target","") or "").strip()
        if tg: _add_run(p, "Target: ", True); _add_run(p, tg); p.add_run("\n")

        # Col 2: MoV
        _set_cell_text(r.cells[2], (k.get("mov") or "").strip() or "—")

        # Col 3: Assumptions at outcome level (leave blank or em dash)
        _set_cell_text(r.cells[3], "—")

        # a little whitespace before the next table
        doc.add_paragraph("")

    # ==== Main Outputs + KPIs table ====
    tbl = doc.add_table(rows=1, cols=4)
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl.autofit = True

    # match your main table header exactly
    hdr = tbl.rows[0]
    labels = ("Output", "KPI", "Means of Verification", "Key Assumptions")
    for i, lab in enumerate(labels):
        _set_cell_text(hdr.cells[i], lab, bold=True, white=True)
        _shade(hdr.cells[i], PRIMARY_SHADE)
    _repeat_header(hdr)

    # set widths for the main table too (tweak if you prefer your prior numbers)
    col_widths_main = (Cm(6.0), Cm(9.0), Cm(6.0), Cm(6.0))
    for i, w in enumerate(col_widths_main):
        for r in tbl.rows:
            r.cells[i].width = w
        tbl.columns[i].width = w

    # --- your existing Outputs & KPIs loop (unchanged) ---
    outputs = st.session_state.get("outputs", [])
    out_nums, kpi_nums = compute_numbers()

    def _sort_by_num(label):
        if not label: return (9999,)
        try: return tuple(int(x) for x in str(label).split("."))
        except: return (9999,)

    outputs = sorted(outputs, key=lambda o: _sort_by_num(out_nums.get(o["id"], "")))

    for out in outputs:
        out_num = out_nums.get(out["id"], "")
        # exclude outcome-level KPIs from this table
        kpis = [k for k in st.session_state.get("kpis", [])
                if k.get("parent_id") == out["id"] and k.get("parent_level") != "Outcome"]

        if not kpis:
            r = tbl.add_row()
            for i, w in enumerate(col_widths_main): r.cells[i].width = w
            _set_cell_text(r.cells[0], f"Output {out_num} — {out.get('name','')}")
            _set_cell_text(r.cells[1], "—")
            _set_cell_text(r.cells[2], "—")
            _set_cell_text(r.cells[3], out.get("assumptions","") or "—")
            continue

        first = len(tbl.rows)
        for k in kpis:
            r = tbl.add_row()
            for i, w in enumerate(col_widths_main): r.cells[i].width = w

            kcell = r.cells[1]; kcell.text = ""
            p = kcell.paragraphs[0]
            k_lab = kpi_nums.get(k["id"], "")
            _add_run(p, f"KPI ({k_lab}) — {k.get('name','')}")
            p.add_run("\n")

            bp = (k.get("baseline","") or "").strip()
            if bp: _add_run(p, "Baseline: ", True); _add_run(p, bp); p.add_run("\n")
            tg = (k.get("target","") or "").strip()
            if tg: _add_run(p, "Target: ", True); _add_run(p, tg); p.add_run("\n")
            _set_cell_text(r.cells[2], (k.get("mov") or "").strip() or "—")

        last = len(tbl.rows) - 1
        _set_cell_text(tbl.cell(first, 0), f"Output {out_num} — {out.get('name','')}")
        _set_cell_text(tbl.cell(first, 3), out.get("assumptions","") or "—")
        if last > first:
            tbl.cell(first, 0).merge(tbl.cell(last, 0))
            tbl.cell(first, 3).merge(tbl.cell(last, 3))

    # ==== WORKPLAN – Activities (table) ====
    from docx.shared import Cm
    _new_section("WORKPLAN")

    # Build tidy DF exactly as in the app
    df_wp = _workplan_df()  # columns: Activity | Output | Start | End

    # Header: same 4 columns as requested
    t_act = doc.add_table(rows=1, cols=4)
    t_act.style = "Table Grid"
    t_act.alignment = WD_TABLE_ALIGNMENT.LEFT

    hdr = t_act.rows[0]
    for i, lab in enumerate(("Output", "Activity", "Start", "End")):
        _set_cell_text(hdr.cells[i], lab, bold=True, white=True)
        _shade(hdr.cells[i], PRIMARY_SHADE)
    _repeat_header(hdr)

    from docx.shared import Cm
    # lock widths, similar proportions to logframe
    COLW = (Cm(6.0), Cm(9.5), Cm(3.0), Cm(3.0))
    for i, w in enumerate(COLW):
        for r in t_act.rows:
            r.cells[i].width = w
        t_act.columns[i].width = w

    # We want outputs numbered (1,2,...) in the same order as logframe
    out_nums, _ = compute_numbers()
    id_to_output = {o["id"]: (o.get("name") or "Output") for o in st.session_state.outputs}

    # Map output name -> number (using id_to_output & out_nums)
    # (df_wp has Output names already; we need to recover numbering by matching)
    name_to_num = {}
    for o in st.session_state.outputs:
        num = out_nums.get(o["id"], "")
        name = id_to_output.get(o["id"], "")
        if name:
            name_to_num[name] = num

    # Group activities by Output (ordered by logframe numbering), render with merged output cells
    if df_wp.empty:
        row = t_act.add_row()
        _set_cell_text(row.cells[0], "—")
        _set_cell_text(row.cells[1], "No scheduled activities")
        _set_cell_text(row.cells[2], "—")
        _set_cell_text(row.cells[3], "—")
    else:
        # 1) Stable within-group order: Start -> Activity
        df_wp = df_wp.sort_values(["Start", "Activity"], kind="stable").reset_index(drop=True)

        # 2) Deterministic group order: by Output number (from logframe)
        def _out_order(name):
            n = name_to_num.get(name, "")
            try:
                return int(n)
            except Exception:
                return 10 ** 9  # push unknowns to the end

        ordered_outputs = sorted(df_wp["Output"].unique().tolist(), key=_out_order)

        for out_name in ordered_outputs:
            sub = df_wp.loc[df_wp["Output"] == out_name]

            start_row = len(t_act.rows)
            for _, r in sub.iterrows():
                row = t_act.add_row()
                # keep widths on the new row
                for i, w in enumerate(COLW):
                    row.cells[i].width = w
                # fill activity row
                _set_cell_text(row.cells[1], str(r["Activity"]))
                _set_cell_text(row.cells[2], r["Start"].strftime("%d/%b/%Y"))
                _set_cell_text(row.cells[3], r["End"].strftime("%d/%b/%Y"))

            end_row = len(t_act.rows) - 1
            if end_row >= start_row:
                merged = t_act.cell(start_row, 0).merge(t_act.cell(end_row, 0))
                label = f"Output {name_to_num.get(out_name, '')} — {out_name}".strip(" —")
                _set_cell_text(merged, label)

    doc.add_paragraph("")  # a small gap after the table

    # ===== WORKPLAN – Gantt (LANDSCAPE section) =====
    try:
        gantt_buf = _gantt_png_buf()
        if gantt_buf:
            sec_land = _new_landscape_section("WORKPLAN – Gantt")
            from docx.shared import Cm
            width_cm = _content_width_cm(sec_land)
            doc.add_picture(gantt_buf, width=Cm(max(1.0, width_cm - 0.5)))  # 0.5 cm safety
    except Exception:
        pass

    # ===== BUDGET =====
    _ensure_portrait_section("BUDGET")

    # Three-column budget table: Budget item | Description | Total Cost (USD)
    bt = doc.add_table(rows=1, cols=3)
    bt.style = "Table Grid"
    bt.alignment = WD_TABLE_ALIGNMENT.LEFT

    bh = bt.rows[0]
    for i, lab in enumerate(("Budget item", "Description", "Total Cost (USD)")):
        _set_cell_text(bh.cells[i], lab, bold=True, white=True)
        _shade(bh.cells[i], PRIMARY_SHADE)
    _repeat_header(bh)

    from docx.shared import Cm
    for i, w in enumerate((Cm(6.0), Cm(11.0), Cm(4.0))):
        for r in bt.rows:
            r.cells[i].width = w
        bt.columns[i].width = w

    total_budget = 0.0
    for r in st.session_state.get("budget", []):
        row = bt.add_row()
        _set_cell_text(row.cells[0], r.get("item", ""))
        _set_cell_text(row.cells[1], r.get("description", ""))
        amt = float(r.get("total_usd") or 0.0)
        total_budget += amt
        _set_cell_text(row.cells[2], f"USD {amt:,.2f}")

    # Total line
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(f"Total budget: USD {total_budget:,.2f}")
    run.bold = True

    # USD→AED note under the budget table
    note = doc.add_paragraph()
    note_run = note.add_run(
        footer_note
    )
    note_run.italic = True

    # ===== DISBURSEMENT SCHEDULE =====
    _ensure_portrait_section("DISBURSEMENT SCHEDULE")

    from datetime import date as _date
    # Use saved disbursement rows (leave empty -> no rows)
    dsp_src = list(st.session_state.get("disbursement", []))

    # Sort consistently: by output label (for stability), then date, then deliverable text
    out_nums, _ = compute_numbers()
    id_to_output = {o["id"]: (o.get("name") or "Output") for o in st.session_state.outputs}

    def _out_label(oid):  # only for stable sort; NOT shown in table
        return f"Output {out_nums.get(oid, '')} — {id_to_output.get(oid, '(unassigned)')}".strip(" —")

    # Sort primarily by date (earliest first), then by Output number, then by KPI title
    out_nums, _ = compute_numbers()

    def _out_num_val(oid):
        n = out_nums.get(oid, "")
        try:
            return int(n)
        except Exception:
            return 10 ** 9  # push unknown/unnumbered outputs to the end

    dsp_rows = sorted(
        dsp_src,
        key=lambda d: (
            d.get("anticipated_date") or _date(2100, 1, 1),  # 1) date
            _out_num_val(d.get("output_id")),  # 2) output number (not shown)
            (d.get("kpi_name") or d.get("deliverable") or "").strip()  # 3) KPI title
        )
    )

    # Table: SAME style as other tables (no title row, no milestone id col)
    t = doc.add_table(rows=1, cols=3)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.LEFT

    # Header row (PRIMARY_SHADE + white)
    hdr = t.rows[0]
    headers = (
        "Anticipated deliverable date",
        "Deliverable",
        "Maximum Grant instalment payable on satisfaction of this deliverable (USD)",
    )
    for i, lab in enumerate(headers):
        _set_cell_text(hdr.cells[i], lab, bold=True, white=True)
        _shade(hdr.cells[i], PRIMARY_SHADE)
    _repeat_header(hdr)

    # Column widths
    for i, w in enumerate((Cm(5.0), Cm(12.0), Cm(5.0))):
        for r in t.rows:
            r.cells[i].width = w
        t.columns[i].width = w

    # Data rows (Deliverable = KPI title only; do NOT prepend Output)
    if dsp_rows:
        for d in dsp_rows:
            r = t.add_row()
            _set_cell_text(r.cells[0], d["anticipated_date"].strftime("%d/%b/%Y") if d.get("anticipated_date") else "")
            _set_cell_text(r.cells[1], (d.get("deliverable") or d.get("kpi_name") or ""))
            _set_cell_text(r.cells[2], f"{float(d.get('amount_usd') or 0.0):,.2f}")
    else:
        r = t.add_row()
        _set_cell_text(r.cells[0], "")
        _set_cell_text(r.cells[1], "No disbursements defined.")
        _set_cell_text(r.cells[2], "")

    # USD→AED note under the budget table
    note = doc.add_paragraph()
    note_run = note.add_run(
        footer_note
    )
    note_run.italic = True

    # ---- Save to buffer
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

def view_activity_readonly(a, label, id_to_output, id_to_kpi):
    out_name = id_to_output.get(a.get("output_id"), "(unassigned)")
    kpis_txt = ", ".join(id_to_kpi.get(kid, "") for kid in (a.get("kpi_ids") or [])) or "—"
    sd = fmt_dd_mmm_yyyy(a.get("start")) or "—"
    ed = fmt_dd_mmm_yyyy(a.get("end"))   or "—"

    body = (
        f"<div class='lf-activity-title'>Activity {escape(label)} — {escape(a.get('name',''))}</div>"
        f"<div class='lf-line'><b>Output:</b> {escape(out_name)}</div>"
        f"<div class='lf-line'><b>Owner:</b> {escape(a.get('owner','') or '—')}</div>"
        f"<div class='lf-line'><b>Start date:</b> {sd} &nbsp;&nbsp;•&nbsp;&nbsp; <b>End date:</b> {ed}</div>"
        f"<div class='lf-line'><b>Linked KPIs:</b> {escape(kpis_txt)}</div>"
    )

    return view_logframe_element(body, kind="activity")

def view_budget_item_card(row, id_to_output) -> str:
    """
    row: [OutputID, Item, Category, Unit, Qty, Unit Cost, Currency, Total]
    id_to_output: {output_id -> output name}
    """
    out_id, item, cat, unit, qty, uc, cur, tot, blid = _budget_unpack(row)
    out_name = id_to_output.get(out_id, "(unassigned)")
    # rows (label/value)
    lines = [
        ("Output", out_name),
        ("Item", item or "—"),
        ("Category", cat or "—"),
        ("Unit", unit or "—"),
        ("Qty", str(qty or 0)),
        ("Unit Cost", fmt_money(uc or 0)),
        ("Currency", cur or "—"),
        ("Total", fmt_money(tot or 0)),
    ]
    body = "".join(
        f"<div class='lf-line'><b>{escape(lbl)}:</b> {escape(val) if lbl not in ('Unit Cost','Total') else val}</div>"
        for (lbl, val) in lines
    )
    # reuse the blue style
    return f"<div class='lf-card lf-card--budget'>{body}</div>"

# ---------------- CSS for cards ----------------
def inject_logframe_css():
    st.markdown("""
    <style>
/* Base card */
.lf-card{
  margin: 14px 0 16px;
  padding: 14px 16px;
  border-radius: 12px;
  box-shadow: 0 1px 2px rgba(0,0,0,.03);
  position: relative;               /* required for the left accent bar */
}

/* ========= OUTPUT (green) ========= */
.lf-card--output{
  background: #E9F3E1;              /* pastel green fill */
  border: 1px solid #91A76A;        /* brand green border (#91A76A) */
}
.lf-card--output::before{
  content:"";
  position:absolute; top:0; left:0; bottom:0; width:8px;
  background:#6E8C4B;               /* darker green accent bar */
  border-top-left-radius:12px;
  border-bottom-left-radius:12px;
}

/* ========= KPI (orange) ========= */
.lf-card--kpi{
  background: #FFEBD6;          /* soft orange */
  border: 1px solid #F2B277;    /* warm orange border */
  border-radius: 12px;
  padding: 12px 16px;
  margin: 10px 0 12px;
  position: relative;
  box-shadow: 0 1px 2px rgba(0,0,0,.03);
}
.lf-card--kpi::before{
  content:"";
  position:absolute; top:0; left:0; bottom:0; width:6px;
  background:#DD7A1A;          /* stronger orange accent */
  border-top-left-radius:12px;
  border-bottom-left-radius:12px;
}

/* ========= Activity (yellow) ========= */
.lf-card--activity{
  background: #FFFBEA;          /* soft yellow */
  border: 1px solid #F6D58E;    /* warm yellow border */
  border-radius: 12px;
  padding: 12px 16px;
  margin: 10px 0 12px;
  position: relative;
  box-shadow: 0 1px 2px rgba(0,0,0,.03);
}
.lf-card--activity::before{
  content:"";
  position:absolute; top:0; left:0; bottom:0; width:6px;
  background:#F59E0B;          /* amber accent */
  border-top-left-radius:12px;
  border-bottom-left-radius:12px;
}

/* Headings & text bits */
.lf-out-header{ margin: 0; font-weight: 700; }
.lf-kpi-title{ font-weight: 600; margin-bottom: 6px; }

/* Assumptions list inside the green Output card */
.lf-card--output .lf-ass-heading{
  font-weight: 600;
  margin: 6px 0 4px;                 /* tighter heading spacing */
}

.lf-card--output .lf-ass-list{
  margin: 4px 0 0 1.15rem;           /* small left indent, no extra top gap */
  padding: 0;
  list-style: disc outside;
}

.lf-card--output .lf-ass-list li{
  margin: 0;                          /* remove extra gaps between items */
  padding: 0;
  font-style: italic;                 /* italics */
  font-size: 0.92rem;                 /* smaller text */
  line-height: 1.25;                  /* single/compact spacing */
}

/* Chips - Check if linked to payment*/
.chip {
  display:inline-block; padding: 2px 8px; border-radius: 999px;
  background:#eef2ff; font-size: .85rem; border:1px solid #e2e8ff;
}
.chip.green { background:#e6f9ee; border-color:#c6f0d8; color:#046c4e; }

/* Dot bullet in Output header */
.dot { font-size: 1.05rem; margin-right: .4rem; }

/* (optional) generic line styling used in KPI details */
.lf-line { margin: 2px 0; font-size: 0.95rem; color: #444; }

/* ========= BUDGET (blue) ========= */
.lf-card--budget{
  background: #E8F0FE;           /* light blue */
  border: 1px solid #90B4FE;
  border-radius: 12px;
  padding: 12px 16px;
  margin: 10px 0 12px;
  box-shadow: 0 1px 2px rgba(0,0,0,.03);
}
.lf-budget-total{ font-weight:700; margin-top:10px; }

.lf-card--budget .budget-table td:nth-child(1){
  max-width: 520px;              /* tune as needed */
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}   

/* --- Compact budget row (blue) --- */
.lf-budget-row{
  background:#E8F0FE;
  border:1px solid #90B4FE;
  border-radius:12px;
  padding:10px 12px;
  margin:8px 0;
  display:flex;
  align-items:center;
  gap:10px;
  box-shadow:0 1px 2px rgba(0,0,0,.03);
}
.lf-budget-cells{
  display:flex;
  gap:14px;
  flex-wrap:wrap;
  align-items:baseline;
  width:100%;
  font-size:.95rem;
  color:#1f2937;
}
.lf-budget-cell{
  white-space:nowrap;
}
.lf-budget-money{
  font-variant-numeric: tabular-nums;
  text-align:right;
}
.lf-budget-chip{
  background:#eef2ff;
  border:1px solid #dbe2ff;
  border-radius:999px;
  padding:2px 8px;
  display:inline-block;
}
.lf-subtotal{ font-weight:700; margin:6px 0 12px; text-align:right; }
.lf-grandtotal{ font-weight:800; margin-top:12px; text-align:right; }

/* Make the Streamlit-native header logo larger */
header [data-testid="stLogo"] img {
  height: 80px;      /* change this */
  width: auto;
}
header [data-testid="stLogo"] {
  margin-left: 8px;  /* optional: shift slightly from edge */
  padding-top: 6px;  /* optional */
  padding-bottom: 6px;
}

    </style>
    """, unsafe_allow_html=True)

# ---------------- Tabs ----------------
tabs = st.tabs([
    "📘 Instructions",
    "🪪 Identification",
    "🧱 Logframe",
    "🗂️ Workplan",
    "💵 Budget",
    "💸 Disbursement Schedule",
    "📤 Export"
])

# ===== TAB 1: Instructions =====
tabs[0].markdown(
    """
# 📝 Welcome to the Falcon Awards Application Portal

Please complete each section of your application:

1. **Identification** – Fill project & contacts.
2. **Logframe** – Add **Goal**, **Outcome**, **Outputs**, then **KPIs**.
3. **Workplan** – Add **Activities** linked to Outputs/KPIs with dates.
4. **Budget** – Add **Budget lines** linked to Outputs.
5. **Export** – Use **Generate Work Document** to produce a .docx that includes the logframe, workplan (table & Gantt) and budget.

### Definitions
- **Goal**: The long-term vision (impact).
- **Outcome**: The specific change expected from the project.
- **Output**: Tangible products/services delivered by the project.
- **KPI**: Quantifiable metric to judge performance (with baseline, target, dates, MoV).
- **Activity**: Tasks that produce Outputs (scheduled in the **Workplan**).
- **Assumptions**: External conditions necessary for success.
- **Means of Verification (MoV)**: Where/how the KPI will be measured.
- **Payment-linked indicator**: KPI that triggers financing when achieved (optional).
- **Budget line**: A costed item (category, unit, quantity, unit cost).

Once done, export your application as an Excel file.

"""
)
# **Contact us**: Anderson E. Stanciole | astanciole@glideae.org

# --- Resume from Excel ---
uploaded_file = tabs[0].file_uploader("Resume Previous Submission (Excel)", type="xlsx")
if uploaded_file is not None:
    try:
        # Build a stable signature for the uploaded file
        file_bytes = uploaded_file.getvalue()
        file_sig = hashlib.md5(file_bytes).hexdigest()

        # Only (re)load if this is a new file or changed content
        if st.session_state.get("_resume_file_sig") != file_sig:
            xls = pd.ExcelFile(BytesIO(file_bytes))

            # ---- RESET state containers
            st.session_state.impacts = []
            st.session_state.outcomes = []
            st.session_state.outputs = []
            st.session_state.kpis = []

            # (optional) clear edit flags so they won't point to old IDs
            for _f in ("edit_goal", "edit_outcome", "edit_output", "edit_kpi"):
                st.session_state[_f] = None

            # ---- Read Summary with explicit IDs ----
            summary_df = pd.read_excel(xls, sheet_name="Summary")
            summary_df.columns = [str(c).strip() for c in summary_df.columns]

            st.session_state.impacts = []
            st.session_state.outcomes = []
            st.session_state.outputs = []

            for _, row in summary_df.iterrows():
                lvl = _s(row.get("RowLevel", ""))
                _id = _s(row.get("ID")) or generate_id()
                pid = _s(row.get("ParentID")) or None
                text = _s(row.get("Text / Title"))
                ass = _s(row.get("Assumptions"))

                if lvl.lower() == "goal":
                    st.session_state.impacts.append({"id": _id, "level": "Goal", "name": text})

                elif lvl.lower() == "outcome":
                    st.session_state.outcomes.append({"id": _id, "level": "Outcome", "name": text, "parent_id": pid})

                elif lvl.lower() == "output":
                    st.session_state.outputs.append({
                        "id": _id,
                        "level": "Output",
                        "name": strip_label_prefix(text, "Output") or "Output",
                        "parent_id": pid,
                        "assumptions": ass
                    })

            # ---- KPI Matrix (ID-based) ----
            if "KPI Matrix" in xls.sheet_names:
                kdf = pd.read_excel(xls, sheet_name="KPI Matrix")
                kdf.columns = [str(c).strip() for c in kdf.columns]

                st.session_state.kpis = []  # reset before loading

                for _, r in kdf.iterrows():
                    kid = _s(r.get("KPIID")) or generate_id()
                    plev = _s(r.get("Parent Level") or "Output")
                    pid = _s(r.get("ParentID")) or None

                    st.session_state.kpis.append({
                        "id": kid,
                        "level": "KPI",
                        "name": _s(r.get("KPI")),
                        "parent_id": pid,
                        "parent_level": "Outcome" if plev.lower() == "outcome" else "Output",
                        "baseline": _s(r.get("Baseline")),
                        "target": _s(r.get("Target")),
                        "start_date": parse_date_like(r.get("Start Date")),
                        "end_date": parse_date_like(r.get("End Date")),
                        "linked_payment": _s(r.get("Linked to Payment")).lower() in ("yes", "y", "true", "1"),
                        "mov": _s(r.get("Means of Verification")),
                    })

            # ---- Workplan (supports both "rich" and "simple" exports)
            if "Workplan" in xls.sheet_names:
                wdf = pd.read_excel(xls, sheet_name="Workplan")
                # normalize headers
                wdf.columns = [str(c).strip() for c in wdf.columns]

                def _split_csv(s):
                    return [t.strip() for t in (_s(s).split(",") if _s(s) else []) if t.strip()]

                # lookups
                outputs_by_name = {(o.get("name") or "").strip(): o["id"] for o in st.session_state.outputs}
                kpis_by_name = {(k.get("name") or "").strip(): k["id"] for k in st.session_state.kpis}
                kpi_id_set = set(kpis_by_name.values())

                # detect "rich" vs "simple" format
                rich = {"Activity ID", "Activity #", "OutputID", "Output", "Activity", "Owner", "Start", "End",
                        "Linked KPI IDs", "Linked KPIs"}.issubset(set(wdf.columns))

                st.session_state.workplan = []  # reset before loading

                if rich:
                    output_ids = {o["id"] for o in st.session_state.outputs}
                    kpi_ids = {k["id"] for k in st.session_state.kpis}


                    def _csv_ids(s):
                        vals = [t.strip() for t in (_s(s).split(",") if _s(s) else []) if t.strip()]
                        return vals


                    for _, row in wdf.iterrows():
                        act_id = _s(row.get("Activity ID")) or generate_id()
                        out_id = _s(row.get("OutputID"))
                        if not out_id:
                            out_name = _s(row.get("Output"))
                            out_id = next((o["id"] for o in st.session_state.outputs if
                                           (o.get("name") or "").strip() == out_name), None)

                        linked_kpi_ids = [kid for kid in _csv_ids(row.get("Linked KPI IDs")) if kid in kpi_ids]
                        dep_ids = _csv_ids(row.get("Dependencies"))

                        st.session_state.workplan.append({
                            "id": act_id,
                            "output_id": out_id if out_id in output_ids else None,
                            "name": _s(row.get("Activity")),
                            "owner": _s(row.get("Owner")),
                            "start": parse_date_like(row.get("Start")),
                            "end": parse_date_like(row.get("End"))
                        })

                else:
                    # SIMPLE legacy format: Activity | Owner | Start Date | End Date | Milestone
                    # No Output/KPI info in this shape; if there is exactly one Output, attach to it.
                    only_output_id = st.session_state.outputs[0]["id"] if len(st.session_state.outputs) == 1 else None
                    for _, row in wdf.iterrows():
                        st.session_state.workplan.append({
                            "id": generate_id(),
                            "output_id": only_output_id,  # None if multiple outputs; user can reassign in UI
                            "name": _s(row.get("Activity")),
                            "owner": _s(row.get("Owner")),
                            "start": parse_date_like(row.get("Start Date")),
                            "end": parse_date_like(row.get("End Date")),
                            "kpi_ids": [],
                        })

            # ---- Budget import (optional sheet)
            if "Budget" in xls.sheet_names:
                bdf = pd.read_excel(xls, sheet_name="Budget")
                bdf.columns = [str(c).strip() for c in bdf.columns]
                st.session_state.budget = []
                for _, r in bdf.iterrows():
                    item = _s(r.get("Budget item"))
                    desc = _s(r.get("Description"))
                    tot = float(r.get("Total Cost (USD)") or 0.0)
                    if item or desc or tot:
                        st.session_state.budget.append({"item": item, "description": desc, "total_usd": tot})

            # ---- Disbursement Schedule import (optional sheet)
            if "Disbursement Schedule" in xls.sheet_names:
                ddf = pd.read_excel(xls, sheet_name="Disbursement Schedule")
                ddf.columns = [str(c).strip() for c in ddf.columns]

                st.session_state.disbursement = []

                k_by_id = {k["id"]: k for k in st.session_state.kpis}

                for _, row in ddf.iterrows():
                    kid = _s(row.get("KPIID"))
                    k = k_by_id.get(kid)

                    st.session_state.disbursement.append({
                        "kpi_id": kid,
                        "output_id": (k.get("parent_id") if k else None),
                        "kpi_name": (k.get("name") if k else ""),
                        "anticipated_date": parse_date_like(row.get("Anticipated deliverable date")),
                        "deliverable": _s(row.get("Deliverable")),
                        "amount_usd": float(row.get(
                            "Maximum Grant instalment payable on satisfaction of this deliverable (USD)") or 0.0),
                    })

            # --- Import Identification sheet (if present) and update ID page state ---
            if "Identification" in xls.sheet_names:
                id_df = pd.read_excel(xls, sheet_name="Identification")
                try:
                    kv = {str(r["Field"]).strip(): str(r["Value"]) if not pd.isna(r["Value"]) else ""
                          for _, r in id_df.iterrows()}
                except Exception:
                    kv = {}

                def _g(field):  # helper to get a string from the kv map
                    return (kv.get(field, "") or "").strip()

                id_info = st.session_state.get("id_info", {}) or {}
                id_info.update({
                    "title": _g("Project title"),
                    "pi_name": _g("Principal Investigator (PI) name"),
                    "pi_email": _g("PI email"),
                    "institution": _g("Institution / Organization"),
                    "start_date": parse_date_like(kv.get("Project start date", "")) or id_info.get("start_date"),
                    "end_date": parse_date_like(kv.get("Project end date", "")) or id_info.get("end_date"),
                    "contact_name": _g("Contact person (optional)"),
                    "contact_email": _g("Contact email"),
                    "contact_phone": _g("Contact phone"),
                })
                st.session_state.id_info = id_info

                # also prime the live widget keys so the inputs show the imported values immediately
                st.session_state["id_title"]        = id_info["title"]
                st.session_state["id_pi_name"]      = id_info["pi_name"]
                st.session_state["id_pi_email"]     = id_info["pi_email"]
                st.session_state["id_institution"]  = id_info["institution"]
                st.session_state["id_start_date"]   = id_info["start_date"]
                st.session_state["id_end_date"]     = id_info["end_date"]
                st.session_state["id_contact_name"] = id_info["contact_name"]
                st.session_state["id_contact_email"]= id_info["contact_email"]
                st.session_state["id_contact_phone"]= id_info["contact_phone"]

            # Remember we loaded this file content; prevents re-import on button clicks
            st.session_state["_resume_file_sig"] = file_sig
            tabs[0].success("✅ Previous submission loaded into session.")
            st.rerun()

        # else: same file uploaded again → skip re-import so edit/delete works
    except Exception as e:
        tabs[0].error(f"Could not parse uploaded Excel: {e}")

# ===== TAB 2: Identification =====
with tabs[1]:
    st.header("🪪 Project Identification")

    # defaults
    if "id_info" not in st.session_state:
        st.session_state.id_info = {
            "title": "", "pi_name": "", "pi_email": "", "institution": "",
            "start_date": None, "end_date": None,
            "contact_name": "", "contact_email": "", "contact_phone": ""
        }

    # ensure widget keys exist (so inputs are persistent & can be set by the importer)
    for k, v in [
        ("id_title", st.session_state.id_info["title"]),
        ("id_pi_name", st.session_state.id_info["pi_name"]),
        ("id_pi_email", st.session_state.id_info["pi_email"]),
        ("id_institution", st.session_state.id_info["institution"]),
        ("id_start_date", st.session_state.id_info["start_date"]),
        ("id_end_date", st.session_state.id_info["end_date"]),
        ("id_contact_name", st.session_state.id_info["contact_name"]),
        ("id_contact_email", st.session_state.id_info["contact_email"]),
        ("id_contact_phone", st.session_state.id_info["contact_phone"]),
    ]:
        if k not in st.session_state:
            st.session_state[k] = v

    c1, c2 = st.columns(2)
    with c1:
        st.session_state.id_info["title"] = st.text_input("Project title*", key="id_title")
        st.session_state.id_info["pi_name"] = st.text_input("Principal Investigator (PI) name*", key="id_pi_name")
        st.session_state.id_info["pi_email"] = st.text_input("PI email*", key="id_pi_email")
        st.session_state.id_info["institution"] = st.text_input("Institution / Organization*", key="id_institution")
    with c2:
        sd = st.date_input("Project start date*", key="id_start_date")
        if sd:
            st.caption(f"Selected: {fmt_dd_mmm_yyyy(sd)}")  # show DD/Mon/YYYY preview

        ed = st.date_input("Project end date*", key="id_end_date")
        if ed:
            st.caption(f"Selected: {fmt_dd_mmm_yyyy(ed)}")  # show DD/Mon/YYYY preview

        # keep canonical values in id_info for export/validation
        st.session_state.id_info["start_date"] = sd
        st.session_state.id_info["end_date"] = ed
        with st.expander("More contact details (optional)"):
            st.session_state.id_info["contact_name"] = st.text_input("Contact person (if different from PI)", key="id_contact_name")
            st.session_state.id_info["contact_email"] = st.text_input("Contact email", key="id_contact_email")
            st.session_state.id_info["contact_phone"] = st.text_input("Contact phone", key="id_contact_phone")

    # inline validation (no button)
    errs = []
    ii = st.session_state.id_info
    if not ii["title"].strip():       errs.append("Project title is required.")
    if not ii["pi_name"].strip():     errs.append("PI name is required.")
    if not ii["institution"].strip(): errs.append("Institution is required.")
    if not ii["pi_email"].strip() or "@" not in ii["pi_email"]: errs.append("Valid PI email is required.")
    if ii["start_date"] and ii["end_date"] and ii["start_date"] > ii["end_date"]:
        errs.append("Project start date must be on or before the end date.")

    # show errors only after the user has typed something in any required field
    touched = any([
        st.session_state["id_title"].strip(),
        st.session_state["id_pi_name"].strip(),
        st.session_state["id_pi_email"].strip(),
        st.session_state["id_institution"].strip(),
    ])
    if touched:
        for e in errs:
            st.error(e)

    # --- Read-only summary (live)
    # Budget total (computed from detailed budget)
    def _sum_budget():
        return sum(float(r.get("total_usd") or 0.0) for r in st.session_state.get("budget", []))

    budget_total = _sum_budget()

    # Live counts
    outputs_count = len(st.session_state.get("outputs", []))
    # activities_count = len(st.session_state.get("activities", []))
    kpis_count = len(st.session_state.get("kpis", []))

    st.markdown("### Summary")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Total funding requested (from Budget):**")
        st.markdown(f"**USD {budget_total:,.2f}**")
    with c2:
        st.markdown("**Logframe indicators:**")
        st.markdown(
            f"""
            <div style="display:flex; gap:10px; align-items:center;">
              <div style="background:#eef2ff;border:1px solid #dbe2ff;border-radius:999px;padding:6px 10px;">Outputs: <b>{outputs_count}</b></div>
              <div style="background:#f7f7f9;border:1px solid #e6e6e6;border-radius:999px;padding:6px 10px;">KPIs: <b>{kpis_count}</b></div>
            </div>
            """,
            unsafe_allow_html=True
        )

    # Cross-check note area (non-blocking warnings placeholder)
    warnings = []
    # (Optional) You can later add KPI due-date vs project dates checks here and append messages to `warnings`.
    if warnings:
        for w in warnings:
            st.warning(w)


# ===== TAB 3: Logframe =====
tabs[2].header("📊 Build Your Logframe")
inject_logframe_css()
# --- numbering for labels shown in UI and preview ---
out_nums, kpi_nums = compute_numbers()   # outputs -> 'n', KPIs -> 'n.p'

# --- Add forms ---
with tabs[2].expander("➕ Add Goal"):
    if len(st.session_state.impacts) >= 1:
        st.info("Only one Goal is allowed. Edit the existing Goal in the preview below.")
    else:
        with st.form("goal_form"):
            goal_text = st.text_area("Goal (single, high-level statement)")
            if st.form_submit_button("Add Goal") and goal_text.strip():
                st.session_state.impacts.append({"id": generate_id(), "level": "Goal", "name": goal_text.strip()})

with tabs[2].expander("➕ Add Outcome"):
    if not st.session_state.impacts:
        st.warning("Add the Goal first.")
    elif len(st.session_state.outcomes) >= 1:
        st.info("Only one Outcome is allowed. Edit the existing Outcome in the preview below.")
    else:
        with st.form("outcome_form"):
            outcome_text = st.text_area("Outcome (statement)")
            # since there is only one goal, no need to pick it; link to the single goal
            linked_goal_id = st.session_state.impacts[0]["id"]
            if st.form_submit_button("Add Outcome") and outcome_text.strip():
                st.session_state.outcomes.append(
                    {"id": generate_id(), "level": "Outcome", "name": outcome_text.strip(), "parent_id": linked_goal_id}
                )

# --- Add ONE Outcome-level KPI ---
with tabs[2].expander("➕ Add Outcome KPI (one per outcome)"):
    if not st.session_state.outcomes:
        tabs[2].warning("Add the Outcome first.")
    else:
        outcome = st.session_state.outcomes[0]   # single-outcome design
        existing_outcome_kpis = [
            k for k in st.session_state.kpis
            if k.get("parent_level") == "Outcome" and k.get("parent_id") == outcome["id"]
        ]

        if existing_outcome_kpis:
            st.info("This outcome already has a KPI. Edit it below in the preview.")
        else:
            with st.form("kpi_outcome_form"):
                kpi_text = st.text_area("Outcome KPI*")
                baseline = st.text_input("Baseline")
                target   = st.text_input("Target")
                payment_linked = st.checkbox("Linked to Payment (optional)")
                mov = st.text_area("Means of Verification")

                if st.form_submit_button("Add Outcome KPI") and kpi_text.strip():
                    st.session_state.kpis.append({
                        "id": generate_id(),
                        "level": "KPI",
                        "name": strip_label_prefix(kpi_text.strip(), "KPI"),
                        "parent_id": outcome["id"],
                        "parent_level": "Outcome",   # <<< key difference
                        "baseline": baseline.strip(),
                        "target": target.strip(),
                        "linked_payment": bool(payment_linked),
                        "mov": mov.strip(),
                    })
                    st.rerun()

with tabs[2].expander("➕ Add Output"):
    if not st.session_state.outcomes:
        tabs[2].warning("Add the Outcome first.")
    else:
        with st.form("output_form"):
            output_title = st.text_input("Output title (e.g., 'Output 1')")
            output_assumptions = st.text_area("Key Assumptions (optional)")
            if st.form_submit_button("Add Output") and output_title.strip():
                linked_outcome_id = st.session_state.outcomes[0]["id"]
                st.session_state.outputs.append(
                    {
                        "id": generate_id(),
                        "level": "Output",
                        "name": output_title.strip(),
                        "parent_id": linked_outcome_id,
                        "assumptions": output_assumptions.strip(),
                    }
                )

with tabs[2].expander("➕ Add KPI"):
    if not st.session_state.outputs:
        tabs[2].warning("Add an Output first.")
    else:
        with st.form("kpi_form"):
            parent = st.selectbox(
                "Parent Output",
                st.session_state.outputs,
                format_func=lambda o: f"Output {out_nums.get(o['id'],'?')} — {o.get('name','Output')}"
            )
            kpi_text = st.text_area("KPI*")
            baseline = st.text_input("Baseline")
            target   = st.text_input("Target")
            payment_linked = st.checkbox("Linked to Payment (optional)")
            mov = st.text_area("Means of Verification")

            if st.form_submit_button("Add KPI") and kpi_text.strip():
                st.session_state.kpis.append({
                    "id": generate_id(),
                    "level": "KPI",
                    "name": strip_label_prefix(kpi_text.strip(), "KPI"),
                    "parent_id": parent["id"],
                    "parent_level": "Output",   # <-- fixed
                    "baseline": baseline.strip(),
                    "target": target.strip(),
                    "linked_payment": bool(payment_linked),
                    "mov": mov.strip(),
                })
                st.rerun()

# ---- View helpers (card layout, compact-aware) ----
def view_goal(g):
    return f"##### 🟦 **Goal:** {g.get('name','')}"

def view_outcome(o):
    return f"##### 🟪 **Outcome:** {o.get('name','')}"

def view_output(out):
    num   = out_nums.get(out["id"], "?")
    title = out.get("name", "Output")

    # header line (keeps your green card style via view_logframe_element)
    header_html = (
        f"<div class='lf-out-header'><strong>Output {num}:</strong> {title}</div>"
    )

    # assumptions -> bullet list (one bullet per line the user typed)
    ass = (out.get("assumptions") or "").strip()
    ass_html = ""
    if ass:
        # strip any leading "-" or "•", ignore empty lines
        items = [
            re.sub(r"^[\-\u2022]\s*", "", ln).strip()
            for ln in ass.splitlines()
            if ln.strip()
        ]
        if items:
            lis = "".join(f"<li>{html.escape(x)}</li>" for x in items)
            ass_html = (
                "<div class='lf-ass'>"
                "<div class='lf-ass-heading'> <b> Key Assumptions </b> </div>"
                f"<ul class='lf-ass-list'>{lis}</ul>"
                "</div>"
            )

    # wrap in the green card
    return view_logframe_element(header_html + ass_html, kind="output")

def view_output_header(out):
    num   = out_nums.get(out["id"], "?")
    title = escape(out.get("name", "Output"))
    header_html = f"<div class='lf-out-header'><strong>Output {num}:</strong> {title}</div>"
    return view_logframe_element(header_html, kind="output")

def view_kpi(k):
    # decide label prefix based on level
    is_outcome = (k.get("parent_level") == "Outcome")
    num  = kpi_nums.get(k["id"], "?") if not is_outcome else ""  # no numbering for outcome-level
    name = k.get("name", "")

    bp  = (k.get("baseline") or "").strip()
    tg  = (k.get("target") or "").strip()
    mov = (k.get("mov") or "").strip()

    chip = (
        "<span class='chip green'>Payment-linked</span>"
        if k.get("linked_payment")
        else "<span class='chip'>Not payment-linked</span>"
    )

    # Title
    if is_outcome:
        header = f"<div class='lf-kpi-title'>Outcome KPI: {name}</div>"
    else:
        header = f"<div class='lf-kpi-title'>KPI {num}: {name}</div>"

    lines = []
    if bp:
        lines.append(f"<div class='lf-line'><b>Baseline:</b> {bp}</div>")
    if tg:
        lines.append(f"<div class='lf-line'><b>Target:</b> {tg}</div>")
    if mov:
        lines.append(f"<div class='lf-line'><b>Means of Verification:</b> {mov}</div>")

    lines.append(f"<div class='lf-line'>{chip}</div>")
    inner = header + "".join(lines)
    return view_logframe_element(inner, kind="kpi")

def view_activity(a: dict, act_label: str, id_to_output: dict, id_to_kpi: dict) -> str:
    """
    Render one activity as a card.
    - a: activity dict from st.session_state.workplan
    - act_label: precomputed label like "1.2"
    - id_to_output: {output_id -> output name}
    - id_to_kpi:    {kpi_id -> kpi name}
    """
    out_name = id_to_output.get(a.get("output_id"), "(unassigned)")
    title    = f"Activity {escape(act_label)} — {escape(a.get('name',''))}"
    owner    = escape(a.get("owner","") or "—")
    sd       = fmt_dd_mmm_yyyy(a.get("start")) or "—"
    ed       = fmt_dd_mmm_yyyy(a.get("end"))   or "—"
    kpis_txt = ", ".join(escape(id_to_kpi.get(kid,"")) for kid in (a.get("kpi_ids") or [])) or "—"

    title_html = f"<div class='lf-activity-title'>{title}</div>"
    rows = [
        ("Output", out_name),
        ("Owner", owner),
        ("Start date", sd),
        ("End date", ed),
        ("Linked KPIs", kpis_txt),
    ]

    # Build rows (simple label/value pairs). You can skip empty ones here if desired.
    body = "".join(
        f"<div class='lf-line'><b>{escape(lbl)}:</b> {val}</div>"
        for (lbl, val) in rows
        if (val is not None and str(val).strip() != "")
    )

    # Use your generic card wrapper (orange indicators style)
    return view_logframe_element(title_html + body, kind="activity")

def make_ext_id(kind: str, text: str) -> str:
    """
    Deterministic short id for external linking across exports/imports.
    kind: 'output' | 'kpi'
    text: any stable text (e.g., output name; for KPI: 'OutputName|KPI text')
    """
    base = f"{kind}:{(text or '').strip().lower()}"
    return hashlib.sha1(base.encode("utf-8")).hexdigest()[:10]

# --- Inline preview with Edit / Delete buttons (refactored, card layout) ---
with tabs[2]:
    st.markdown("---")
    st.subheader("Current Logframe (preview) — click ✏️ to edit, 🗑️ to delete")

    for g in st.session_state.get("impacts", []):
        render_editable_item(
            item=g,
            list_name="impacts",
            edit_flag_key="edit_goal",
            view_md_func=view_goal,
            default_label="Goal",
            on_delete=lambda _id=g["id"]: (delete_cascade(goal_id=_id), st.rerun()),
            key_prefix="lf"
        )

        outcomes_here = [o for o in st.session_state.get("outcomes", []) if o.get("parent_id") == g["id"]]
        for oc in outcomes_here:
            render_editable_item(
                item=oc,
                list_name="outcomes",
                edit_flag_key="edit_outcome",
                view_md_func=view_outcome,
                default_label="Outcome",
                on_delete=lambda _id=oc["id"]: (delete_cascade(outcome_id=_id), st.rerun()),
                key_prefix="lf"
            )

            # --- Outcome-level KPI preview (edit/delete same as output KPI) ---
            outcome_kpis = [k for k in st.session_state.get("kpis", [])
                            if k.get("parent_id") == oc["id"] and k.get("parent_level") == "Outcome"]

            for k in outcome_kpis:
                render_editable_item(
                    item=k, list_name="kpis", edit_flag_key="edit_kpi",
                    view_md_func=lambda kk: view_kpi(kk),  # reuse the same renderer
                    fields=[
                        ("name", st.text_area, "KPI"),
                        ("baseline", st.text_input, "Baseline"),
                        ("target", st.text_input, "Target"),
                        ("linked_payment",
                         lambda label, value, key: st.checkbox(label, value=bool(value), key=key),
                         "Linked to Payment"),
                        ("mov", st.text_area, "Means of Verification"),
                    ],
                    on_delete=lambda _id=k["id"]: (
                        setattr(st.session_state, "kpis", [x for x in st.session_state.kpis if x["id"] != _id]),
                        st.rerun()
                    ),
                    key_prefix="lf"
                )

            outs_here = [o for o in st.session_state.get("outputs", []) if o.get("parent_id") == oc["id"]]
            for out in outs_here:
                with st.container():  # now this container lives inside the Logframe tab
                    render_editable_item(
                        item=out,
                        list_name="outputs",
                        edit_flag_key="edit_output",
                        view_md_func=view_output,
                        fields=[
                            ("name", st.text_input, "Output title"),
                            ("assumptions", st.text_area, "Key Assumptions"),
                        ],
                        on_delete=lambda _id=out["id"]: (delete_cascade(output_id=_id), st.rerun()),
                        key_prefix="lf"
                    )

                    k_children = [k for k in st.session_state.get("kpis", [])
                                  if (k.get("parent_id") == out["id"])]  # parent_level check optional if you dropped outcome-level KPIs
                    for k in k_children:
                        render_editable_item(
                            item=k, list_name="kpis", edit_flag_key="edit_kpi",
                            view_md_func=view_kpi,
                            fields=[
                                ("parent_id",
                                 lambda label, value, key: select_output_id(label, value, key),
                                 "Linked Output"),                                ("name", st.text_area, "KPI"),
                                ("baseline", st.text_input, "Baseline"),
                                ("target", st.text_input, "Target"),
                                ("linked_payment",
                                 lambda label, value, key: st.checkbox(label, value=bool(value), key=key),
                                 "Linked to Payment"),
                                ("mov", st.text_area, "Means of Verification"),
                            ],
                            on_delete=lambda _id=k["id"]: (
                                setattr(st.session_state, "kpis", [x for x in st.session_state.kpis if x["id"] != _id]),
                                st.rerun()
                            ),
                            key_prefix="lf"
                        )

# ===== TAB 4: Workplan =====
with tabs[3]:
    st.header("📆 Workplan")
    with st.expander("➕ Add Activity"):
        with st.form("workplan_form_v2"):
            # Required: link to Output
            output_parent = st.selectbox(
                "Linked Output*",
                st.session_state.outputs,
                format_func=lambda x: x.get("name") or "Output"
            )

            # Optional: link to KPI(s) under that Output
            output_id = output_parent["id"] if output_parent else None
            kpis_for_output = [k for k in st.session_state.kpis if k.get("parent_id") == output_id]
            kpi_links = st.multiselect(
                "Linked KPI(s) (optional)",
                kpis_for_output,
                format_func=lambda k: f"{k.get('name','')}"
            )

            name = st.text_input("Activity*")
            owner = st.text_input("Responsible person/institution*")
            c1, c2 = st.columns(2)
            with c1:
                start = st.date_input("Start date*")
            with c2:
                end = st.date_input("End date*")

            submitted = st.form_submit_button("Add to Workplan")
            if submitted and name.strip() and owner.strip() and output_parent and start and end and start <= end:
                st.session_state.workplan.append({
                    "id": generate_id(),
                    "output_id": output_id,
                    "name": name.strip(),
                    "kpi_ids": [k["id"] for k in kpi_links],
                    "owner": owner.strip(),
                    "start": start,
                    "end": end
                })
                st.rerun()
            elif submitted:
                st.warning("Please fill required fields (Output, Activity, Owner, Start≤End).")

    # ------- Gantt (web) -------
    df_g = _workplan_df()
    with st.expander("📈 Gantt chart (Workplan)", expanded=bool(len(df_g))):
        if df_g.empty:
            st.info("Add activities with start & end dates to see the Gantt.")
        else:
            height = min(1.2 + 0.35 * len(df_g), 12)
            fig, ax = plt.subplots(figsize=(11, height))
            _draw_gantt(ax, df_g, show_today=False)
            fig.subplots_adjust(left=0.26, right=0.98, top=0.96, bottom=0.30)
            st.pyplot(fig, clear_figure=True, use_container_width=True)

    # --- Card view (optional: put this below or instead of the table) ---
    out_nums, kpi_nums, act_nums = compute_numbers(include_activities=True)
    id_to_output = {o["id"]: (o.get("name") or "Output") for o in st.session_state.outputs}
    id_to_kpi    = {k["id"]: (k.get("name") or "")       for k in st.session_state.kpis}

    for oc in st.session_state.outcomes:
        outs_here = [o for o in st.session_state.outputs if o.get("parent_id") == oc["id"]]
        for out in outs_here:
            # green Output header card
            st.markdown(view_output_header(out), unsafe_allow_html=True)

            # orange Activity cards (with edit/delete)
            acts_here = [a for a in st.session_state.workplan if a.get("output_id") == out["id"]]
            for a in acts_here:
                label = act_nums.get(a["id"], "?")
                # Edit mode?
                if st.session_state.get("edit_activity") == a["id"]:
                    e1, e2, e3 = st.columns([0.90, 0.05, 0.05])
                    with e1:
                        new_output_id = select_output_id("Linked Output", a.get("output_id"), f"a_out_{a['id']}")
                        new_name = st.text_input("Activity", value=a.get("name", ""), key=f"a_name_{a['id']}")
                        new_owner = st.text_input("Owner", value=a.get("owner", ""), key=f"a_owner_{a['id']}")
                        cA, cB = st.columns(2)
                        with cA:
                            new_start = st.date_input("Start date", value=a.get("start"), key=f"a_start_{a['id']}")
                        with cB:
                            new_end = st.date_input("End date", value=a.get("end"), key=f"a_end_{a['id']}")
                    if e2.button("💾", key=f"a_save_{a['id']}"):
                        idx = _find_by_id(st.session_state.workplan, a["id"])
                        if idx is not None:
                            st.session_state.workplan[idx].update({
                                "output_id": new_output_id,
                                "name": new_name.strip(),
                                "owner": new_owner.strip(),
                                "start": new_start,
                                "end": new_end,
                            })
                        st.session_state["edit_activity"] = None
                        st.rerun()
                    if e3.button("✖️", key=f"a_cancel_{a['id']}"):
                        st.session_state["edit_activity"] = None
                        st.rerun()
                else:
                    v1, v2, v3 = st.columns([0.90, 0.05, 0.05])
                    v1.markdown(
                        view_activity_readonly(a, label, id_to_output, id_to_kpi),
                        unsafe_allow_html=True
                    )
                    if v2.button("✏️", key=f"a_edit_{a['id']}"):
                        st.session_state["edit_activity"] = a["id"]
                        st.rerun()
                    if v3.button("🗑️", key=f"a_del_{a['id']}"):
                        st.session_state.workplan = [x for x in st.session_state.workplan if x["id"] != a["id"]]
                        st.rerun()

# ===== TAB 5: Budget =====
with tabs[4]:
    st.header("💵 Define Budget\nEnter amounts in USD")

    # ---------- Add new budget item ----------
    with st.expander("➕ Add Budget Item"):
        with st.form("budget_form_simple"):
            item = st.text_input("Budget item*")
            desc = st.text_area("Description*", placeholder="Brief description…")
            total = st.number_input("Total Cost (USD)*", min_value=0.0, value=0.0,
                                    step=500.0, format="%.2f")

            submitted = st.form_submit_button("Add")
            if submitted:
                if not item.strip():
                    st.warning("Budget item is required.")
                elif not desc.strip():
                    st.warning("Description is required.")
                elif total <= 0:
                    st.warning("Please enter a positive amount.")
                else:
                    st.session_state.budget.append({
                        "id": generate_id(),                 # <- stable row id for unique widget keys
                        "item": item.strip(),
                        "description": desc.strip(),
                        "total_usd": float(total),
                    })
                    st.rerun()

    # Ensure each budget row has a stable id for widget keys (backfill if missing)
    for r in st.session_state.budget:
        if "id" not in r:
            r["id"] = generate_id()

    # ---------- Current budget list (summary: item / description / total) ----------
    st.markdown("### Current Budget (Summary)")
    if not st.session_state.budget:
        st.info("No budget items yet.")
    else:
        for r in st.session_state.budget:
            rid = r["id"]
            c1, c2, c3, c4 = st.columns([0.28, 0.50, 0.12, 0.10])

            # Use row id in keys to avoid duplicates across the app
            c1_val = c1.text_input("Budget item", value=r.get("item", ""), key=f"bud_item_{rid}")
            c2_val = c2.text_area("Description", value=r.get("description", ""), key=f"bud_desc_{rid}", height=70)
            c3_val = c3.number_input("Total (USD)", min_value=0.0, value=float(r.get("total_usd") or 0.0),
                                     step=500.0, format="%.2f", key=f"bud_total_{rid}")

            with c4:
                if st.button("💾", key=f"bud_save_{rid}"):
                    r["item"]        = c1_val.strip()
                    r["description"] = c2_val.strip()
                    r["total_usd"]   = float(c3_val)
                    st.rerun()
                if st.button("🗑️", key=f"bud_del_{rid}"):
                    st.session_state.budget = [x for x in st.session_state.budget if x["id"] != rid]
                    st.rerun()

        total_usd = sum(float(x.get("total_usd") or 0.0) for x in st.session_state.budget)
        st.markdown(f"**Total: USD {total_usd:,.2f}**")
        st.caption(footer_note)

# ===== TAB 6: Disbursement Schedule =====
with tabs[5]:
    st.header("💸 Disbursement Schedule")

    # Keep this table derived from KPIs (add/update/remove) while preserving date/amount
    sync_disbursement_from_kpis()

    if not st.session_state.disbursement:
        st.info("No KPIs are marked as **Linked to Payment** in the Logframe.")
    else:
        out_nums, _ = compute_numbers()
        id_to_output = {o["id"]: (o.get("name") or "Output") for o in st.session_state.outputs}

        def _out_label(d):
            n = out_nums.get(d.get("output_id",""), "")
            return f"{n} | {id_to_output.get(d.get('output_id'), '(unassigned)')}"

        st.caption("Enter the anticipated date and the amount (USD). The **Linked-KPI** comes from the Logframe and is not editable here.")
        # Render (sorted for stability)
        for row in sorted(st.session_state.disbursement, key=lambda x: (_out_label(x), x.get("kpi_name",""))):
            kpid = row["kpi_id"]
            with st.container():
                # Output label (disabled), Linked-KPI (disabled), Date (editable), Amount (editable)
                c1, c2, c3, c4 = st.columns([0.25, 0.35, 0.20, 0.20])
                c1.text_input("Output", value=_out_label(row), key=f"dsp_out_{kpid}", disabled=True)
                # make Linked-KPI explicitly read-only and visually greyed out
                c2.text_input("Linked-KPI", value=row.get("kpi_name",""), key=f"dsp_kpi_{kpid}", disabled=True)
                new_date = c3.date_input("Anticipated date", value=row.get("anticipated_date"), key=f"dsp_date_{kpid}")
                new_amt  = c4.number_input("Amount (USD)", min_value=0.0, value=float(row.get("amount_usd") or 0.0),
                                           step=1000.0, key=f"dsp_amt_{kpid}")

                # Persist only editable fields (date, amount)
                row["anticipated_date"] = new_date
                row["amount_usd"] = float(new_amt)
    st.caption(footer_note)

# ===== TAB 7: Export =====
tabs[6].header("📤 Export Your Application")
if tabs[6].button("Generate Excel Backup File"):
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    # --- Sheet 0: Identification (Project ID page) ---
    def _sum_budget_for_export():
        return sum(float(r.get("total_usd") or 0.0) for r in st.session_state.get("budget", []))

    id_info = st.session_state.get("id_info", {}) or {}

    proj_title = id_info.get("title", "")
    pi_name = id_info.get("pi_name", "")
    pi_email = id_info.get("pi_email", "")
    institution = id_info.get("institution", "")
    start_date = id_info.get("start_date", "")
    end_date = id_info.get("end_date", "")
    contact_name = id_info.get("contact_name", "")
    contact_mail = id_info.get("contact_email", "")
    contact_phone = id_info.get("contact_phone", "")

    # live computed
    budget_total = _sum_budget_for_export()
    outputs_count = len(st.session_state.get("outputs", []))
    kpis_count = len(st.session_state.get("kpis", []))

    ws_id = wb.create_sheet("Identification", 0)  # put it first
    ws_id.append(["Field", "Value"])
    ws_id.append(["Project title", proj_title])
    ws_id.append(["Principal Investigator (PI) name", pi_name])
    ws_id.append(["PI email", pi_email])
    ws_id.append(["Institution / Organization", institution])
    ws_id.append(["Project start date", fmt_dd_mmm_yyyy(start_date)])
    ws_id.append(["Project end date", fmt_dd_mmm_yyyy(end_date)])
    ws_id.append(["Contact person (optional)", contact_name])
    ws_id.append(["Contact email", contact_mail])
    ws_id.append(["Contact phone", contact_phone])

    # read-only summary values
    ws_id.append(["Funding requested (from Budget)", f"{budget_total:,.2f}"])
    ws_id.append(["Outputs (count)", outputs_count])
    ws_id.append(["KPIs (count)", kpis_count])

    # Sheet 1: Summary (Goal/Outcome/Output) — with explicit IDs
    s1 = wb.create_sheet("Summary", 1)
    s1.append(["RowLevel", "ID", "ParentID", "Text / Title", "Assumptions"])

    # Goal rows
    for row in st.session_state.get("impacts", []):
        s1.append(["Goal", row.get("id", ""), "", row.get("name", ""), ""])

    # Outcome rows
    for row in st.session_state.get("outcomes", []):
        s1.append(["Outcome", row.get("id", ""), row.get("parent_id", ""), row.get("name", ""), ""])

    # Output rows (include assumptions)
    for row in st.session_state.get("outputs", []):
        s1.append([
            "Output",
            row.get("id", ""),
            row.get("parent_id", ""),
            row.get("name", ""),
            row.get("assumptions", "")
        ])

    # Sheet 2: KPI Matrix — include KPIID and ParentID (labels are just helpers)
    s2 = wb.create_sheet("KPI Matrix")
    s2.append([
        "KPIID", "Parent Level", "ParentID", "Parent (label)",
        "KPI", "Baseline", "Target",
        "Linked to Payment", "Means of Verification"
    ])

    out_nums, kpi_nums = compute_numbers()
    output_title = {o["id"]: (o.get("name") or "Output") for o in st.session_state.outputs}
    outcome_name = (st.session_state.outcomes[0]["name"] if st.session_state.outcomes else "")

    for k in st.session_state.kpis:
        plevel = k.get("parent_level", "Output")
        pid = k.get("parent_id", "")
        parent_label = (
            f"Outcome — {outcome_name}" if plevel == "Outcome"
            else f"Output {out_nums.get(pid, '')} — {output_title.get(pid, '')}"
        )
        s2.append([
            k.get("id", ""),
            plevel,
            pid,
            parent_label,
            k.get("name", ""),
            k.get("baseline", ""),
            k.get("target", ""),
            "Yes" if k.get("linked_payment") else "No",
            k.get("mov", ""),
        ])

    # Workplan (export)
    out_nums, kpi_nums, act_nums = compute_numbers(include_activities=True)
    ws2 = wb.create_sheet("Workplan")
    ws2.append([
        "Activity ID", "Activity #",
        "OutputID", "Output",
        "Activity", "Owner", "Start", "End",
        "Linked KPI IDs", "Linked KPIs"
    ])

    id_to_output = {o["id"]: (o.get("name") or "Output") for o in st.session_state.outputs}
    id_to_kpi = {k["id"]: (k.get("name") or "") for k in st.session_state.kpis}

    for a in st.session_state.workplan:
        ws2.append([
            a.get("id", ""),
            act_nums.get(a["id"], ""),
            a.get("output_id", ""),
            id_to_output.get(a.get("output_id"), ""),
            a.get("name", ""),
            a.get("owner", ""),
            fmt_dd_mmm_yyyy(a.get("start")),
            fmt_dd_mmm_yyyy(a.get("end")),
            ",".join(a.get("kpi_ids") or []),  # machine IDs
            ", ".join(id_to_kpi.get(i, "") for i in (a.get("kpi_ids") or [])),  # display names
        ])

    # --- Budget (export) ---
    ws3 = wb.create_sheet("Budget")
    ws3.append(["Budget item", "Description", "Total Cost (USD)"])

    # Write simplified budget rows (keep amounts numeric for Excel)
    for r in st.session_state.get("budget", []):
        ws3.append([r.get("item", ""), r.get("description", ""), float(r.get("total_usd") or 0.0)])

    # Number format for Total (col C)
    for row in ws3.iter_rows(min_row=2):
        row[2].number_format = '#,##0.00'

    # --- Disbursement Schedule (export) ---
    wsd = wb.create_sheet("Disbursement Schedule")
    wsd.append([
        "KPIID",  # machine id for mapping back
        "Anticipated deliverable date",
        "Deliverable",
        "Maximum Grant instalment payable on satisfaction of this deliverable (USD)"
    ])

    from datetime import date as _date

    src = list(st.session_state.get("disbursement", [])) or [
        {
            "kpi_id": k["id"],
            "output_id": k.get("parent_id"),
            "anticipated_date": k.get("end_date") or k.get("start_date") or None,
            "deliverable": k.get("name", ""),
            "amount_usd": 0.0,
        }
        for k in st.session_state.kpis if bool(k.get("linked_payment"))
    ]

    rows = sorted(
        src,
        key=lambda d: (d.get("output_id"), d.get("anticipated_date") or _date(2100, 1, 1),
                       d.get("deliverable") or "")
    )

    for d in rows:
        # write actual date object; openpyxl will store a true Excel date
        wsd.append([
            d.get("kpi_id") or "",
            d.get("anticipated_date") or None,  # <-- date, not string
            (d.get("deliverable") or ""),
            float(d.get("amount_usd") or 0.0),
        ])

    # Number formats: Date (col B) and Amount (col D)
    for r in wsd.iter_rows(min_row=2):
        r[1].number_format = 'DD/mmm/YYYY'  # e.g., 01/Oct/2025
        r[3].number_format = '#,##0.00'

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    tabs[6].download_button(
        "📥 Download Excel File",
        data=buf,
        file_name="Application_Submission.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# --- Word export (Logframe as table)
if tabs[6].button("Generate Word Document"):
    try:
        word_buf = build_logframe_docx()
        proj_title = (st.session_state.get("id_info", {}) or {}).get("title", "") or "Project"
        safe = re.sub(r"[^A-Za-z0-9]+", "_", proj_title).strip("_") or "Project"
        tabs[6].download_button(
            "📥 Download Word Document",
            data=word_buf,
            file_name=f"Logframe_{safe}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except ModuleNotFoundError:
        tabs[6].error("`python-docx` is required. Install it with: pip install python-docx")
    except Exception as e:
        tabs[6].error(f"Could not generate the Word Document: {e}")
