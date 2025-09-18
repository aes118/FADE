# app.py
import streamlit as st

st.set_page_config(page_title="Style Toggle Demo", layout="wide")

# ----------------------------
# Fake data (replace as needed)
# ----------------------------
DATA = {
    "user_name": "Anderson",
    "kpis": {"total_apps": 2, "in_progress": 1, "submitted": 1, "approved": 0},
    "recent": [
        {"title": "Rural Water Access Improvement", "date": "Mar 1, 2024", "amount": 750_000, "status": "draft"},
        {"title": "Digital Literacy for Primary Schools", "date": "Jun 19, 2024", "amount": 150_000, "status": "submitted"},
    ],
}

# ----------------------------
# CSS injectors
# ----------------------------
def inject_original_css():
    """Close to vanilla Streamlit."""
    st.markdown(
        """
        <style>
          .hero h1{margin-bottom:8px}
          .hero p{color:#6b7280;margin-top:0}
          .card{border:1px solid #e5e7eb;border-radius:10px;padding:12px;background:#fff}
          .kpi-value{font-size:1.6rem;font-weight:700;margin:0}
          .kpi-title{color:#6b7280;margin-bottom:6px}
          .chip{display:inline-block;padding:2px 8px;border-radius:999px;border:1px solid #e5e7eb;background:#f9fafb;font-size:.85rem;color:#374151}
          .app{border:1px solid #e5e7eb;border-radius:10px;padding:12px;margin:8px 0;background:#fff}
          .meta{color:#6b7280;font-size:.9rem}
          .money{font-variant-numeric:tabular-nums;font-weight:600}
        </style>
        """,
        unsafe_allow_html=True,
    )

def inject_base44_css():
    st.markdown("""
    <style>
      /* Card shell */
      .kpi-fancy{
        position:relative;
        background:#fff;
        border:1px solid #eef2f7;
        border-radius:22px;
        padding:22px 22px 18px;
        box-shadow:0 10px 28px rgba(2,6,23,.06);
        overflow:hidden; /* clip corner shapes */
      }

      .kpi-title{color:#334155;font-weight:700;font-size:1.05rem;margin:0 0 14px}
      .kpi-value{color:#0f172a;font-size:2.35rem;font-weight:800;line-height:1;margin:0 0 10px}
      .kpi-trend{display:flex;gap:8px;align-items:center;color:#15803d;font-weight:600;font-size:.95rem}
      .kpi-trend svg{width:16px;height:16px}

      /* ===== Corner shapes (match Base44) =====
         ::after  = small quarter-circle splash (tight)
         ::before = soft rounded square "tile" behind the icon
      */
      .kpi-fancy::after{
        content:"";
        position:absolute;
        top:-18px;          /* nudge out so only a quarter shows */
        right:-18px;
        width:120px;        /* small arc, not sweeping */
        height:120px;
        border-bottom-left-radius:120px;   /* quarter circle */
        background:#e0e7ff; /* default, overridden by variants */
        pointer-events:none;
        z-index:1;
      }
      .kpi-fancy::before{
        content:"";
        position:absolute;
        top:30px;           /* sits under the icon */
        right:30px;
        width:64px;
        height:64px;
        border-radius:16px; /* rounded square tile */
        background:rgba(37,99,235,.06);  /* default tint; overridden by variants */
        box-shadow: inset 0 0 0 1px rgba(37,99,235,.06);
        z-index:1;
      }

      /* Icon badge (above both shapes) */
      .kpi-icon{
        position:absolute; top:34px; right:34px;
        width:44px; height:44px; border-radius:14px;
        display:grid; place-items:center;
        background:#fff;
        border:1px solid #e5e7eb;
        box-shadow:0 2px 10px rgba(2,6,23,.06);
        z-index:2;
      }
      .kpi-icon svg{width:22px;height:22px}

      /* ===== Color variants (splash + tile + icon tint) ===== */
      .kpi--blue::after   { background:#e0e7ff; }
      .kpi--blue::before  { background:rgba(37,99,235,.06); box-shadow: inset 0 0 0 1px rgba(37,99,235,.08); }
      .kpi--blue  .kpi-icon{ color:#2563eb; border-color:#dbeafe }

      .kpi--amber::after  { background:#fef3c7; }
      .kpi--amber::before { background:rgba(217,119,6,.08); box-shadow: inset 0 0 0 1px rgba(217,119,6,.10); }
      .kpi--amber .kpi-icon{ color:#d97706; border-color:#fde68a }

      .kpi--violet::after { background:#ede9fe; }
      .kpi--violet::before{ background:rgba(124,58,237,.08); box-shadow: inset 0 0 0 1px rgba(124,58,237,.10); }
      .kpi--violet .kpi-icon{ color:#7c3aed; border-color:#ddd6fe }

      .kpi--green::after  { background:#dcfce7; }
      .kpi--green::before { background:rgba(22,163,74,.08); box-shadow: inset 0 0 0 1px rgba(22,163,74,.10); }
      .kpi--green .kpi-icon{ color:#16a34a; border-color:#bbf7d0 }
    </style>
    """, unsafe_allow_html=True)


def kpi_card_fancy(title, value, sub, color="blue", icon_svg=""):
    """Base44-style KPI card."""
    arrow_up = """
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M7 11l5-5 5 5"></path><path d="M12 6v12"></path>
      </svg>
    """
    icon = icon_svg or """
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <rect x="4" y="3" width="16" height="18" rx="2"></rect>
        <path d="M8 7h8M8 11h8M8 15h5"></path>
      </svg>
    """
    st.markdown(
        f"""
        <div class="kpi-fancy kpi--{color}">
          <div class="kpi-icon">{icon}</div>
          <div class="kpi-title">{title}</div>
          <div class="kpi-value">{value}</div>
          <div class="kpi-trend">{arrow_up} <span>{sub}</span></div>
        </div>
        """,
        unsafe_allow_html=True
    )

def app_row(title, date_txt, amount_txt, status="draft", base44=False):
    chip_class = {"draft":"yellow","submitted":"blue","approved":"green"}.get(status,"") if base44 else ""
    container_cls = "row-card" if base44 else "app"
    st.markdown(
        f"""
        <div class="{container_cls}">
          <div>
            <div style="font-weight:600">{title}</div>
            <div class="meta">📅 {date_txt} &nbsp;&nbsp; <span class="money">$ {amount_txt}</span></div>
          </div>
          <div><span class="chip {chip_class}">{status}</span></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# ----------------------------
# UI – style toggle + same content
# ----------------------------
st.sidebar.title("Style")
style = st.sidebar.radio("Choose look", ["Original", "Base44 (fancy)"], index=1)

if style.startswith("Base44"):
    inject_base44_css()
else:
    inject_original_css()

# Header row with optional CTA (simple vs fancy)
left, right = st.columns([0.7, 0.3])
with left:
    st.markdown(
        f"""
        <div class="hero">
          <h1>Welcome back, {DATA['user_name']}</h1>
          <p>Manage your grant applications and track progress</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
with right:
    if style.startswith("Base44"):
        st.markdown('<div style="text-align:right;"><button style="display:inline-flex;gap:8px;align-items:center;padding:10px 14px;border-radius:12px;background:#2563eb;color:#fff;border:0;cursor:pointer;box-shadow:0 6px 18px rgba(37,99,235,.25);">➕ New Application</button></div>', unsafe_allow_html=True)
    else:
        st.button("New Application", use_container_width=True)

st.divider()

# KPI cards
k = DATA["kpis"]
if style.startswith("Base44"):
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card_fancy("Total Applications", k["total_apps"], "+12% from last month",
                       color="blue",
                       icon_svg="""
                         <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                           <path d="M8 21V9m4 12V3m4 18v-6"></path>
                         </svg>""")
    with c2:
        kpi_card_fancy("In Progress", k["in_progress"], "3 pending review",
                       color="amber",
                       icon_svg="""
                         <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                           <circle cx="12" cy="12" r="9"></circle>
                           <path d="M12 7v5l3 2"></path>
                         </svg>""")
    with c3:
        kpi_card_fancy("Submitted", k["submitted"], "2 this week",
                       color="violet",
                       icon_svg="""
                         <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                           <path d="M3 17l6-6 4 4 8-8"></path>
                         </svg>""")
    with c4:
        kpi_card_fancy("Approved", k["approved"], "85% success rate",
                       color="green",
                       icon_svg="""
                         <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                           <path d="M20 6l-11 11-5-5"></path>
                         </svg>""")
else:
    c1, c2, c3, c4 = st.columns(4)
    with c1: kpi_card_simple("Total Applications", k["total_apps"])
    with c2: kpi_card_simple("In Progress", k["in_progress"])
    with c3: kpi_card_simple("Submitted", k["submitted"])
    with c4: kpi_card_simple("Approved", k["approved"])

st.markdown("### Recent Applications")
for row in DATA["recent"]:
    app_row(row["title"], row["date"], fmt_money(row["amount"]),
            status=row["status"], base44=style.startswith("Base44"))

st.markdown("### Notes")
st.write("Flip the **Style** in the left sidebar to compare spacing, shadows, chips, icons, and typography.")
