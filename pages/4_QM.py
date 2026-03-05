# pages/4_QM.py - Quality Maintenance Pillar
import os, io, sys
from textwrap import fill
from datetime import datetime
from urllib.parse import quote

import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import matplotlib as mpl
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib.colors import HexColor

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from qm_pipeline import (
    read_excel_from_upload, build_dataset_final_issued, build_dataset_ncr,
    FINAL_APPROVAL_COL, CREATION_DATETIME_COL,
)
from utils.supabase_client import get_supabase

os.environ["MPLCONFIGDIR"] = os.path.join(os.getcwd(), ".mplconfig")
st.set_page_config(page_title="QM Pillar", page_icon="✅", layout="wide")

# ── Auth guard ──
if "user" not in st.session_state:
    st.switch_page("app.py")

supabase = get_supabase()
role     = st.session_state.get("role", "member")
pillar   = st.session_state.get("pillar", "ALL")
name     = st.session_state.get("name", "User")
can_edit = (role == "plant_manager") or (role == "pillar_leader" and pillar == "QM")

# ── Header ──
col1, col2 = st.columns([5, 1])
with col1:
    st.markdown("# ✅ Quality Maintenance")
    st.markdown(f"Logged in as **{name}** · `{role}` · {'✏️ Full Access' if can_edit else '👁️ View Only'}")
with col2:
    if st.button("🏠 Home", use_container_width=True):
        st.switch_page("pages/1_Home.py")

st.divider()

tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Maturity Level",
    "📈 Quality Indicators",
    "📁 Documents",
    "🎯 Action Plans",
])

# ════════════════════════════════════════
# TAB 1 — MATURITY LEVEL
# ════════════════════════════════════════
with tab1:
    st.markdown("### TPM Quality Maintenance Maturity Level Tracker")
    st.markdown("Current Level: 🥇 **Gold**")
    st.progress(1.0)
    levels = [
        {"name": "Entry", "icon": "⚪", "focus": "QM by basic methods. Identification of Q factors.",
         "criteria": ["QM Pillar and targets defined (Cost of non-Quality, CC, Internal Rejections)",
                      "Basic Root Cause Analysis implemented",
                      "Action plans generated to prevent defect re-occurrence",
                      "Identification of Q factors (Machine, Material, Manpower, Method)",
                      "QA Matrix tool developed","KPIs monitored and suggestions for improvement available",
                      "Customer Complaints trend is decreasing"]},
        {"name": "Bronze", "icon": "🥉", "focus": "Reduce Waste and Complaints through a Quality system.",
         "criteria": ["Analysis implemented for all Q factors",
                      "QA Matrix used to tackle high risk defects",
                      "Action plans effective (IR and CC data reduced by 10%)",
                      "Accurate and updated MPCC for all FGs",
                      "Quality skills assessed and improved for all shopfloor employees",
                      "QC council meetings implemented",
                      "Shopfloor/QC/MT employees have basic knowledge on Q factors"]},
        {"name": "Silver", "icon": "🥈", "focus": "Quality Awareness is clear.",
         "criteria": ["Reduction of NCR & CC ratio by 10% vs last year",
                      "Evidence of 2 major Quality improvement projects",
                      "QA matrix showing decrease in RPN of 50% of defects",
                      "Cost of Poor Quality actual equals target COPQ",
                      "Evidence of shopfloor involvement in NCR and CC RCAs",
                      "Quality circles used to transfer knowledge",
                      "Clear evidence of Q Points for machines, materials, methods, men and measurement"]},
        {"name": "Gold", "icon": "🥇", "focus": "Quality is embedded in the BU culture.",
         "criteria": ["Decreasing monthly trend of NCR & CC ratio",
                      "Quality circles transfer knowledge from engineers to operators and vice versa",
                      "QA Matrix used for all machine groups, updated quarterly",
                      "Evidence of 2 major Quality improvement projects",
                      "Formalised links between defects and Q Factors with physical evidence on shopfloor",
                      "Quality defined for the BU"]},
    ]
    for level in levels:
        is_current = level["name"] == "Gold"
        with st.expander(f"{level['icon']} {level['name']} — {level['focus']}", expanded=is_current):
            if is_current: st.success("✅ Current Level")
            for criterion in level["criteria"]:
                st.checkbox(criterion, value=is_current, key=f"qm_lvl_{level['name']}_{criterion[:30]}")

# ════════════════════════════════════════
# TAB 2 — QUALITY INDICATORS (full original tool)
# ════════════════════════════════════════
with tab2:
    st.markdown("### 📈 CRM Quality Dashboard")

    # matplotlib settings (same as original)
    mpl.rcParams["font.family"] = "DejaVu Sans"
    mpl.rcParams["font.size"] = 11
    mpl.rcParams["figure.dpi"] = 160
    mpl.rcParams["savefig.dpi"] = 320
    mpl.rcParams["savefig.bbox"] = "tight"
    COLORS = {"Service": "#C1A02E", "Quality": "#006394", "Invalid": "#D8C37D"}

    # ── File uploads ──
    col1, col2, col3 = st.columns(3)
    with col1: final_file  = st.file_uploader("📂 FINAL Approval file", type=["xls","xlsx"], key="qm_final")
    with col2: issued_file = st.file_uploader("📂 ISSUED file",         type=["xls","xlsx"], key="qm_issued")
    with col3: ncr_file    = st.file_uploader("📂 NCR file",            type=["xls","xlsx"], key="qm_ncr")

    if not final_file or not issued_file or not ncr_file:
        st.info("⬆️ Upload all 3 files above to enable the dashboard.")
    else:
        try:
            with st.spinner("Loading data..."):
                df_final_loaded  = read_excel_from_upload(final_file)
                df_issued_loaded = read_excel_from_upload(issued_file)
                df_ncr_loaded    = read_excel_from_upload(ncr_file)

                final_pkg  = build_dataset_final_issued(df_final_loaded,  date_col=FINAL_APPROVAL_COL,    dataset_name="FINAL")
                issued_pkg = build_dataset_final_issued(df_issued_loaded, date_col=CREATION_DATETIME_COL, dataset_name="ISSUED")
                ncr_pkg    = build_dataset_ncr(df_ncr_loaded,             date_col=FINAL_APPROVAL_COL,    dataset_name="NCR")

                df_final  = final_pkg["cleaned_flagged"].copy()
                df_issued = issued_pkg["cleaned_flagged"].copy()
                df_ncr    = ncr_pkg["cleaned_flagged"].copy()

                for _d in (df_final, df_issued, df_ncr):
                    _d["Year"]  = pd.to_numeric(_d["Year"],  errors="coerce").astype("Int64")
                    _d["Month"] = pd.to_numeric(_d["Month"], errors="coerce").astype("Int64")
                df_final  = df_final.dropna(subset=["Year","Month"])
                df_issued = df_issued.dropna(subset=["Year","Month"])
                df_ncr    = df_ncr.dropna(subset=["Year","Month"])
                df_final["Year"]  = df_final["Year"].astype(int)
                df_final["Month"] = df_final["Month"].astype(int)
                df_issued["Year"] = df_issued["Year"].astype(int)
                df_issued["Month"]= df_issued["Month"].astype(int)
                df_ncr["Year"]    = df_ncr["Year"].astype(int)
                df_ncr["Month"]   = df_ncr["Month"].astype(int)

                df_final_raw_flagged = final_pkg["raw_flagged"]
                df = df_final
                df_raw_flagged = df_final_raw_flagged
                df_ncr_dash = df_ncr

            st.success("✅ Files loaded!")

            # ── Period selector ──
            df_dates = df_final[["Year","Month"]].drop_duplicates().sort_values(["Year","Month"])
            years = sorted(df_dates["Year"].unique().tolist())
            col1, col2 = st.columns(2)
            with col1: selected_year  = st.selectbox("Year",  years, index=len(years)-1, key="qm_year")
            with col2:
                months_avail = sorted(df_dates[df_dates["Year"]==selected_year]["Month"].unique().tolist())
                selected_month = int(st.selectbox("Month", months_avail, index=len(months_avail)-1, key="qm_month"))

            # ── Top-N controls ──
            with st.expander("⚙️ Display Controls", expanded=False):
                c1,c2,c3,c4,c5,c6,c7,c8 = st.columns(8)
                def topn(col, label, default, max_n=40):
                    opts = ["All"]+list(range(3,max_n+1))
                    v = col.selectbox(label, opts, index=opts.index(default) if default in opts else 0)
                    return None if v=="All" else int(v)
                TOPN_QUALITY_DEFECT = topn(c1,"Q Defects CM vs YTD",15)
                TOPN_QUALITY_CM     = topn(c2,"Q Defects CM",10)
                TOPN_SERVICE_REASON = topn(c3,"Svc Reasons CM vs YTD",15)
                TOPN_SERVICE_CM     = topn(c4,"Svc Reasons CM",10)
                TOPN_COST_DEFECT    = topn(c5,"Cost Defects",15)
                TOPN_NCR_DEFECT     = topn(c6,"NCR Defects",20,50)
                TOPN_CORRELATION    = topn(c7,"NCR vs CRM Corr",12)
                TOPN_CUSTOMER       = topn(c8,"Top Customers",5,20)

            run = st.button("🔄 Generate Charts", type="primary", key="qm_run")

            if run:
                # ── Helpers (exact same as original) ──
                def fig_to_png_bytes(fig, dpi=300):
                    buf = io.BytesIO()
                    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight")
                    buf.seek(0)
                    return buf

                def show_fig(fig, dpi=300):
                    st.image(fig_to_png_bytes(fig, dpi=dpi))
                    plt.close(fig)

                def ppt_slide_title_bar(c, title, W, H):
                    c.setFillColor(HexColor("#006394")); c.setFont("Helvetica-Bold", 26)
                    c.drawString(40, H-58, title)
                    c.setStrokeColor(HexColor("#006394")); c.setLineWidth(4)
                    c.line(40, H-78, W-40, H-78)

                def build_ppt_pdf(slides, dpi=300):
                    W, H = 960, 540
                    pdf_buf = io.BytesIO()
                    c = canvas.Canvas(pdf_buf, pagesize=(W, H))
                    for s in slides:
                        ppt_slide_title_bar(c, s["title"], W, H)
                        img = ImageReader(fig_to_png_bytes(s["fig"], dpi=dpi))
                        c.drawImage(img, 40, 40, width=W-80, height=H-92-40, preserveAspectRatio=True, anchor="c")
                        c.showPage()
                        plt.close(s["fig"])
                    c.save(); pdf_buf.seek(0)
                    return pdf_buf

                def donut_chart(ax, counts):
                    labels = ["Service","Quality","Invalid"]
                    values = np.array([int(counts.get(l,0)) for l in labels], dtype=float)
                    if values.sum()==0:
                        ax.text(0.5,0.5,"No data",ha="center",va="center"); ax.axis("off"); return
                    rw=0.45; lr=(1.0-rw)+rw/2
                    wedges,_=ax.pie(values,labels=None,colors=[COLORS[l] for l in labels],
                                    startangle=90,counterclock=False,radius=1.0,
                                    wedgeprops=dict(width=rw,edgecolor="white",linewidth=1))
                    for w,v in zip(wedges,values):
                        if v<=0: continue
                        ang=(w.theta2+w.theta1)/2.0
                        ax.text(lr*np.cos(np.deg2rad(ang)),lr*np.sin(np.deg2rad(ang)),
                                f"{int(v)}",ha="center",va="center",fontsize=11,color="#4D4D4D")
                    ax.axis("equal")

                def add_simple_value_labels(ax, bars, fmt_fn, pad):
                    for b in bars:
                        h=b.get_height()
                        if np.isnan(h) or h==0: continue
                        ax.text(b.get_x()+b.get_width()/2,h+pad,fmt_fn(h),
                                ha="center",va="bottom",fontsize=10,color="#4D4D4D",clip_on=False)

                MONTH_LABELS=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
                prev_year = selected_year - 1
                months_range = list(range(1, selected_month+1))
                slides_for_pdf = []

                # ══ FINAL ══
                st.header("FINAL (CRM Approved)")

               # 5) Slide figs (each returns fig only)
# =========================
def slide_1_final_donuts_fig(year, month):
    df_month = df[(df["Year"] == year) & (df["Month"] == month)]
    df_ytd   = df[(df["Year"] == year) & (df["Month"].between(1, month))]

    month_counts = df_month["Complaint_Category"].value_counts()
    ytd_counts   = df_ytd["Complaint_Category"].value_counts()

    fig, axes = plt.subplots(1, 2, figsize=(13.33, 7.5), dpi=300)

    donut_chart(axes[0], month_counts)
    donut_chart(axes[1], ytd_counts)

    # small captions under donuts (NOT titles)
    m_lbl = pd.to_datetime(f"{year}-{month:02d}-01").strftime("%b-%y")
    axes[0].text(0.5, -0.10, m_lbl, transform=axes[0].transAxes, ha="center", va="top", fontsize=13, color="#4D4D4D")
    axes[1].text(0.5, -0.10, f"{year} YTD", transform=axes[1].transAxes, ha="center", va="top", fontsize=13, color="#4D4D4D")

    fig.legend(["Service", "Quality", "Invalid"], loc="lower center", ncol=3, frameon=False, prop={"size": 14})
    plt.tight_layout(rect=[0, 0.08, 1, 1])
    return fig

def build_unique_crm_leadtime_table(df_in):
    tmp = df_in.copy()
    tmp["First_Approval_dt"] = pd.to_datetime(tmp["First Approval Date"], errors="coerce")
    tmp["Final_Approval_dt"] = pd.to_datetime(tmp["Final Approval Date"], errors="coerce")

    g = (
        tmp.groupby("Ref NB", as_index=False)
           .agg(
               First_Approval_dt=("First_Approval_dt", "min"),
               Final_Approval_dt=("Final_Approval_dt", "max"),
           )
    )

    g["LeadTime_days"] = (g["Final_Approval_dt"] - g["First_Approval_dt"]).dt.total_seconds() / 86400.0
    g["Year"]  = g["Final_Approval_dt"].dt.year
    g["Month"] = g["Final_Approval_dt"].dt.month

    g = g.dropna(subset=["Year", "Month", "LeadTime_days"])
    g = g[g["LeadTime_days"] >= 0]
    return g

df_lt = build_unique_crm_leadtime_table(df)

def slide_2_leadtime_fig(selected_year, selected_month):
    prev_year = selected_year - 1
    month_labels_all = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

    months = list(range(1, selected_month + 1))
    month_labels = month_labels_all[:selected_month]

    def monthly_avg(year):
        return (
            df_lt[df_lt["Year"] == year]
            .groupby("Month")["LeadTime_days"]
            .mean()
            .reindex(months)
        )

    def ytd_avg(year):
        d = df_lt[(df_lt["Year"] == year) & (df_lt["Month"].between(1, selected_month))]
        return float(d["LeadTime_days"].mean()) if len(d) else np.nan

    m_prev = monthly_avg(prev_year)
    m_sel  = monthly_avg(selected_year)
    ytd_prev = ytd_avg(prev_year)
    ytd_sel  = ytd_avg(selected_year)

    cats = month_labels + ["AVG-YTD"]
    prev_vals = list(m_prev.values) + [ytd_prev]
    sel_vals  = list(m_sel.values)  + [ytd_sel]

    x = np.arange(len(cats))
    w = 0.28

    fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)
    b1 = ax.bar(x - w/2, prev_vals, width=w, label=str(prev_year), color="#006394")
    b2 = ax.bar(x + w/2, sel_vals,  width=w, label=str(selected_year), color="#C1A02E")

    add_simple_value_labels(ax, b1, lambda v: f"{v:.0f}", pad=0.6)
    add_simple_value_labels(ax, b2, lambda v: f"{v:.0f}", pad=0.6)

    ax.set_xticks(x)
    ax.set_xticklabels(cats)
    ax.set_ylabel("Lead Time (days)")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.08), ncol=2, frameon=False)

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(False)

    plt.tight_layout(rect=[0, 0.05, 1, 1])
    return fig

def slide_3_valid_quality_count_fig(selected_year, selected_month):
    prev_year = selected_year - 1
    month_labels_all = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

    months = list(range(1, selected_month + 1))
    month_labels = month_labels_all[:selected_month]
    cats = month_labels + ["YTD"]

    base = df[(df["Is_Valid"] == True) & (df["Complaint_Category"] == "Quality")].copy()

    def monthly_counts(year):
        return (
            base[base["Year"] == year]
            .groupby("Month")
            .size()
            .reindex(months, fill_value=0)
        )

    def ytd_count(year):
        return int(base[(base["Year"] == year) & (base["Month"].between(1, selected_month))].shape[0])

    c_prev = monthly_counts(prev_year)
    c_sel  = monthly_counts(selected_year)
    ytd_prev = ytd_count(prev_year)
    ytd_sel  = ytd_count(selected_year)

    prev_vals = list(c_prev.values) + [ytd_prev]
    sel_vals  = list(c_sel.values)  + [ytd_sel]

    x = np.arange(len(cats))
    w = 0.28

    fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)
    b1 = ax.bar(x - w/2, prev_vals, width=w, label=str(prev_year), color="#006394")
    b2 = ax.bar(x + w/2, sel_vals,  width=w, label=str(selected_year), color="#C1A02E")

    add_simple_value_labels(ax, b1, lambda v: f"{int(v)}", pad=1.0)
    add_simple_value_labels(ax, b2, lambda v: f"{int(v)}", pad=1.0)

    # trendline for selected year months only (exclude YTD)
    y = np.array(list(c_sel.values), dtype=float)
    x_months = np.arange(len(months), dtype=float)
    if len(months) >= 2 and np.any(y > 0):
        coeff = np.polyfit(x_months, y, 1)
        y_fit = np.polyval(coeff, x_months)
        ax.plot(x[:len(months)], y_fit, linestyle="--", linewidth=1.5, color="#C1A02E", label=f"Trend {selected_year}")

    ax.set_xticks(x)
    ax.set_xticklabels(cats)
    ax.set_ylabel("Count")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.08), ncol=3, frameon=False)

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(False)

    plt.tight_layout(rect=[0, 0.05, 1, 1])
    return fig

def slide_4_quality_defect_cm_vs_ytd_fig(selected_year, selected_month, top_n=15):
    prev_year = selected_year - 1

    base = df[(df["Is_Valid"] == True) & (df["Complaint_Category"] == "Quality")].copy()
    base["Defect"] = base["Reason"].astype(str).str.strip()

    cm_prev  = base[(base["Year"] == prev_year) & (base["Month"] == selected_month)]
    cm_sel   = base[(base["Year"] == selected_year) & (base["Month"] == selected_month)]
    ytd_prev = base[(base["Year"] == prev_year) & (base["Month"].between(1, selected_month))]
    ytd_sel  = base[(base["Year"] == selected_year) & (base["Month"].between(1, selected_month))]

    cm_prev_counts  = cm_prev["Defect"].value_counts()
    cm_sel_counts   = cm_sel["Defect"].value_counts()
    ytd_prev_counts = ytd_prev["Defect"].value_counts()
    ytd_sel_counts  = ytd_sel["Defect"].value_counts()

    all_defects = (
        pd.Index(ytd_sel_counts.index)
        .union(ytd_prev_counts.index)
        .union(cm_sel_counts.index)
        .union(cm_prev_counts.index)
    )

    summary = pd.DataFrame(index=all_defects)
    summary["CM_prev"]  = cm_prev_counts.reindex(all_defects, fill_value=0).astype(int)
    summary["CM_sel"]   = cm_sel_counts.reindex(all_defects, fill_value=0).astype(int)
    summary["YTD_prev"] = ytd_prev_counts.reindex(all_defects, fill_value=0).astype(int)
    summary["YTD_sel"]  = ytd_sel_counts.reindex(all_defects, fill_value=0).astype(int)

    summary = summary.sort_values(["YTD_sel", "CM_sel"], ascending=False)
    if top_n is not None:
        summary = summary.head(int(top_n))

    cm_title_prev = pd.to_datetime(f"{prev_year}-{selected_month:02d}-01").strftime("%b-%y")
    cm_title_sel  = pd.to_datetime(f"{selected_year}-{selected_month:02d}-01").strftime("%b-%y")

    defects = summary.index.tolist()
    wrapped_labels = [fill(d, width=16) for d in defects]

    x = np.arange(len(defects))
    w = 0.18

    fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)

    b1 = ax.bar(x - 1.5*w, summary["CM_prev"],  width=w, color="#0F68B9", label=f"CM {cm_title_prev}")
    b2 = ax.bar(x - 0.5*w, summary["CM_sel"],   width=w, color="#D8C37D", label=f"CM {cm_title_sel}")
    b3 = ax.bar(x + 0.5*w, summary["YTD_prev"], width=w, color="#006394", label=f"YTD {prev_year}")
    b4 = ax.bar(x + 1.5*w, summary["YTD_sel"],  width=w, color="#C1A02E", label=f"YTD {selected_year}")

    add_simple_value_labels(ax, b1, lambda v: f"{int(v)}", pad=1.0)
    add_simple_value_labels(ax, b2, lambda v: f"{int(v)}", pad=1.0)
    add_simple_value_labels(ax, b3, lambda v: f"{int(v)}", pad=1.0)
    add_simple_value_labels(ax, b4, lambda v: f"{int(v)}", pad=1.0)

    ax.set_xticks(x)
    ax.set_xticklabels(wrapped_labels, rotation=35, ha="right", fontsize=9)

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(False)

    ax.legend(
        loc="upper center",
        bbox_to_anchor=(0.5, -0.22),   # lower legend
        ncol=4,
        frameon=False,
        fontsize=10
    )

    fig.subplots_adjust(bottom=0.35)   # give legend breathing room

    return fig

def slide_5_quality_defect_current_month_fig(selected_year, selected_month, top_n=10):
    base = df[(df["Is_Valid"] == True) & (df["Complaint_Category"] == "Quality")].copy()
    base["Defect"] = base["Reason"].astype(str).str.strip()

    cm = base[(base["Year"] == selected_year) & (base["Month"] == selected_month)]
    counts = cm["Defect"].value_counts()

    if top_n is not None:
        counts = counts.head(int(top_n))

    fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)

    x = np.arange(len(counts))
    bars = ax.bar(x, counts.values, color="#C1A02E", width=0.6)

    add_simple_value_labels(ax, bars, lambda v: f"{int(v)}", pad=0.6)

    labels = [fill(str(s), width=18) for s in counts.index.tolist()]
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=10, rotation=25, ha="right")

    ax.set_ylim(0, max(1, counts.max()) * 1.15)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(False)

    plt.tight_layout()
    return fig

def slide_6_quality_cost_cm_vs_ytd_fig(selected_year, selected_month, top_n=15, defect_col="Reason"):
    prev_year = selected_year - 1
    COST_COL = "Cost Amount"

    # --- detect Decision column ---
    decision_candidates = [c for c in df_raw_flagged.columns if "decision" in str(c).lower()]
    if not decision_candidates:
        raise KeyError("Could not find Decision column.")
    DECISION_COL = decision_candidates[0]

    # --- base filtering ---
    base = df_raw_flagged[
        (df_raw_flagged["Is_Valid"] == True) &
        (df_raw_flagged["Complaint_Category"] == "Quality")
    ].copy()

    base[DECISION_COL] = base[DECISION_COL].astype(str).str.strip()
    base = base[base[DECISION_COL].str.lower() == "credit note"]

    base["Defect"] = base[defect_col].astype(str).str.replace("\n", " ").replace("\r", " ").str.strip()
    base[COST_COL] = pd.to_numeric(base[COST_COL], errors="coerce").fillna(0)

    # --- period splits ---
    cm_prev  = base[(base["Year"] == prev_year) & (base["Month"] == selected_month)]
    cm_sel   = base[(base["Year"] == selected_year) & (base["Month"] == selected_month)]
    ytd_prev = base[(base["Year"] == prev_year) & (base["Month"].between(1, selected_month))]
    ytd_sel  = base[(base["Year"] == selected_year) & (base["Month"].between(1, selected_month))]

    summary = pd.DataFrame({
        "CM_prev":  cm_prev.groupby("Defect")[COST_COL].sum(),
        "CM_sel":   cm_sel.groupby("Defect")[COST_COL].sum(),
        "YTD_prev": ytd_prev.groupby("Defect")[COST_COL].sum(),
        "YTD_sel":  ytd_sel.groupby("Defect")[COST_COL].sum(),
    }).fillna(0)

    summary = summary.sort_values(["YTD_sel", "CM_sel"], ascending=False)
    if top_n:
        summary = summary.head(int(top_n))

    defects = summary.index.tolist()
    labels  = [fill(d, 18) for d in defects]

    x = np.arange(len(defects))
    w = 0.18

    fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)

    b1 = ax.bar(x - 1.5*w, summary["CM_prev"],  w, label=f"CM {prev_year}", color="#0F68B9")
    b2 = ax.bar(x - 0.5*w, summary["CM_sel"],   w, label=f"CM {selected_year}", color="#D8C37D")
    b3 = ax.bar(x + 0.5*w, summary["YTD_prev"], w, label=f"YTD {prev_year}", color="#006394")
    b4 = ax.bar(x + 1.5*w, summary["YTD_sel"],  w, label=f"YTD {selected_year}", color="#C1A02E")

    # ---------- VALUE LABELS (ALL, VERTICAL, ABOVE BAR) ----------
    ymax = summary.to_numpy().max()
    pad  = ymax * 0.015

    def fmt(v):
        return f"SAR {int(v/1000)}K" if v >= 1000 else f"SAR {int(v)}"

    for bars in (b1, b2, b3, b4):
        for bar in bars:
            h = bar.get_height()
            if h > 0:
                ax.text(
                    bar.get_x() + bar.get_width()/2,
                    h + pad,
                    fmt(h),
                    ha="center",
                    va="bottom",
                    fontsize=9,
                    rotation=90,
                    color="#333333",
                    clip_on=False
                )

    # ---------- AXES & LABELS ----------
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=30, ha="right", fontsize=9)
    ax.set_ylabel("Cost Amount (SAR)")

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(False)

    # ---------- LEGEND LOWER ----------
    ax.legend(
        loc="upper center",
        bbox_to_anchor=(0.5, -0.22),
        ncol=4,
        frameon=False,
        fontsize=10
    )

    fig.subplots_adjust(bottom=0.35)
    return fig



def slide_s1_valid_service_count_fig(selected_year, selected_month):
    prev_year = selected_year - 1
    month_labels_all = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    months = list(range(1, selected_month + 1))
    month_labels = month_labels_all[:selected_month]
    cats = month_labels + ["YTD"]

    base = df[(df["Is_Valid"] == True) & (df["Complaint_Category"] == "Service")].copy()

    def monthly_counts(year):
        return (
            base[base["Year"] == year]
            .groupby("Month")
            .size()
            .reindex(months, fill_value=0)
        )

    def ytd_count(year):
        return int(base[(base["Year"] == year) & (base["Month"].between(1, selected_month))].shape[0])

    c_prev = monthly_counts(prev_year)
    c_sel  = monthly_counts(selected_year)
    ytd_prev = ytd_count(prev_year)
    ytd_sel  = ytd_count(selected_year)

    prev_vals = list(c_prev.values) + [ytd_prev]
    sel_vals  = list(c_sel.values)  + [ytd_sel]

    x = np.arange(len(cats))
    w = 0.28

    fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)

    b1 = ax.bar(x - w/2, prev_vals, width=w, label=str(prev_year), color="#006394")
    b2 = ax.bar(x + w/2, sel_vals,  width=w, label=str(selected_year), color="#C1A02E")

    add_simple_value_labels(ax, b1, lambda v: f"{int(v)}", pad=1.0)
    add_simple_value_labels(ax, b2, lambda v: f"{int(v)}", pad=1.0)

    y = np.array(list(c_sel.values), dtype=float)
    x_months = np.arange(len(months), dtype=float)
    if len(months) >= 2 and np.any(y > 0):
        coeff = np.polyfit(x_months, y, 1)
        y_fit = np.polyval(coeff, x_months)
        ax.plot(x[:len(months)], y_fit, linestyle="--", linewidth=1.5, color="#C1A02E", label=f"Trend {selected_year}")

    ax.set_xticks(x)
    ax.set_xticklabels(cats)
    ax.set_ylabel("Count")

    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.08), ncol=3, frameon=False)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(False)

    plt.tight_layout(rect=[0, 0.05, 1, 1])
    return fig

def slide_s2_service_reason_cm_vs_ytd_fig(selected_year, selected_month, top_n=15):
    prev_year = selected_year - 1

    base = df[(df["Is_Valid"] == True) & (df["Complaint_Category"] == "Service")].copy()
    base["Reason_Svc"] = base["Reason"].astype(str).str.strip()

    cm_prev  = base[(base["Year"] == prev_year) & (base["Month"] == selected_month)]
    cm_sel   = base[(base["Year"] == selected_year) & (base["Month"] == selected_month)]
    ytd_prev = base[(base["Year"] == prev_year) & (base["Month"].between(1, selected_month))]
    ytd_sel  = base[(base["Year"] == selected_year) & (base["Month"].between(1, selected_month))]

    cm_prev_counts  = cm_prev["Reason_Svc"].value_counts()
    cm_sel_counts   = cm_sel["Reason_Svc"].value_counts()
    ytd_prev_counts = ytd_prev["Reason_Svc"].value_counts()
    ytd_sel_counts  = ytd_sel["Reason_Svc"].value_counts()

    all_reasons = (
        pd.Index(ytd_sel_counts.index)
        .union(ytd_prev_counts.index)
        .union(cm_sel_counts.index)
        .union(cm_prev_counts.index)
    )

    summary = pd.DataFrame(index=all_reasons)
    summary["CM_prev"]  = cm_prev_counts.reindex(all_reasons, fill_value=0).astype(int)
    summary["CM_sel"]   = cm_sel_counts.reindex(all_reasons, fill_value=0).astype(int)
    summary["YTD_prev"] = ytd_prev_counts.reindex(all_reasons, fill_value=0).astype(int)
    summary["YTD_sel"]  = ytd_sel_counts.reindex(all_reasons, fill_value=0).astype(int)

    summary = summary.sort_values(["YTD_sel", "CM_sel"], ascending=False)
    if top_n is not None:
        summary = summary.head(int(top_n))

    cm_title_prev = pd.to_datetime(f"{prev_year}-{selected_month:02d}-01").strftime("%b-%y")
    cm_title_sel  = pd.to_datetime(f"{selected_year}-{selected_month:02d}-01").strftime("%b-%y")

    reasons = summary.index.tolist()
    wrapped_labels = [fill(r, width=18) for r in reasons]

    x = np.arange(len(reasons))
    w = 0.18

    fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)

    b1 = ax.bar(x - 1.5*w, summary["CM_prev"],  width=w, color="#0F68B9", label=f"CM {cm_title_prev}")
    b2 = ax.bar(x - 0.5*w, summary["CM_sel"],   width=w, color="#D8C37D", label=f"CM {cm_title_sel}")
    b3 = ax.bar(x + 0.5*w, summary["YTD_prev"], width=w, color="#006394", label=f"YTD {prev_year}")
    b4 = ax.bar(x + 1.5*w, summary["YTD_sel"],  width=w, color="#C1A02E", label=f"YTD {selected_year}")

    add_simple_value_labels(ax, b1, lambda v: f"{int(v)}", pad=1.0)
    add_simple_value_labels(ax, b2, lambda v: f"{int(v)}", pad=1.0)
    add_simple_value_labels(ax, b3, lambda v: f"{int(v)}", pad=1.0)
    add_simple_value_labels(ax, b4, lambda v: f"{int(v)}", pad=1.0)

    ax.set_xticks(x)
    ax.set_xticklabels(wrapped_labels, rotation=35, ha="right", fontsize=9)

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(False)

    ax.legend(
        loc="upper center",
        bbox_to_anchor=(0.5, -0.22),   # lower legend
        ncol=4,
        frameon=False,
        fontsize=10
    )

    fig.subplots_adjust(bottom=0.35)   # extra space for legend

    return fig

def slide_s3_service_reason_current_month_fig(selected_year, selected_month, top_n=10):
    base = df[(df["Is_Valid"] == True) & (df["Complaint_Category"] == "Service")].copy()
    base["Svc_Reason"] = base["Reason"].astype(str).str.strip()

    cm = base[(base["Year"] == selected_year) & (base["Month"] == selected_month)]
    counts = cm["Svc_Reason"].value_counts()
    if top_n is not None:
        counts = counts.head(int(top_n))

    fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)
    x = np.arange(len(counts))
    bars = ax.bar(x, counts.values, color="#C1A02E", width=0.6)

    add_simple_value_labels(ax, bars, lambda v: f"{int(v)}", pad=0.6)

    labels = [fill(str(s), width=22) for s in counts.index.tolist()]
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=10, rotation=30, ha="right")

    ax.set_ylim(0, max(1, counts.max()) * 1.15)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(False)

    plt.tight_layout()
    return fig

# ===== ISSUED (subset you included) =====
def slide_issued_1_donuts_fig(year, month):
    required = ["Year", "Month", "Complaint_Category"]
    missing = [c for c in required if c not in df_issued.columns]
    if missing:
        raise KeyError(f"df_issued missing columns: {missing}")

    d = df_issued.copy()
    df_month = d[(d["Year"] == year) & (d["Month"] == month)]
    df_ytd   = d[(d["Year"] == year) & (d["Month"].between(1, month))]

    month_counts = df_month["Complaint_Category"].value_counts()
    ytd_counts   = df_ytd["Complaint_Category"].value_counts()

    fig, axes = plt.subplots(1, 2, figsize=(13.33, 7.5), dpi=300)

    donut_chart(axes[0], month_counts)
    donut_chart(axes[1], ytd_counts)

    m_lbl = pd.to_datetime(f"{year}-{month:02d}-01").strftime("%b-%y")
    axes[0].text(0.5, -0.10, f"ISSUED – {m_lbl}", transform=axes[0].transAxes, ha="center", va="top", fontsize=13, color="#4D4D4D")
    axes[1].text(0.5, -0.10, f"ISSUED – {year} YTD", transform=axes[1].transAxes, ha="center", va="top", fontsize=13, color="#4D4D4D")

    fig.legend(["Service", "Quality", "Invalid"], loc="lower center", ncol=3, frameon=False, prop={"size": 14})
    plt.tight_layout(rect=[0, 0.08, 1, 1])
    return fig

def slide_issued_valid_count_fig(selected_year, selected_month):
    prev_year = selected_year - 1
    month_labels_all = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    months = list(range(1, selected_month + 1))
    month_labels = month_labels_all[:selected_month]
    cats = month_labels + ["YTD"]

    base = df_issued[df_issued["Is_Valid"] == True].copy()

    def monthly_counts(year):
        return (
            base[base["Year"] == year]
            .groupby("Month")
            .size()
            .reindex(months, fill_value=0)
        )

    def ytd_count(year):
        return int(base[(base["Year"] == year) & (base["Month"].between(1, selected_month))].shape[0])

    c_prev = monthly_counts(prev_year)
    c_sel  = monthly_counts(selected_year)
    ytd_prev = ytd_count(prev_year)
    ytd_sel  = ytd_count(selected_year)

    prev_vals = list(c_prev.values) + [ytd_prev]
    sel_vals  = list(c_sel.values)  + [ytd_sel]

    x = np.arange(len(cats))
    w = 0.28

    fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)
    b1 = ax.bar(x - w/2, prev_vals, width=w, label=str(prev_year), color="#006394")
    b2 = ax.bar(x + w/2, sel_vals,  width=w, label=str(selected_year), color="#C1A02E")

    add_simple_value_labels(ax, b1, lambda v: f"{int(v)}", pad=1.0)
    add_simple_value_labels(ax, b2, lambda v: f"{int(v)}", pad=1.0)

    # trendline selected year months only
    y = np.array(list(c_sel.values), dtype=float)
    x_months = np.arange(len(months), dtype=float)
    if len(months) >= 2 and np.any(y > 0):
        coeff = np.polyfit(x_months, y, 1)
        y_fit = np.polyval(coeff, x_months)
        ax.plot(x[:len(months)], y_fit, linestyle="--", linewidth=1.5, color="#4D4D4D", label="Trendline")

    ax.set_xticks(x)
    ax.set_xticklabels(cats)
    ax.set_ylabel("Count")

    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.08), ncol=3, frameon=False)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(False)

    plt.tight_layout(rect=[0, 0.05, 1, 1])
    return fig

def slide_issued_quality_reason_current_month_fig(selected_year, selected_month, top_n=12):
    base = df_issued[
        (df_issued["Is_Valid"] == True) &
        (df_issued["Complaint_Category"] == "Quality")
    ].copy()
    base["Reason_Q"] = base["Reason"].astype(str).str.strip()

    cm = base[(base["Year"] == selected_year) & (base["Month"] == selected_month)]
    counts = cm["Reason_Q"].value_counts()
    if top_n is not None:
        counts = counts.head(int(top_n))

    fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)
    x = np.arange(len(counts))
    bars = ax.bar(x, counts.values, color="#C1A02E", width=0.6)

    add_simple_value_labels(ax, bars, lambda v: f"{int(v)}", pad=0.6)

    labels = [fill(str(s), width=28) for s in counts.index.tolist()]
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=10, rotation=35, ha="right")

    ax.set_ylim(0, max(1, counts.max()) * 1.15)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(False)

    plt.tight_layout()
    return fig

def slide_issued_service_reason_current_month_fig(selected_year, selected_month, top_n=10):
    base = df_issued[
        (df_issued["Is_Valid"] == True) &
        (df_issued["Complaint_Category"] == "Service")
    ].copy()
    base["Reason_S"] = base["Reason"].astype(str).str.strip()

    cm = base[(base["Year"] == selected_year) & (base["Month"] == selected_month)]
    counts = cm["Reason_S"].value_counts()
    if top_n is not None:
        counts = counts.head(int(top_n))

    fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)
    x = np.arange(len(counts))
    bars = ax.bar(x, counts.values, color="#C1A02E", width=0.6)

    add_simple_value_labels(ax, bars, lambda v: f"{int(v)}", pad=0.6)

    labels = [fill(str(s), width=28) for s in counts.index.tolist()]
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=10, rotation=35, ha="right")

    ax.set_ylim(0, max(1, counts.max()) * 1.15)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(False)

    plt.tight_layout()
    return fig

# ===== NCR (subset you included) =====
def slide_ncr_quality_defect_cm_vs_ytd_fig(selected_year, selected_month, top_n=10):
    base = df_ncr_dash[
        (df_ncr_dash["Is_Valid"] == True) &
        (df_ncr_dash["Complaint_Category"] == "Quality")
    ].copy()

    base["Defect"] = base["Reason"].astype(str).str.strip()

    cm  = base[(base["Year"] == selected_year) & (base["Month"] == selected_month)]
    ytd = base[(base["Year"] == selected_year) & (base["Month"].between(1, selected_month))]

    cm_counts  = cm["Defect"].value_counts()
    ytd_counts = ytd["Defect"].value_counts()

    all_defects = pd.Index(ytd_counts.index).union(cm_counts.index)
    summary = pd.DataFrame(index=all_defects)
    summary["CM"]  = cm_counts.reindex(all_defects, fill_value=0).astype(int)
    summary["YTD"] = ytd_counts.reindex(all_defects, fill_value=0).astype(int)

    summary = summary.sort_values(["YTD", "CM"], ascending=False)
    if top_n is not None:
        summary = summary.head(int(top_n))

    defects = summary.index.tolist()
    wrapped = [fill(str(d), width=18) for d in defects]

    x = np.arange(len(defects))
    w = 0.35

    fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)

    b1 = ax.bar(x - w/2, summary["CM"],  width=w, color="#006394", label="Current Month")
    b2 = ax.bar(x + w/2, summary["YTD"], width=w, color="#C1A02E", label="YTD")

    add_simple_value_labels(ax, b1, lambda v: f"{int(v)}", pad=1.0)
    add_simple_value_labels(ax, b2, lambda v: f"{int(v)}", pad=1.0)

    ax.set_xticks(x)
    ax.set_xticklabels(wrapped, rotation=35, ha="right", fontsize=10)

    ax.set_ylim(0, max(1, summary[["CM","YTD"]].to_numpy().max()) * 1.18)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(False)

    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.08), ncol=2, frameon=False, fontsize=12)
    fig.subplots_adjust(bottom=0.22)
    return fig

def table_ncr_valid_count_by_month(selected_year, selected_month):
    base = df_ncr_dash[
        (df_ncr_dash["Year"] == selected_year) &
        (df_ncr_dash["Month"].between(1, selected_month)) &
        (df_ncr_dash["Is_Valid"] == True)
    ].copy()

    months = list(range(1, selected_month + 1))
    month_labels_all = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    labels = month_labels_all[:selected_month]

    monthly = (
        base.groupby("Month")
            .size()
            .reindex(months, fill_value=0)
            .astype(int)
    )
    ytd = int(monthly.sum())

    row = {lab: int(monthly[i]) for i, lab in zip(months, labels)}
    row["YTD"] = ytd

    df_out = pd.DataFrame([row], index=[str(selected_year)])
    return df_out

# (Your long root-cause slides are kept out of the PDF in this rewrite to keep it stable.
# If you want them included, paste them back as *_fig functions that RETURN fig, same style.)

def slide_ncr_vs_crm_correlation_ytd_fig(selected_year, selected_month, top_n=10, wrap_width=28):
    from matplotlib.ticker import FuncFormatter

    required = ["Year", "Month", "Is_Valid", "Complaint_Category", "Reason"]
    for name, dfx in [("df_final", df_final), ("df_ncr", df_ncr)]:
        missing = [c for c in required if c not in dfx.columns]
        if missing:
            raise KeyError(f"{name} missing columns: {missing}")

    crm = df_final[
        (df_final["Is_Valid"] == True) &
        (df_final["Complaint_Category"] == "Quality") &
        (df_final["Year"] == selected_year) &
        (df_final["Month"] <= selected_month)
    ].copy()
    crm["Reason"] = crm["Reason"].astype(str).str.strip()
    crm_counts = crm["Reason"].value_counts()

    ncr = df_ncr[
        (df_ncr["Is_Valid"] == True) &
        (df_ncr["Complaint_Category"] == "Quality") &
        (df_ncr["Year"] == selected_year) &
        (df_ncr["Month"] <= selected_month)
    ].copy()
    ncr["Reason"] = ncr["Reason"].astype(str).str.strip()
    ncr_counts = ncr["Reason"].value_counts()

    all_reasons = sorted(set(crm_counts.index) | set(ncr_counts.index))
    df_cmp = pd.DataFrame({
        "Reason": all_reasons,
        "CRM": [int(crm_counts.get(r, 0)) for r in all_reasons],
        "NCR": [int(ncr_counts.get(r, 0)) for r in all_reasons],
    })

    df_cmp = df_cmp.sort_values(["CRM", "NCR", "Reason"], ascending=[False, False, True]).head(int(top_n))

    reasons  = df_cmp["Reason"].tolist()
    crm_vals = df_cmp["CRM"].to_numpy(dtype=float)
    ncr_vals = df_cmp["NCR"].to_numpy(dtype=float)
    ncr_left = -ncr_vals

    y = np.arange(len(reasons))
    fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)

    c_ncr = "#D8C37D"
    c_crm = "#C1A02E"

    bar_h = 0.48
    ax.barh(y, ncr_left, color=c_ncr, height=bar_h, label="NCR")
    ax.barh(y, crm_vals, color=c_crm, height=bar_h, label="CRM")

    ax.axvline(0, color="#666666", linewidth=1)

    wrapped = [fill(r, wrap_width) for r in reasons]
    ax.set_yticks(y)
    ax.set_yticklabels(wrapped, fontsize=11)
    ax.invert_yaxis()

    max_side = float(max(crm_vals.max(), ncr_vals.max(), 1.0))
    ax.set_xlim(-max_side * 1.25, max_side * 1.25)
    ax.xaxis.set_major_formatter(FuncFormatter(lambda v, pos: f"{int(abs(v))}"))

    pad = max_side * 0.02
    for i, (ncr_v, crm_v) in enumerate(zip(ncr_vals, crm_vals)):
        if ncr_v > 0:
            ax.text(-ncr_v - pad, i, f"{int(ncr_v)}", va="center", ha="right", fontsize=11, color="#333333")
        if crm_v > 0:
            ax.text(crm_v + pad, i, f"{int(crm_v)}", va="center", ha="left", fontsize=11, color="#333333")

    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.08), ncol=2, frameon=False, fontsize=12)

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.tick_params(axis="x", labelsize=11)
    ax.grid(False)

    fig.subplots_adjust(left=0.28, right=0.98, top=0.90, bottom=0.18)
    return fig

def slide_valid_quality_by_customer_top5_cm_vs_ytd_fig(selected_year, selected_month, top_n=5):
    base = df[
        (df["Is_Valid"] == True) &
        (df["Complaint_Category"] == "Quality")
    ].copy()

    CUST_COL = "Customer"
    if CUST_COL not in base.columns:
        cand = [c for c in base.columns if "customer" in str(c).lower()]
        raise KeyError(f"Customer column '{CUST_COL}' not found. Similar columns: {cand}")

    base[CUST_COL] = base[CUST_COL].astype(str).str.strip()

    cm  = base[(base["Year"] == selected_year) & (base["Month"] == selected_month)]
    ytd = base[(base["Year"] == selected_year) & (base["Month"].between(1, selected_month))]

    cm_counts  = cm[CUST_COL].value_counts()
    ytd_counts = ytd[CUST_COL].value_counts()

    all_cust = pd.Index(ytd_counts.index).union(cm_counts.index)
    summary = pd.DataFrame(index=all_cust)
    summary["CM"]  = cm_counts.reindex(all_cust, fill_value=0).astype(int)
    summary["YTD"] = ytd_counts.reindex(all_cust, fill_value=0).astype(int)

    summary = summary.sort_values(["YTD", "CM"], ascending=False).head(int(top_n))

    customers = summary.index.tolist()
    wrapped = [fill(str(c), width=22) for c in customers]

    x = np.arange(len(customers))
    w = 0.35

    fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)

    b1 = ax.bar(x - w/2, summary["CM"],  width=w, color="#006394", label="Current Month")
    b2 = ax.bar(x + w/2, summary["YTD"], width=w, color="#C1A02E", label="YTD")

    add_simple_value_labels(ax, b1, lambda v: f"{int(v)}", pad=1.0)
    add_simple_value_labels(ax, b2, lambda v: f"{int(v)}", pad=1.0)

    ax.set_xticks(x)
    ax.set_xticklabels(wrapped, rotation=0, ha="center", fontsize=12)
    ax.set_ylabel("Count")

    ax.set_ylim(0, max(1, summary[["CM","YTD"]].to_numpy().max()) * 1.18)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(False)

    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.08), ncol=2, frameon=False, fontsize=12)
    fig.subplots_adjust(bottom=0.20)
    return fig

MONTHS_ORDER = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

def month_label(year: int, month: int) -> str:
    # month: 1..12
    return f"{MONTHS_ORDER[month-1]}-{str(year)[-2:]}"

def parse_fg_table_tsv(tsv_text: str) -> pd.DataFrame:
    """
    Expects columns: Month, FG_Invoiced
    """
    df = pd.read_csv(pd.io.common.StringIO(tsv_text.strip()), sep="\t")
    df.columns = [c.strip() for c in df.columns]
    if "Month" not in df.columns or "FG_Invoiced" not in df.columns:
        raise ValueError("FG table must have two columns: Month and FG_Invoiced (tab-separated).")
    df["Month"] = df["Month"].astype(str).str.strip()
    df["FG_Invoiced"] = pd.to_numeric(df["FG_Invoiced"], errors="coerce")
    return df

def build_cc_counts_by_month(df_final, selected_year: int,
                             year_col="Year", month_col="Month",
                             valid_col="Is_Valid", category_col="Complaint_Category") -> pd.DataFrame:
    """
    Returns columns: MonthLabel, Q_CC, S_CC, T_CC
    Assumes month_col is 1..12 or month names; adjust if needed.
    """
    base = df_final[
        (df_final[valid_col] == True) &
        (df_final[category_col].isin(["Quality", "Service"])) &
        (df_final[year_col] == selected_year)
    ].copy()

    # Normalize month to int 1..12 if it's not already
    if not np.issubdtype(base[month_col].dtype, np.number):
        # If you have "Jan-25" or "Jan" etc, you’ll need a mapping. Otherwise remove this.
        month_map = {m:i+1 for i,m in enumerate(MONTHS_ORDER)}
        base[month_col] = base[month_col].astype(str).str[:3].map(month_map)

    base["MonthNum"] = base[month_col].astype(int)

    q = (base[base[category_col] == "Quality"]
         .groupby("MonthNum").size().rename("Q_CC"))
    s = (base[base[category_col] == "Service"]
         .groupby("MonthNum").size().rename("S_CC"))

    out = pd.concat([q, s], axis=1).fillna(0).reset_index()
    out["T_CC"] = out["Q_CC"] + out["S_CC"]
    out["MonthLabel"] = out["MonthNum"].apply(lambda m: month_label(selected_year, m))
    out = out.sort_values("MonthNum")
    return out[["MonthNum", "MonthLabel", "Q_CC", "S_CC", "T_CC"]]

def make_fg_input_template(month_labels: list[str]) -> str:
    lines = ["Month\tFG_Invoiced"]
    lines += [f"{ml}\t" for ml in month_labels]
    return "\n".join(lines)


# =========================
# 6) Run charts + Export PDF
# =========================
if run:
    st.subheader(f"Period: {selected_year}-{selected_month:02d}")




    # Collect slides for PDF
    slides_for_pdf = []

    def render_and_collect(title, fig):
        st.subheader(title)
        show_fig(fig)
        # IMPORTANT: recreate fig for PDF (because show_fig closes it)
        return

    # ========== FINAL ==========
    st.header("FINAL (CRM Approved)")

    fig1 = slide_1_final_donuts_fig(selected_year, selected_month)
    st.subheader("Overview")
    show_fig(fig1)
    slides_for_pdf.append({"title": "FINAL - Total Complaints Issued", "fig": slide_1_final_donuts_fig(selected_year, selected_month)})

    fig2 = slide_2_leadtime_fig(selected_year, selected_month)
    st.subheader("Lead Time")
    show_fig(fig2)
    slides_for_pdf.append({"title": "FINAL - Lead Time (Days)", "fig": slide_2_leadtime_fig(selected_year, selected_month)})

    st.subheader("Quality")
    fig3 = slide_3_valid_quality_count_fig(selected_year, selected_month)
    show_fig(fig3)
    slides_for_pdf.append({"title": "FINAL - Valid Quality Count", "fig": slide_3_valid_quality_count_fig(selected_year, selected_month)})

    fig4 = slide_4_quality_defect_cm_vs_ytd_fig(selected_year, selected_month, top_n=TOPN_QUALITY_DEFECT)
    show_fig(fig4)
    slides_for_pdf.append({"title": "FINAL - Quality Defects (CM vs YTD)", "fig": slide_4_quality_defect_cm_vs_ytd_fig(selected_year, selected_month, top_n=TOPN_QUALITY_DEFECT)})

    fig5 = slide_5_quality_defect_current_month_fig(selected_year, selected_month, top_n=TOPN_QUALITY_CM)
    show_fig(fig5)
    slides_for_pdf.append({"title": "FINAL - Quality Defects (Current Month)", "fig": slide_5_quality_defect_current_month_fig(selected_year, selected_month, top_n=TOPN_QUALITY_DEFECT)})

    fig6 = slide_6_quality_cost_cm_vs_ytd_fig(selected_year, selected_month, top_n=TOPN_COST_DEFECT)
    show_fig(fig6)
    slides_for_pdf.append({"title": "FINAL - Quality Cost (Credit Note)", "fig": slide_6_quality_cost_cm_vs_ytd_fig(selected_year, selected_month, top_n=TOPN_QUALITY_DEFECT)})

    st.subheader("Service")
    figS1 = slide_s1_valid_service_count_fig(selected_year, selected_month)
    show_fig(figS1)
    slides_for_pdf.append({"title": "FINAL - Valid Service Count", "fig": slide_s1_valid_service_count_fig(selected_year, selected_month)})

    figS2 = slide_s2_service_reason_cm_vs_ytd_fig(selected_year, selected_month, top_n=TOPN_SERVICE_REASON)
    show_fig(figS2)
    slides_for_pdf.append({"title": "FINAL - Service Reasons (CM vs YTD)", "fig": slide_s2_service_reason_cm_vs_ytd_fig(selected_year, selected_month, top_n=TOPN_QUALITY_DEFECT)})

    figS3 = slide_s3_service_reason_current_month_fig(selected_year, selected_month, top_n=TOPN_SERVICE_CM)
    show_fig(figS3)
    slides_for_pdf.append({"title": "FINAL - Service Reasons (Current Month)", "fig": slide_s3_service_reason_current_month_fig(selected_year, selected_month, top_n=TOPN_QUALITY_DEFECT)})

    st.divider()

    # ========== ISSUED ==========
    st.header("ISSUED")

    figI1 = slide_issued_1_donuts_fig(selected_year, selected_month)
    st.subheader("Overview")
    show_fig(figI1)
    slides_for_pdf.append({"title": "ISSUED - Total Complaints Issued", "fig": slide_issued_1_donuts_fig(selected_year, selected_month)})

    st.subheader("Valid Count")
    figI2 = slide_issued_valid_count_fig(selected_year, selected_month)
    show_fig(figI2)
    slides_for_pdf.append({"title": "ISSUED - Valid Count", "fig": slide_issued_valid_count_fig(selected_year, selected_month)})

    st.subheader("Quality (Current Month)")
    figIQ = slide_issued_quality_reason_current_month_fig(selected_year, selected_month, top_n=TOPN_QUALITY_CM)
    show_fig(figIQ)
    slides_for_pdf.append({"title": "ISSUED - Quality Reasons (Current Month)", "fig": slide_issued_quality_reason_current_month_fig(selected_year, selected_month, top_n=TOPN_QUALITY_DEFECT)})

    st.subheader("Service (Current Month)")
    figIS = slide_issued_service_reason_current_month_fig(selected_year, selected_month, top_n=TOPN_SERVICE_CM)
    show_fig(figIS)
    slides_for_pdf.append({"title": "ISSUED - Service Reasons (Current Month)", "fig": slide_issued_service_reason_current_month_fig(selected_year, selected_month, top_n=TOPN_QUALITY_DEFECT)})

    st.divider()

    # ========== NCR ==========
    st.header("NCR")

    st.subheader("Quality (CM vs YTD)")
    figN1 = slide_ncr_quality_defect_cm_vs_ytd_fig(selected_year, selected_month, top_n=TOPN_NCR_DEFECT)

    show_fig(figN1)
    slides_for_pdf.append({"title": "NCR - Quality Defects (CM vs YTD)", "fig": slide_ncr_quality_defect_cm_vs_ytd_fig(selected_year, selected_month, top_n=TOPN_QUALITY_DEFECT)})

    st.subheader("Valid NCR Count Table")
    df_tbl = table_ncr_valid_count_by_month(selected_year, selected_month)
    st.dataframe(df_tbl, use_container_width=True)

    st.subheader("NCR vs CRM Correlation (YTD)")
    figCorr = slide_ncr_vs_crm_correlation_ytd_fig(selected_year, selected_month, top_n=TOPN_CORRELATION)

    show_fig(figCorr)
    slides_for_pdf.append({"title": "NCR vs CRM - Correlation (YTD)", "fig": slide_ncr_vs_crm_correlation_ytd_fig(selected_year, selected_month, top_n=TOPN_QUALITY_DEFECT)})

    st.subheader("Top Customers (FINAL)")
    figCust = slide_valid_quality_by_customer_top5_cm_vs_ytd_fig(selected_year, selected_month, top_n=TOPN_CUSTOMER)
    show_fig(figCust)
    slides_for_pdf.append({"title": "FINAL - Top Customers (CM vs YTD)", "fig": slide_valid_quality_by_customer_top5_cm_vs_ytd_fig(selected_year, selected_month, top_n=TOPN_QUALITY_DEFECT)})

    st.divider()

                # ── PDF Export ──
                st.divider()
                st.subheader("📥 Export")
                pdf_buf = build_ppt_pdf(slides_for_pdf)
                st.download_button(
                    label="📥 Download PPT-style PDF",
                    data=pdf_buf,
                    file_name=f"QM_Dashboard_{selected_year}-{selected_month:02d}.pdf",
                    mime="application/pdf"
                )

        except Exception as e:
            st.error(f"Error loading files: {str(e)}")

# ════════════════════════════════════════
# TAB 3 — DOCUMENTS
# ════════════════════════════════════════
with tab3:
    st.markdown("### 📁 Required Documents")
    try:
        req_docs    = supabase.table("qm_required_docs").select("*").order("created_at").execute()
        all_uploads = supabase.table("qm_documents").select("*").order("created_at", desc=True).execute()
    except Exception as e:
        st.error(f"Could not load documents: {str(e)}"); st.stop()

    if can_edit:
        with st.expander("➕ Add / Remove Required Document", expanded=False):
            col1, col2 = st.columns([3,1])
            with col1:
                new_doc_name = st.text_input("Document Name", key="qm_new_doc_name")
                new_doc_desc = st.text_input("Description (optional)", key="qm_new_doc_desc")
            with col2:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("➕ Add", use_container_width=True, key="qm_add_req"):
                    if not new_doc_name: st.error("Please enter a document name.")
                    else:
                        supabase.table("qm_required_docs").insert({"doc_name":new_doc_name,"description":new_doc_desc or "","added_by":name}).execute()
                        st.success(f"✅ '{new_doc_name}' added!"); st.rerun()
            if req_docs.data:
                st.divider()
                del_map={r['doc_name']:r['id'] for r in req_docs.data}
                col1,col2=st.columns([3,1])
                with col1: sel_del=st.selectbox("Select to remove",list(del_map.keys()),key="qm_del_req")
                with col2:
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("🗑️ Remove",use_container_width=True,key="qm_btn_del_req"):
                        supabase.table("qm_required_docs").delete().eq("id",del_map[sel_del]).execute()
                        st.success(f"Removed '{sel_del}'"); st.rerun()

    st.divider()
    if not req_docs.data:
        st.info("No required documents defined yet.")
    else:
        for req in req_docs.data:
            doc_name=req['doc_name']; doc_desc=req.get('description','')
            uploads=[u for u in (all_uploads.data or []) if u['doc_type']==doc_name]
            has_up=len(uploads)>0
            st.markdown(f"""
                <div style="border:1px solid {'#2ea04355' if has_up else '#f8514955'};border-radius:10px;
                    padding:16px 20px;margin-bottom:14px;background:{'#2ea04308' if has_up else '#f8514908'};">
                    <div style="display:flex;align-items:center;gap:10px;">
                        <span style="font-size:18px">{'✅' if has_up else '❌'}</span>
                        <span style="font-weight:700;font-size:16px;">{doc_name}</span>
                        <span style="font-size:11px;color:{'#2ea043' if has_up else '#f85149'};margin-left:auto;">
                            {len(uploads)} file{'s' if len(uploads)!=1 else ''} uploaded</span>
                    </div>
                    {f'<div style="font-size:12px;color:#8B949E;margin-top:4px;">{doc_desc}</div>' if doc_desc else ''}
                </div>
            """, unsafe_allow_html=True)
            for upload in uploads:
                ext=upload['file_name'].split('.')[-1].lower()
                file_url=f"https://sjcwzbftzpfylwdqiknh.supabase.co/storage/v1/object/public/asset/{quote(upload['file_path'],safe='/')}"
                col1,col2,col3=st.columns([4,2,1])
                with col1:
                    if ext in ['png','jpg','jpeg']: st.image(file_url,width=200)
                    else: st.markdown(f"📄 [{upload['file_name']}]({file_url})")
                    st.caption(f"Uploaded by {upload['uploaded_by']} · {pd.to_datetime(upload['created_at']).strftime('%d %b %Y %H:%M')}")
                with col3:
                    if can_edit and st.button("🗑️",key=f"qm_del_file_{upload['id']}"):
                        supabase.table("qm_documents").delete().eq("id",upload['id']).execute()
                        st.success("Deleted!"); st.rerun()
            if can_edit:
                with st.expander(f"⬆️ Upload file for '{doc_name}'",expanded=False):
                    doc_file=st.file_uploader("Choose file",type=["xlsx","pdf","png","jpg","docx"],key=f"qm_upload_{req['id']}")
                    if doc_file and st.button("Upload",key=f"qm_btn_upload_{req['id']}",type="primary"):
                        with st.spinner("Uploading..."):
                            try:
                                file_bytes=doc_file.read()
                                file_path=f"QM/{datetime.now().strftime('%Y%m%d_%H%M%S')}_{doc_file.name}"
                                supabase.storage.from_("asset").upload(file_path,file_bytes)
                                supabase.table("qm_documents").insert({"doc_type":doc_name,"file_name":doc_file.name,"file_path":file_path,"uploaded_by":name}).execute()
                                st.success("✅ Uploaded!"); st.rerun()
                            except Exception as e: st.error(f"Upload failed: {str(e)}")
            st.markdown("---")

# ════════════════════════════════════════
# TAB 4 — ACTION PLANS
# ════════════════════════════════════════
with tab4:
    st.markdown("### 🎯 Action Plans")
    try:
        actions=supabase.table("qm_action_plans").select("*").order("created_at",desc=True).execute()
        if actions.data:
            act_df=pd.DataFrame(actions.data)
            act_df["due_date"]=pd.to_datetime(act_df["due_date"]).dt.strftime("%d %b %Y")
            st.dataframe(act_df[["action","owner","due_date","status","notes"]].rename(columns={"action":"Action","owner":"Owner","due_date":"Due Date","status":"Status","notes":"Notes"}),use_container_width=True)
            if can_edit:
                act_map={f"{r['action'][:40]} — {r['owner']}":r['id'] for r in actions.data}
                col1,col2=st.columns([3,1])
                with col1: sel_act=st.selectbox("Select action",list(act_map.keys()),key="qm_sel_act")
                with col2: new_act_s=st.selectbox("Status",["Open","In Progress","Done"],key="qm_act_status")
                if st.button("✅ Update Status",key="qm_update_act"):
                    supabase.table("qm_action_plans").update({"status":new_act_s}).eq("id",act_map[sel_act]).execute()
                    st.success("Updated!"); st.rerun()
                st.markdown("#### 🗑️ Delete")
                del_act_map={f"{r['action'][:40]} — {r['owner']}":r['id'] for r in actions.data}
                sel_del_act=st.selectbox("Select to delete",list(del_act_map.keys()),key="qm_del_act")
                if st.button("🗑️ Delete Action Plan",key="qm_btn_del_act"):
                    supabase.table("qm_action_plans").delete().eq("id",del_act_map[sel_del_act]).execute()
                    st.success("Deleted!"); st.rerun()
        else:
            st.info("No action plans yet.")
    except Exception as e:
        st.error(f"Could not load action plans: {str(e)}")
    if can_edit:
        st.divider()
        st.markdown("#### ➕ Add Action Plan")
        act_text=st.text_input("Action",key="qm_act_text")
        own_text=st.text_input("Owner",key="qm_own_text")
        due_d=st.date_input("Due Date",key="qm_due_date")
        stat_ap=st.selectbox("Status",["Open","In Progress","Done"],key="qm_new_act_status")
        notes_ap=st.text_area("Notes",key="qm_notes_ap")
        if st.button("➕ Add Action Plan",type="primary",key="qm_add_act"):
            if not act_text or not own_text: st.error("Please fill in action and owner.")
            else:
                try:
                    supabase.table("qm_action_plans").insert({"action":act_text,"owner":own_text,"due_date":str(due_d),"status":stat_ap,"notes":notes_ap}).execute()
                    st.success("✅ Action plan added!"); st.rerun()
                except Exception as e: st.error(f"Error: {str(e)}")
