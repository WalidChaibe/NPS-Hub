# pages/4_QM.py - Quality Maintenance Pillar
import os, sys, io
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
os.environ["MPLCONFIGDIR"] = os.path.join(os.getcwd(), ".mplconfig")

from qm_pipeline import (
    read_excel_from_upload, build_dataset_final_issued, build_dataset_ncr,
    FINAL_APPROVAL_COL, CREATION_DATETIME_COL,
)
from utils.supabase_client import get_supabase

st.set_page_config(page_title="QM Pillar", page_icon="✅", layout="wide")

if "user" not in st.session_state:
    st.switch_page("app.py")

supabase = get_supabase()
role     = st.session_state.get("role", "member")
pillar   = st.session_state.get("pillar", "ALL")
name     = st.session_state.get("name", "User")
can_edit = (role == "plant_manager") or (role == "pillar_leader" and pillar == "QM")

col1, col2 = st.columns([5, 1])
with col1:
    st.markdown("# ✅ Quality Maintenance")
    st.markdown(f"Logged in as **{name}** · `{role}` · {'✏️ Full Access' if can_edit else '👁️ View Only'}")
with col2:
    if st.button("🏠 Home", use_container_width=True):
        st.switch_page("pages/1_Home.py")

st.divider()

tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Maturity Level", "📈 Quality Indicators", "📁 Documents", "🎯 Action Plans",
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
                      "QA Matrix tool developed", "KPIs monitored and suggestions for improvement available",
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
            if is_current:
                st.success("✅ Current Level")
            for criterion in level["criteria"]:
                st.checkbox(criterion, value=is_current, key=f"qm_lvl_{level['name']}_{criterion[:30]}")

# ════════════════════════════════════════
# TAB 2 — QUALITY INDICATORS
# ════════════════════════════════════════
with tab2:
    st.markdown("### 📈 CRM Quality Dashboard")

    mpl.rcParams["font.family"] = "DejaVu Sans"
    mpl.rcParams["font.size"] = 11
    mpl.rcParams["figure.dpi"] = 160
    mpl.rcParams["savefig.dpi"] = 320
    mpl.rcParams["savefig.bbox"] = "tight"
    COLORS = {"Service": "#C1A02E", "Quality": "#006394", "Invalid": "#D8C37D"}

    def fig_to_png_bytes(fig, dpi=300):
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight")
        buf.seek(0)
        return buf

    def show_fig(fig, dpi=300):
        st.image(fig_to_png_bytes(fig, dpi=dpi))
        plt.close(fig)

    def ppt_slide_title_bar(c, title, W, H):
        c.setFillColor(HexColor("#006394"))
        c.setFont("Helvetica-Bold", 26)
        c.drawString(40, H - 58, title)
        c.setStrokeColor(HexColor("#006394"))
        c.setLineWidth(4)
        c.line(40, H - 78, W - 40, H - 78)

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
        c.save()
        pdf_buf.seek(0)
        return pdf_buf

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
                df_final["Year"]   = df_final["Year"].astype(int)
                df_final["Month"]  = df_final["Month"].astype(int)
                df_issued["Year"]  = df_issued["Year"].astype(int)
                df_issued["Month"] = df_issued["Month"].astype(int)
                df_ncr["Year"]     = df_ncr["Year"].astype(int)
                df_ncr["Month"]    = df_ncr["Month"].astype(int)

                df_final_raw_flagged = final_pkg["raw_flagged"]
                df            = df_final
                df_raw_flagged= df_final_raw_flagged
                df_ncr_dash   = df_ncr

            st.success("✅ Files loaded!")

            # Period selector
            df_dates = df_final[["Year","Month"]].drop_duplicates().sort_values(["Year","Month"])
            years = sorted(df_dates["Year"].unique().tolist())
            col1, col2 = st.columns(2)
            with col1:
                selected_year = st.selectbox("Year", years, index=len(years)-1, key="qm_year")
            with col2:
                months_avail = sorted(df_dates[df_dates["Year"]==selected_year]["Month"].unique().tolist())
                selected_month = int(st.selectbox("Month", months_avail, index=len(months_avail)-1, key="qm_month"))

            # Top-N controls
            with st.expander("⚙️ Display Controls", expanded=False):
                c1,c2,c3,c4,c5,c6,c7,c8 = st.columns(8)
                def topn(col, label, default, max_n=40):
                    opts = ["All"]+list(range(3,max_n+1))
                    v = col.selectbox(label, opts, index=opts.index(default) if default in opts else 0)
                    return None if v=="All" else int(v)
                TOPN_QUALITY_DEFECT = topn(c1, "Q Defects CM vs YTD", 15)
                TOPN_QUALITY_CM     = topn(c2, "Q Defects CM", 10)
                TOPN_SERVICE_REASON = topn(c3, "Svc Reasons CM vs YTD", 15)
                TOPN_SERVICE_CM     = topn(c4, "Svc Reasons CM", 10)
                TOPN_COST_DEFECT    = topn(c5, "Cost Defects", 15)
                TOPN_NCR_DEFECT     = topn(c6, "NCR Defects", 20, 50)
                TOPN_CORRELATION    = topn(c7, "NCR vs CRM", 12)
                TOPN_CUSTOMER       = topn(c8, "Top Customers", 5, 20)

            run = st.button("🔄 Generate Charts", type="primary", key="qm_run")

            if run:
                prev_year    = selected_year - 1
                months_range = list(range(1, selected_month+1))
                MONTH_LABELS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
                slides_for_pdf = []

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

                # ── FINAL ──
                st.header("FINAL (CRM Approved)")
                st.subheader("Overview")

                def slide_1_final_donuts_fig(year, month):
                    fig,axes=plt.subplots(1,2,figsize=(13.33,7.5),dpi=300)
                    donut_chart(axes[0],df[(df["Year"]==year)&(df["Month"]==month)]["Complaint_Category"].value_counts())
                    donut_chart(axes[1],df[(df["Year"]==year)&(df["Month"].between(1,month))]["Complaint_Category"].value_counts())
                    m_lbl=pd.to_datetime(f"{year}-{month:02d}-01").strftime("%b-%y")
                    axes[0].text(0.5,-0.10,m_lbl,transform=axes[0].transAxes,ha="center",va="top",fontsize=13,color="#4D4D4D")
                    axes[1].text(0.5,-0.10,f"{year} YTD",transform=axes[1].transAxes,ha="center",va="top",fontsize=13,color="#4D4D4D")
                    fig.legend(["Service","Quality","Invalid"],loc="lower center",ncol=3,frameon=False,prop={"size":14})
                    plt.tight_layout(rect=[0,0.08,1,1]); return fig
                show_fig(slide_1_final_donuts_fig(selected_year,selected_month))
                slides_for_pdf.append({"title":"FINAL - Total Complaints Issued","fig":slide_1_final_donuts_fig(selected_year,selected_month)})

                # Lead Time
                st.subheader("Lead Time")
                def build_unique_crm_leadtime_table(df_in):
                    tmp=df_in.copy()
                    tmp["First_Approval_dt"]=pd.to_datetime(tmp["First Approval Date"],errors="coerce")
                    tmp["Final_Approval_dt"]=pd.to_datetime(tmp["Final Approval Date"],errors="coerce")
                    g=tmp.groupby("Ref NB",as_index=False).agg(First_Approval_dt=("First_Approval_dt","min"),Final_Approval_dt=("Final_Approval_dt","max"))
                    g["LeadTime_days"]=(g["Final_Approval_dt"]-g["First_Approval_dt"]).dt.total_seconds()/86400.0
                    g["Year"]=g["Final_Approval_dt"].dt.year; g["Month"]=g["Final_Approval_dt"].dt.month
                    g=g.dropna(subset=["Year","Month","LeadTime_days"]); g=g[g["LeadTime_days"]>=0]
                    return g
                df_lt=build_unique_crm_leadtime_table(df)

                def slide_2_leadtime_fig(selected_year,selected_month):
                    def mavg(yr): return df_lt[df_lt["Year"]==yr].groupby("Month")["LeadTime_days"].mean().reindex(months_range)
                    def ytda(yr):
                        d=df_lt[(df_lt["Year"]==yr)&(df_lt["Month"].between(1,selected_month))]
                        return float(d["LeadTime_days"].mean()) if len(d) else np.nan
                    cats=MONTH_LABELS[:selected_month]+["AVG-YTD"]; x=np.arange(len(cats)); w=0.28
                    fig,ax=plt.subplots(figsize=(13.33,7.5),dpi=300)
                    b1=ax.bar(x-w/2,list(mavg(prev_year).values)+[ytda(prev_year)],w,label=str(prev_year),color="#006394")
                    b2=ax.bar(x+w/2,list(mavg(selected_year).values)+[ytda(selected_year)],w,label=str(selected_year),color="#C1A02E")
                    add_simple_value_labels(ax,b1,lambda v:f"{v:.0f}",0.6)
                    add_simple_value_labels(ax,b2,lambda v:f"{v:.0f}",0.6)
                    ax.set_xticks(x); ax.set_xticklabels(cats); ax.set_ylabel("Lead Time (days)")
                    ax.legend(loc="upper center",bbox_to_anchor=(0.5,-0.08),ncol=2,frameon=False)
                    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.grid(False)
                    plt.tight_layout(rect=[0,0.05,1,1]); return fig
                show_fig(slide_2_leadtime_fig(selected_year,selected_month))
                slides_for_pdf.append({"title":"FINAL - Lead Time (Days)","fig":slide_2_leadtime_fig(selected_year,selected_month)})

                # Quality Count
                st.subheader("Quality")
                def slide_3_valid_quality_count_fig(selected_year,selected_month):
                    base=df[(df["Is_Valid"]==True)&(df["Complaint_Category"]=="Quality")]
                    def mc(yr): return base[base["Year"]==yr].groupby("Month").size().reindex(months_range,fill_value=0)
                    def ytdc(yr): return int(base[(base["Year"]==yr)&(base["Month"].between(1,selected_month))].shape[0])
                    cats=MONTH_LABELS[:selected_month]+["YTD"]; x=np.arange(len(cats)); w=0.28
                    fig,ax=plt.subplots(figsize=(13.33,7.5),dpi=300)
                    b1=ax.bar(x-w/2,list(mc(prev_year).values)+[ytdc(prev_year)],w,label=str(prev_year),color="#006394")
                    b2=ax.bar(x+w/2,list(mc(selected_year).values)+[ytdc(selected_year)],w,label=str(selected_year),color="#C1A02E")
                    add_simple_value_labels(ax,b1,lambda v:f"{int(v)}",1); add_simple_value_labels(ax,b2,lambda v:f"{int(v)}",1)
                    y=np.array(list(mc(selected_year).values),dtype=float); xm=np.arange(len(months_range),dtype=float)
                    if len(months_range)>=2 and np.any(y>0):
                        coeff=np.polyfit(xm,y,1)
                        ax.plot(x[:len(months_range)],np.polyval(coeff,xm),linestyle="--",linewidth=1.5,color="#C1A02E",label=f"Trend {selected_year}")
                    ax.set_xticks(x); ax.set_xticklabels(cats); ax.set_ylabel("Count")
                    ax.legend(loc="upper center",bbox_to_anchor=(0.5,-0.08),ncol=3,frameon=False)
                    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.grid(False)
                    plt.tight_layout(rect=[0,0.05,1,1]); return fig
                show_fig(slide_3_valid_quality_count_fig(selected_year,selected_month))
                slides_for_pdf.append({"title":"FINAL - Valid Quality Count","fig":slide_3_valid_quality_count_fig(selected_year,selected_month)})

                def slide_4_quality_defect_cm_vs_ytd_fig(selected_year,selected_month,top_n=15):
                    base=df[(df["Is_Valid"]==True)&(df["Complaint_Category"]=="Quality")].copy()
                    base["Defect"]=base["Reason"].astype(str).str.strip()
                    cp=base[(base["Year"]==prev_year)&(base["Month"]==selected_month)]["Defect"].value_counts()
                    cs=base[(base["Year"]==selected_year)&(base["Month"]==selected_month)]["Defect"].value_counts()
                    yp=base[(base["Year"]==prev_year)&(base["Month"].between(1,selected_month))]["Defect"].value_counts()
                    ys=base[(base["Year"]==selected_year)&(base["Month"].between(1,selected_month))]["Defect"].value_counts()
                    all_d=pd.Index(ys.index).union(yp.index).union(cs.index).union(cp.index)
                    summ=pd.DataFrame({"CM_prev":cp.reindex(all_d,fill_value=0),"CM_sel":cs.reindex(all_d,fill_value=0),"YTD_prev":yp.reindex(all_d,fill_value=0),"YTD_sel":ys.reindex(all_d,fill_value=0)})
                    summ=summ.sort_values(["YTD_sel","CM_sel"],ascending=False)
                    if top_n: summ=summ.head(int(top_n))
                    cmp=pd.to_datetime(f"{prev_year}-{selected_month:02d}-01").strftime("%b-%y")
                    cms=pd.to_datetime(f"{selected_year}-{selected_month:02d}-01").strftime("%b-%y")
                    x=np.arange(len(summ)); w=0.18
                    fig,ax=plt.subplots(figsize=(13.33,7.5),dpi=300)
                    b1=ax.bar(x-1.5*w,summ["CM_prev"],w,color="#0F68B9",label=f"CM {cmp}")
                    b2=ax.bar(x-0.5*w,summ["CM_sel"],w,color="#D8C37D",label=f"CM {cms}")
                    b3=ax.bar(x+0.5*w,summ["YTD_prev"],w,color="#006394",label=f"YTD {prev_year}")
                    b4=ax.bar(x+1.5*w,summ["YTD_sel"],w,color="#C1A02E",label=f"YTD {selected_year}")
                    for b in [b1,b2,b3,b4]: add_simple_value_labels(ax,b,lambda v:f"{int(v)}",1)
                    ax.set_xticks(x); ax.set_xticklabels([fill(d,16) for d in summ.index],rotation=35,ha="right",fontsize=9)
                    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.grid(False)
                    ax.legend(loc="upper center",bbox_to_anchor=(0.5,-0.22),ncol=4,frameon=False,fontsize=10)
                    fig.subplots_adjust(bottom=0.35); return fig
                st.subheader("Quality Defects (CM vs YTD)")
                show_fig(slide_4_quality_defect_cm_vs_ytd_fig(selected_year,selected_month,TOPN_QUALITY_DEFECT))
                slides_for_pdf.append({"title":"FINAL - Quality Defects (CM vs YTD)","fig":slide_4_quality_defect_cm_vs_ytd_fig(selected_year,selected_month,TOPN_QUALITY_DEFECT)})

                def slide_5_quality_defect_current_month_fig(selected_year,selected_month,top_n=10):
                    base=df[(df["Is_Valid"]==True)&(df["Complaint_Category"]=="Quality")].copy()
                    base["Defect"]=base["Reason"].astype(str).str.strip()
                    counts=base[(base["Year"]==selected_year)&(base["Month"]==selected_month)]["Defect"].value_counts()
                    if top_n: counts=counts.head(int(top_n))
                    fig,ax=plt.subplots(figsize=(13.33,7.5),dpi=300)
                    bars=ax.bar(np.arange(len(counts)),counts.values,color="#C1A02E",width=0.6)
                    add_simple_value_labels(ax,bars,lambda v:f"{int(v)}",0.6)
                    ax.set_xticks(np.arange(len(counts))); ax.set_xticklabels([fill(str(s),18) for s in counts.index],fontsize=10,rotation=25,ha="right")
                    ax.set_ylim(0,max(1,counts.max())*1.15)
                    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.grid(False)
                    plt.tight_layout(); return fig
                st.subheader("Quality Defects (Current Month)")
                show_fig(slide_5_quality_defect_current_month_fig(selected_year,selected_month,TOPN_QUALITY_CM))
                slides_for_pdf.append({"title":"FINAL - Quality Defects (Current Month)","fig":slide_5_quality_defect_current_month_fig(selected_year,selected_month,TOPN_QUALITY_CM)})

                def slide_6_quality_cost_cm_vs_ytd_fig(selected_year,selected_month,top_n=15):
                    decision_candidates=[c for c in df_raw_flagged.columns if "decision" in str(c).lower()]
                    if not decision_candidates: raise KeyError("Could not find Decision column.")
                    DECISION_COL=decision_candidates[0]
                    base=df_raw_flagged[(df_raw_flagged["Is_Valid"]==True)&(df_raw_flagged["Complaint_Category"]=="Quality")].copy()
                    base[DECISION_COL]=base[DECISION_COL].astype(str).str.strip()
                    base=base[base[DECISION_COL].str.lower()=="credit note"]
                    base["Defect"]=base["Reason"].astype(str).str.strip()
                    base["Cost Amount"]=pd.to_numeric(base["Cost Amount"],errors="coerce").fillna(0)
                    cp=base[(base["Year"]==prev_year)&(base["Month"]==selected_month)].groupby("Defect")["Cost Amount"].sum()
                    cs=base[(base["Year"]==selected_year)&(base["Month"]==selected_month)].groupby("Defect")["Cost Amount"].sum()
                    yp=base[(base["Year"]==prev_year)&(base["Month"].between(1,selected_month))].groupby("Defect")["Cost Amount"].sum()
                    ys=base[(base["Year"]==selected_year)&(base["Month"].between(1,selected_month))].groupby("Defect")["Cost Amount"].sum()
                    summ=pd.DataFrame({"CM_prev":cp,"CM_sel":cs,"YTD_prev":yp,"YTD_sel":ys}).fillna(0)
                    summ=summ.sort_values(["YTD_sel","CM_sel"],ascending=False)
                    if top_n: summ=summ.head(int(top_n))
                    x=np.arange(len(summ)); w=0.18
                    fig,ax=plt.subplots(figsize=(13.33,7.5),dpi=300)
                    b1=ax.bar(x-1.5*w,summ["CM_prev"],w,label=f"CM {prev_year}",color="#0F68B9")
                    b2=ax.bar(x-0.5*w,summ["CM_sel"],w,label=f"CM {selected_year}",color="#D8C37D")
                    b3=ax.bar(x+0.5*w,summ["YTD_prev"],w,label=f"YTD {prev_year}",color="#006394")
                    b4=ax.bar(x+1.5*w,summ["YTD_sel"],w,label=f"YTD {selected_year}",color="#C1A02E")
                    ymax=summ.to_numpy().max(); pad=ymax*0.015
                    def fmt(v): return f"SAR {int(v/1000)}K" if v>=1000 else f"SAR {int(v)}"
                    for bars_g in [b1,b2,b3,b4]:
                        for bar in bars_g:
                            h=bar.get_height()
                            if h>0: ax.text(bar.get_x()+bar.get_width()/2,h+pad,fmt(h),ha="center",va="bottom",fontsize=9,rotation=90,color="#333333",clip_on=False)
                    ax.set_xticks(x); ax.set_xticklabels([fill(d,18) for d in summ.index],rotation=30,ha="right",fontsize=9)
                    ax.set_ylabel("Cost Amount (SAR)")
                    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.grid(False)
                    ax.legend(loc="upper center",bbox_to_anchor=(0.5,-0.22),ncol=4,frameon=False,fontsize=10)
                    fig.subplots_adjust(bottom=0.35); return fig
                try:
                    st.subheader("Quality Cost (Credit Note)")
                    show_fig(slide_6_quality_cost_cm_vs_ytd_fig(selected_year,selected_month,TOPN_COST_DEFECT))
                    slides_for_pdf.append({"title":"FINAL - Quality Cost (Credit Note)","fig":slide_6_quality_cost_cm_vs_ytd_fig(selected_year,selected_month,TOPN_COST_DEFECT)})
                except Exception as e:
                    st.warning(f"Cost slide skipped: {e}")

                # Service
                st.subheader("Service")
                def slide_s1_valid_service_count_fig(selected_year,selected_month):
                    base=df[(df["Is_Valid"]==True)&(df["Complaint_Category"]=="Service")]
                    def mc(yr): return base[base["Year"]==yr].groupby("Month").size().reindex(months_range,fill_value=0)
                    def ytdc(yr): return int(base[(base["Year"]==yr)&(base["Month"].between(1,selected_month))].shape[0])
                    cats=MONTH_LABELS[:selected_month]+["YTD"]; x=np.arange(len(cats)); w=0.28
                    fig,ax=plt.subplots(figsize=(13.33,7.5),dpi=300)
                    b1=ax.bar(x-w/2,list(mc(prev_year).values)+[ytdc(prev_year)],w,label=str(prev_year),color="#006394")
                    b2=ax.bar(x+w/2,list(mc(selected_year).values)+[ytdc(selected_year)],w,label=str(selected_year),color="#C1A02E")
                    add_simple_value_labels(ax,b1,lambda v:f"{int(v)}",1); add_simple_value_labels(ax,b2,lambda v:f"{int(v)}",1)
                    y=np.array(list(mc(selected_year).values),dtype=float); xm=np.arange(len(months_range),dtype=float)
                    if len(months_range)>=2 and np.any(y>0):
                        coeff=np.polyfit(xm,y,1)
                        ax.plot(x[:len(months_range)],np.polyval(coeff,xm),linestyle="--",linewidth=1.5,color="#C1A02E",label=f"Trend {selected_year}")
                    ax.set_xticks(x); ax.set_xticklabels(cats); ax.set_ylabel("Count")
                    ax.legend(loc="upper center",bbox_to_anchor=(0.5,-0.08),ncol=3,frameon=False)
                    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.grid(False)
                    plt.tight_layout(rect=[0,0.05,1,1]); return fig
                st.subheader("Valid Service Count")
                show_fig(slide_s1_valid_service_count_fig(selected_year,selected_month))
                slides_for_pdf.append({"title":"FINAL - Valid Service Count","fig":slide_s1_valid_service_count_fig(selected_year,selected_month)})

                def slide_s2_service_reason_cm_vs_ytd_fig(selected_year,selected_month,top_n=15):
                    base=df[(df["Is_Valid"]==True)&(df["Complaint_Category"]=="Service")].copy()
                    base["Reason_Svc"]=base["Reason"].astype(str).str.strip()
                    cp=base[(base["Year"]==prev_year)&(base["Month"]==selected_month)]["Reason_Svc"].value_counts()
                    cs=base[(base["Year"]==selected_year)&(base["Month"]==selected_month)]["Reason_Svc"].value_counts()
                    yp=base[(base["Year"]==prev_year)&(base["Month"].between(1,selected_month))]["Reason_Svc"].value_counts()
                    ys=base[(base["Year"]==selected_year)&(base["Month"].between(1,selected_month))]["Reason_Svc"].value_counts()
                    all_r=pd.Index(ys.index).union(yp.index).union(cs.index).union(cp.index)
                    summ=pd.DataFrame({"CM_prev":cp.reindex(all_r,fill_value=0),"CM_sel":cs.reindex(all_r,fill_value=0),"YTD_prev":yp.reindex(all_r,fill_value=0),"YTD_sel":ys.reindex(all_r,fill_value=0)})
                    summ=summ.sort_values(["YTD_sel","CM_sel"],ascending=False)
                    if top_n: summ=summ.head(int(top_n))
                    cmp=pd.to_datetime(f"{prev_year}-{selected_month:02d}-01").strftime("%b-%y")
                    cms=pd.to_datetime(f"{selected_year}-{selected_month:02d}-01").strftime("%b-%y")
                    x=np.arange(len(summ)); w=0.18
                    fig,ax=plt.subplots(figsize=(13.33,7.5),dpi=300)
                    b1=ax.bar(x-1.5*w,summ["CM_prev"],w,color="#0F68B9",label=f"CM {cmp}")
                    b2=ax.bar(x-0.5*w,summ["CM_sel"],w,color="#D8C37D",label=f"CM {cms}")
                    b3=ax.bar(x+0.5*w,summ["YTD_prev"],w,color="#006394",label=f"YTD {prev_year}")
                    b4=ax.bar(x+1.5*w,summ["YTD_sel"],w,color="#C1A02E",label=f"YTD {selected_year}")
                    for b in [b1,b2,b3,b4]: add_simple_value_labels(ax,b,lambda v:f"{int(v)}",1)
                    ax.set_xticks(x); ax.set_xticklabels([fill(r,18) for r in summ.index],rotation=35,ha="right",fontsize=9)
                    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.grid(False)
                    ax.legend(loc="upper center",bbox_to_anchor=(0.5,-0.22),ncol=4,frameon=False,fontsize=10)
                    fig.subplots_adjust(bottom=0.35); return fig
                st.subheader("Service Reasons (CM vs YTD)")
                show_fig(slide_s2_service_reason_cm_vs_ytd_fig(selected_year,selected_month,TOPN_SERVICE_REASON))
                slides_for_pdf.append({"title":"FINAL - Service Reasons (CM vs YTD)","fig":slide_s2_service_reason_cm_vs_ytd_fig(selected_year,selected_month,TOPN_SERVICE_REASON)})

                def slide_s3_service_reason_current_month_fig(selected_year,selected_month,top_n=10):
                    base=df[(df["Is_Valid"]==True)&(df["Complaint_Category"]=="Service")].copy()
                    base["Svc_Reason"]=base["Reason"].astype(str).str.strip()
                    counts=base[(base["Year"]==selected_year)&(base["Month"]==selected_month)]["Svc_Reason"].value_counts()
                    if top_n: counts=counts.head(int(top_n))
                    fig,ax=plt.subplots(figsize=(13.33,7.5),dpi=300)
                    bars=ax.bar(np.arange(len(counts)),counts.values,color="#C1A02E",width=0.6)
                    add_simple_value_labels(ax,bars,lambda v:f"{int(v)}",0.6)
                    ax.set_xticks(np.arange(len(counts))); ax.set_xticklabels([fill(str(s),22) for s in counts.index],fontsize=10,rotation=30,ha="right")
                    ax.set_ylim(0,max(1,counts.max())*1.15)
                    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.grid(False)
                    plt.tight_layout(); return fig
                st.subheader("Service Reasons (Current Month)")
                show_fig(slide_s3_service_reason_current_month_fig(selected_year,selected_month,TOPN_SERVICE_CM))
                slides_for_pdf.append({"title":"FINAL - Service Reasons (Current Month)","fig":slide_s3_service_reason_current_month_fig(selected_year,selected_month,TOPN_SERVICE_CM)})

                st.divider()

                # ── ISSUED ──
                st.header("ISSUED")
                st.subheader("Overview")
                def slide_issued_1_donuts_fig(year,month):
                    fig,axes=plt.subplots(1,2,figsize=(13.33,7.5),dpi=300)
                    donut_chart(axes[0],df_issued[(df_issued["Year"]==year)&(df_issued["Month"]==month)]["Complaint_Category"].value_counts())
                    donut_chart(axes[1],df_issued[(df_issued["Year"]==year)&(df_issued["Month"].between(1,month))]["Complaint_Category"].value_counts())
                    m_lbl=pd.to_datetime(f"{year}-{month:02d}-01").strftime("%b-%y")
                    axes[0].text(0.5,-0.10,f"ISSUED – {m_lbl}",transform=axes[0].transAxes,ha="center",va="top",fontsize=13,color="#4D4D4D")
                    axes[1].text(0.5,-0.10,f"ISSUED – {year} YTD",transform=axes[1].transAxes,ha="center",va="top",fontsize=13,color="#4D4D4D")
                    fig.legend(["Service","Quality","Invalid"],loc="lower center",ncol=3,frameon=False,prop={"size":14})
                    plt.tight_layout(rect=[0,0.08,1,1]); return fig
                show_fig(slide_issued_1_donuts_fig(selected_year,selected_month))
                slides_for_pdf.append({"title":"ISSUED - Total Complaints","fig":slide_issued_1_donuts_fig(selected_year,selected_month)})

                st.subheader("Valid Count")
                def slide_issued_valid_count_fig(selected_year,selected_month):
                    base=df_issued[df_issued["Is_Valid"]==True]
                    def mc(yr): return base[base["Year"]==yr].groupby("Month").size().reindex(months_range,fill_value=0)
                    def ytdc(yr): return int(base[(base["Year"]==yr)&(base["Month"].between(1,selected_month))].shape[0])
                    cats=MONTH_LABELS[:selected_month]+["YTD"]; x=np.arange(len(cats)); w=0.28
                    fig,ax=plt.subplots(figsize=(13.33,7.5),dpi=300)
                    b1=ax.bar(x-w/2,list(mc(prev_year).values)+[ytdc(prev_year)],w,label=str(prev_year),color="#006394")
                    b2=ax.bar(x+w/2,list(mc(selected_year).values)+[ytdc(selected_year)],w,label=str(selected_year),color="#C1A02E")
                    add_simple_value_labels(ax,b1,lambda v:f"{int(v)}",1); add_simple_value_labels(ax,b2,lambda v:f"{int(v)}",1)
                    y=np.array(list(mc(selected_year).values),dtype=float); xm=np.arange(len(months_range),dtype=float)
                    if len(months_range)>=2 and np.any(y>0):
                        coeff=np.polyfit(xm,y,1)
                        ax.plot(x[:len(months_range)],np.polyval(coeff,xm),linestyle="--",linewidth=1.5,color="#4D4D4D",label="Trendline")
                    ax.set_xticks(x); ax.set_xticklabels(cats); ax.set_ylabel("Count")
                    ax.legend(loc="upper center",bbox_to_anchor=(0.5,-0.08),ncol=3,frameon=False)
                    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.grid(False)
                    plt.tight_layout(rect=[0,0.05,1,1]); return fig
                show_fig(slide_issued_valid_count_fig(selected_year,selected_month))
                slides_for_pdf.append({"title":"ISSUED - Valid Count","fig":slide_issued_valid_count_fig(selected_year,selected_month)})

                st.subheader("Quality (Current Month)")
                def slide_issued_quality_reason_current_month_fig(selected_year,selected_month,top_n=12):
                    base=df_issued[(df_issued["Is_Valid"]==True)&(df_issued["Complaint_Category"]=="Quality")].copy()
                    base["Reason_Q"]=base["Reason"].astype(str).str.strip()
                    counts=base[(base["Year"]==selected_year)&(base["Month"]==selected_month)]["Reason_Q"].value_counts()
                    if top_n: counts=counts.head(int(top_n))
                    fig,ax=plt.subplots(figsize=(13.33,7.5),dpi=300)
                    bars=ax.bar(np.arange(len(counts)),counts.values,color="#C1A02E",width=0.6)
                    add_simple_value_labels(ax,bars,lambda v:f"{int(v)}",0.6)
                    ax.set_xticks(np.arange(len(counts))); ax.set_xticklabels([fill(str(s),28) for s in counts.index],fontsize=10,rotation=35,ha="right")
                    ax.set_ylim(0,max(1,counts.max())*1.15)
                    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.grid(False)
                    plt.tight_layout(); return fig
                show_fig(slide_issued_quality_reason_current_month_fig(selected_year,selected_month,TOPN_QUALITY_CM))
                slides_for_pdf.append({"title":"ISSUED - Quality Reasons (CM)","fig":slide_issued_quality_reason_current_month_fig(selected_year,selected_month,TOPN_QUALITY_CM)})

                st.subheader("Service (Current Month)")
                def slide_issued_service_reason_current_month_fig(selected_year,selected_month,top_n=10):
                    base=df_issued[(df_issued["Is_Valid"]==True)&(df_issued["Complaint_Category"]=="Service")].copy()
                    base["Reason_S"]=base["Reason"].astype(str).str.strip()
                    counts=base[(base["Year"]==selected_year)&(base["Month"]==selected_month)]["Reason_S"].value_counts()
                    if top_n: counts=counts.head(int(top_n))
                    fig,ax=plt.subplots(figsize=(13.33,7.5),dpi=300)
                    bars=ax.bar(np.arange(len(counts)),counts.values,color="#C1A02E",width=0.6)
                    add_simple_value_labels(ax,bars,lambda v:f"{int(v)}",0.6)
                    ax.set_xticks(np.arange(len(counts))); ax.set_xticklabels([fill(str(s),28) for s in counts.index],fontsize=10,rotation=35,ha="right")
                    ax.set_ylim(0,max(1,counts.max())*1.15)
                    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.grid(False)
                    plt.tight_layout(); return fig
                show_fig(slide_issued_service_reason_current_month_fig(selected_year,selected_month,TOPN_SERVICE_CM))
                slides_for_pdf.append({"title":"ISSUED - Service Reasons (CM)","fig":slide_issued_service_reason_current_month_fig(selected_year,selected_month,TOPN_SERVICE_CM)})

                st.divider()

                # ── NCR ──
                st.header("NCR")
                st.subheader("Quality (CM vs YTD)")
                def slide_ncr_quality_defect_cm_vs_ytd_fig(selected_year,selected_month,top_n=20):
                    base=df_ncr_dash[(df_ncr_dash["Is_Valid"]==True)&(df_ncr_dash["Complaint_Category"]=="Quality")].copy()
                    base["Defect"]=base["Reason"].astype(str).str.strip()
                    cm=base[(base["Year"]==selected_year)&(base["Month"]==selected_month)]["Defect"].value_counts()
                    ytd=base[(base["Year"]==selected_year)&(base["Month"].between(1,selected_month))]["Defect"].value_counts()
                    all_d=pd.Index(ytd.index).union(cm.index)
                    summ=pd.DataFrame({"CM":cm.reindex(all_d,fill_value=0),"YTD":ytd.reindex(all_d,fill_value=0)})
                    summ=summ.sort_values(["YTD","CM"],ascending=False)
                    if top_n: summ=summ.head(int(top_n))
                    x=np.arange(len(summ)); w=0.35
                    fig,ax=plt.subplots(figsize=(13.33,7.5),dpi=300)
                    b1=ax.bar(x-w/2,summ["CM"],w,color="#006394",label="Current Month")
                    b2=ax.bar(x+w/2,summ["YTD"],w,color="#C1A02E",label="YTD")
                    add_simple_value_labels(ax,b1,lambda v:f"{int(v)}",1); add_simple_value_labels(ax,b2,lambda v:f"{int(v)}",1)
                    ax.set_xticks(x); ax.set_xticklabels([fill(str(d),18) for d in summ.index],rotation=35,ha="right",fontsize=10)
                    ax.set_ylim(0,max(1,summ[["CM","YTD"]].to_numpy().max())*1.18)
                    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.grid(False)
                    ax.legend(loc="upper center",bbox_to_anchor=(0.5,-0.08),ncol=2,frameon=False,fontsize=12)
                    fig.subplots_adjust(bottom=0.22); return fig
                show_fig(slide_ncr_quality_defect_cm_vs_ytd_fig(selected_year,selected_month,TOPN_NCR_DEFECT))
                slides_for_pdf.append({"title":"NCR - Quality Defects (CM vs YTD)","fig":slide_ncr_quality_defect_cm_vs_ytd_fig(selected_year,selected_month,TOPN_NCR_DEFECT)})

                st.subheader("Valid NCR Count Table")
                base_tbl=df_ncr_dash[(df_ncr_dash["Year"]==selected_year)&(df_ncr_dash["Month"].between(1,selected_month))&(df_ncr_dash["Is_Valid"]==True)]
                monthly_tbl=base_tbl.groupby("Month").size().reindex(months_range,fill_value=0).astype(int)
                row={MONTH_LABELS[i-1]:int(monthly_tbl[i]) for i in months_range}; row["YTD"]=int(monthly_tbl.sum())
                st.dataframe(pd.DataFrame([row],index=[str(selected_year)]),use_container_width=True)

                st.subheader("NCR vs CRM Correlation (YTD)")
                def slide_ncr_vs_crm_correlation_ytd_fig(selected_year,selected_month,top_n=12):
                    from matplotlib.ticker import FuncFormatter
                    crm_c=df[(df["Is_Valid"]==True)&(df["Complaint_Category"]=="Quality")&(df["Year"]==selected_year)&(df["Month"]<=selected_month)]["Reason"].astype(str).str.strip().value_counts()
                    ncr_c=df_ncr_dash[(df_ncr_dash["Is_Valid"]==True)&(df_ncr_dash["Complaint_Category"]=="Quality")&(df_ncr_dash["Year"]==selected_year)&(df_ncr_dash["Month"]<=selected_month)]["Reason"].astype(str).str.strip().value_counts()
                    all_r=sorted(set(crm_c.index)|set(ncr_c.index))
                    df_cmp=pd.DataFrame({"Reason":all_r,"CRM":[int(crm_c.get(r,0)) for r in all_r],"NCR":[int(ncr_c.get(r,0)) for r in all_r]})
                    df_cmp=df_cmp.sort_values(["CRM","NCR","Reason"],ascending=[False,False,True]).head(int(top_n))
                    crm_v=df_cmp["CRM"].to_numpy(dtype=float); ncr_v=df_cmp["NCR"].to_numpy(dtype=float)
                    y=np.arange(len(df_cmp))
                    fig,ax=plt.subplots(figsize=(13.33,7.5),dpi=300)
                    ax.barh(y,-ncr_v,color="#D8C37D",height=0.48,label="NCR")
                    ax.barh(y,crm_v,color="#C1A02E",height=0.48,label="CRM")
                    ax.axvline(0,color="#666666",linewidth=1)
                    ax.set_yticks(y); ax.set_yticklabels([fill(r,28) for r in df_cmp["Reason"].tolist()],fontsize=11)
                    ax.invert_yaxis()
                    ms=float(max(crm_v.max(),ncr_v.max(),1.0))
                    ax.set_xlim(-ms*1.25,ms*1.25)
                    ax.xaxis.set_major_formatter(FuncFormatter(lambda v,pos:f"{int(abs(v))}"))
                    pad=ms*0.02
                    for i,(nv,cv) in enumerate(zip(ncr_v,crm_v)):
                        if nv>0: ax.text(-nv-pad,i,f"{int(nv)}",va="center",ha="right",fontsize=11,color="#333333")
                        if cv>0: ax.text(cv+pad,i,f"{int(cv)}",va="center",ha="left",fontsize=11,color="#333333")
                    ax.legend(loc="upper center",bbox_to_anchor=(0.5,-0.08),ncol=2,frameon=False,fontsize=12)
                    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.grid(False)
                    fig.subplots_adjust(left=0.28,right=0.98,top=0.90,bottom=0.18); return fig
                show_fig(slide_ncr_vs_crm_correlation_ytd_fig(selected_year,selected_month,TOPN_CORRELATION))
                slides_for_pdf.append({"title":"NCR vs CRM - Correlation (YTD)","fig":slide_ncr_vs_crm_correlation_ytd_fig(selected_year,selected_month,TOPN_CORRELATION)})

                st.subheader("Top Customers (FINAL)")
                def slide_valid_quality_by_customer_top5_cm_vs_ytd_fig(selected_year,selected_month,top_n=5):
                    base=df[(df["Is_Valid"]==True)&(df["Complaint_Category"]=="Quality")].copy()
                    CUST_COL="Customer"
                    if CUST_COL not in base.columns:
                        cand=[c for c in base.columns if "customer" in str(c).lower()]
                        raise KeyError(f"Customer column not found. Similar: {cand}")
                    base[CUST_COL]=base[CUST_COL].astype(str).str.strip()
                    cm=base[(base["Year"]==selected_year)&(base["Month"]==selected_month)][CUST_COL].value_counts()
                    ytd=base[(base["Year"]==selected_year)&(base["Month"].between(1,selected_month))][CUST_COL].value_counts()
                    all_c=pd.Index(ytd.index).union(cm.index)
                    summ=pd.DataFrame({"CM":cm.reindex(all_c,fill_value=0),"YTD":ytd.reindex(all_c,fill_value=0)})
                    summ=summ.sort_values(["YTD","CM"],ascending=False).head(int(top_n))
                    x=np.arange(len(summ)); w=0.35
                    fig,ax=plt.subplots(figsize=(13.33,7.5),dpi=300)
                    b1=ax.bar(x-w/2,summ["CM"],w,color="#006394",label="Current Month")
                    b2=ax.bar(x+w/2,summ["YTD"],w,color="#C1A02E",label="YTD")
                    add_simple_value_labels(ax,b1,lambda v:f"{int(v)}",1); add_simple_value_labels(ax,b2,lambda v:f"{int(v)}",1)
                    ax.set_xticks(x); ax.set_xticklabels([fill(str(c),22) for c in summ.index],rotation=0,ha="center",fontsize=12)
                    ax.set_ylabel("Count"); ax.set_ylim(0,max(1,summ[["CM","YTD"]].to_numpy().max())*1.18)
                    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.grid(False)
                    ax.legend(loc="upper center",bbox_to_anchor=(0.5,-0.08),ncol=2,frameon=False,fontsize=12)
                    fig.subplots_adjust(bottom=0.20); return fig
                try:
                    show_fig(slide_valid_quality_by_customer_top5_cm_vs_ytd_fig(selected_year,selected_month,TOPN_CUSTOMER))
                    slides_for_pdf.append({"title":"FINAL - Top Customers (CM vs YTD)","fig":slide_valid_quality_by_customer_top5_cm_vs_ytd_fig(selected_year,selected_month,TOPN_CUSTOMER)})
                except Exception as e:
                    st.warning(f"Customer slide skipped: {e}")

                st.divider()

                # ── CC RATIO CHARTS ──
                st.header("CC Ratio")

                # Load FG Invoiced from Supabase
                try:
                    fg_rows = supabase.table("qm_fg_invoiced").select("*").execute()
                    fg_data = {(r["year"], r["month"]): r["fg_invoiced"] for r in (fg_rows.data or [])}
                except Exception:
                    fg_data = {}

                # Check which months in selected year are missing FG Invoiced
                missing_months = [m for m in months_range if fg_data.get((selected_year, m)) is None]

                if missing_months:
                    st.warning(f"FG Invoiced data missing for: {', '.join([MONTH_LABELS[m-1]+'-'+str(selected_year)[-2:] for m in missing_months])}")
                    with st.form("qm_fg_form"):
                        st.markdown("**Enter missing FG Invoiced values:**")
                        fg_inputs = {}
                        cols = st.columns(min(len(missing_months), 6))
                        for i, m in enumerate(missing_months):
                            fg_inputs[m] = cols[i % 6].number_input(
                                f"{MONTH_LABELS[m-1]}-{str(selected_year)[-2:]}",
                                min_value=0, value=0, step=1, key=f"fg_{selected_year}_{m}"
                            )
                        submitted = st.form_submit_button("Save FG Invoiced", type="primary")
                        if submitted:
                            for m, val in fg_inputs.items():
                                if val > 0:
                                    try:
                                        supabase.table("qm_fg_invoiced").upsert({
                                            "year": selected_year, "month": m,
                                            "fg_invoiced": val, "updated_by": name
                                        }, on_conflict="year,month").execute()
                                        fg_data[(selected_year, m)] = val
                                    except Exception as e:
                                        st.error(f"Could not save {MONTH_LABELS[m-1]}: {e}")
                            st.success("Saved! Regenerating charts...")
                            st.rerun()

                months_with_fg = [m for m in months_range if fg_data.get((selected_year, m)) is not None]

                if months_with_fg:
                    def slide_cc_ratio_fig(selected_year, category_filter=None, title_label="Total CC Ratio"):
                        # Source: ISSUED file, Is_Valid == True
                        base = df_issued[df_issued["Is_Valid"] == True].copy()
                        if category_filter:
                            base = base[base["Complaint_Category"] == category_filter]
                        else:
                            base = base[base["Complaint_Category"].isin(["Quality", "Service"])]

                        ratios, cc_counts, fg_vals, labels = [], [], [], []
                        for m in months_with_fg:
                            cc = int(base[(base["Year"] == selected_year) & (base["Month"] == m)].shape[0])
                            fg = float(fg_data.get((selected_year, m), 0))
                            ratio = (cc / fg * 100) if fg > 0 else 0.0
                            ratios.append(ratio)
                            cc_counts.append(cc)
                            fg_vals.append(fg)
                            labels.append(MONTH_LABELS[m-1])

                        # YTD totals
                        ytd_cc  = sum(cc_counts)
                        ytd_fg  = sum(fg_vals)
                        ytd_ratio = (ytd_cc / ytd_fg * 100) if ytd_fg > 0 else 0.0
                        all_labels  = labels + ["YTD"]
                        all_cc      = cc_counts + [ytd_cc]
                        all_fg      = fg_vals   + [ytd_fg]
                        all_ratios  = ratios    + [ytd_ratio]

                        color_cc = "#006394" if category_filter == "Quality" else "#C1A02E"
                        color_fg = "#D8C37D"
                        n = len(all_labels)
                        x = np.arange(n)
                        w = 0.35

                        fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)

                        # Side-by-side bars: CC and FG
                        bars_cc = ax.bar(x - w/2, all_cc, width=w, color=color_cc, label="CC Count", alpha=0.9)
                        bars_fg = ax.bar(x + w/2, all_fg, width=w, color=color_fg, label="FG Invoiced", alpha=0.9)

                        # CC count labels above CC bars
                        for bar, val in zip(bars_cc, all_cc):
                            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(all_fg)*0.01,
                                    f"{int(val)}", ha="center", va="bottom", fontsize=9,
                                    color=color_cc, fontweight="bold")

                        # FG labels above FG bars
                        for bar, val in zip(bars_fg, all_fg):
                            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(all_fg)*0.01,
                                    f"{int(val)}", ha="center", va="bottom", fontsize=9, color="#666666")

                        # Ratio % labels above each month group (centered between the two bars)
                        ratio_y = max(all_fg) * 1.10
                        for i, ratio in enumerate(all_ratios):
                            ax.text(x[i], ratio_y, f"{ratio:.2f}%",
                                    ha="center", va="bottom", fontsize=10,
                                    color="#333333", fontweight="bold")

                        # Connect ratio labels with a line (monthly only, exclude YTD)
                        ax.plot(x[:-1], [max(all_fg)*1.10]*len(months_with_fg),
                                alpha=0)  # invisible anchor; draw real line below
                        ratio_line_y = [ratio_y + r * max(all_fg) * 0.003 for r in ratios]
                        ax.plot(x[:-1], ratio_line_y, color="#E63946", linewidth=2,
                                marker="o", markersize=5, label="Ratio % (monthly)", zorder=5)

                        # Trendline through monthly ratios only
                        if len(ratios) >= 2:
                            xf = np.arange(len(ratios), dtype=float)
                            coeff = np.polyfit(xf, ratios, 1)
                            trend_y = [np.polyval(coeff, xi) * max(all_fg) * 0.003 + ratio_y for xi in xf]
                            ax.plot(x[:-1], trend_y, linestyle="--", linewidth=1.5,
                                    color="#999999", label="Trend", zorder=4)

                        ax.set_xticks(x)
                        ax.set_xticklabels(all_labels, fontsize=11)
                        ax.set_ylabel("Count")
                        ax.set_ylim(0, max(all_fg) * 1.25)
                        ax.spines["top"].set_visible(False)
                        ax.spines["right"].set_visible(False)
                        ax.grid(False)
                        ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.08), ncol=4, frameon=False, fontsize=10)
                        plt.tight_layout(rect=[0, 0.05, 1, 1])
                        return fig

                    st.subheader("Total Customer Complaints Ratio (Quality + Service)")
                    show_fig(slide_cc_ratio_fig(selected_year, category_filter=None, title_label="Total CC Ratio"))
                    slides_for_pdf.append({"title": "Total CC Ratio", "fig": slide_cc_ratio_fig(selected_year, category_filter=None, title_label="Total CC Ratio")})

                    st.subheader("Quality Customer Complaints Ratio")
                    show_fig(slide_cc_ratio_fig(selected_year, category_filter="Quality", title_label="Quality CC Ratio"))
                    slides_for_pdf.append({"title": "Quality CC Ratio", "fig": slide_cc_ratio_fig(selected_year, category_filter="Quality", title_label="Quality CC Ratio")})

                else:
                    st.info("Enter FG Invoiced values above to generate CC Ratio charts.")

                # PDF Export
                st.divider()
                st.header("Export")
                pdf_buf = build_ppt_pdf(slides_for_pdf, dpi=300)
                st.download_button(
                    label="📥 Download PPT-style PDF",
                    data=pdf_buf,
                    file_name=f"QM_Dashboard_{selected_year}-{selected_month:02d}.pdf",
                    mime="application/pdf",
                )

        except Exception as e:
            st.error(f"Error: {str(e)}")

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
                                file_path=f"QM/{datetime.now().strftime('%Y%m%d_%H%M%S')}_{doc_file.name}"
                                supabase.storage.from_("asset").upload(file_path,doc_file.read())
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
