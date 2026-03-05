# pages/4_QM.py - Quality Maintenance Pillar
import os
import io
import sys
import tempfile
import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import matplotlib as mpl
from textwrap import fill
from datetime import datetime
from urllib.parse import quote

# Add utils to path for pipeline import
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib.colors import HexColor

from utils.supabase_client import get_supabase
from utils.qm_pipeline import (
    read_excel_from_upload, build_dataset_final_issued, build_dataset_ncr,
    FINAL_APPROVAL_COL, CREATION_DATETIME_COL,
)

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

mpl.rcParams["font.family"] = "DejaVu Sans"
mpl.rcParams["font.size"]   = 11
mpl.rcParams["figure.dpi"]  = 120

COLORS = {"Service": "#C1A02E", "Quality": "#006394", "Invalid": "#D8C37D"}

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
        {
            "name": "Entry", "icon": "⚪",
            "focus": "QM by basic methods. Identification of Q factors.",
            "criteria": [
                "QM Pillar and targets defined (Cost of non-Quality, CC, Internal Rejections)",
                "Basic Root Cause Analysis implemented",
                "Action plans generated to prevent defect re-occurrence",
                "Identification of Q factors (Machine, Material, Manpower, Method)",
                "QA Matrix tool developed",
                "KPIs monitored and suggestions for improvement available",
                "Customer Complaints trend is decreasing",
            ]
        },
        {
            "name": "Bronze", "icon": "🥉",
            "focus": "Reduce Waste and Complaints through a Quality system.",
            "criteria": [
                "Analysis implemented for all Q factors",
                "QA Matrix used to tackle high risk defects",
                "Action plans effective (IR and CC data reduced by 10%)",
                "Accurate and updated MPCC for all FGs",
                "Quality skills assessed and improved for all shopfloor employees",
                "QC council meetings implemented",
                "Shopfloor/QC/MT employees have basic knowledge on Q factors",
            ]
        },
        {
            "name": "Silver", "icon": "🥈",
            "focus": "Quality Awareness is clear.",
            "criteria": [
                "Reduction of NCR & CC ratio by 10% vs last year",
                "Evidence of 2 major Quality improvement projects",
                "QA matrix showing decrease in RPN of 50% of defects",
                "Cost of Poor Quality actual equals target COPQ",
                "Evidence of shopfloor involvement in NCR and CC RCAs",
                "Quality circles used to transfer knowledge",
                "Clear evidence of Q Points for machines, materials, methods, men and measurement",
            ]
        },
        {
            "name": "Gold", "icon": "🥇",
            "focus": "Quality is embedded in the BU culture.",
            "criteria": [
                "Decreasing monthly trend of NCR & CC ratio",
                "Quality circles transfer knowledge from engineers to operators and vice versa",
                "QA Matrix used for all machine groups, updated quarterly",
                "Evidence of 2 major Quality improvement projects",
                "Formalised links between defects and Q Factors with physical evidence on shopfloor",
                "Quality defined for the BU",
            ]
        },
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
    st.caption("Upload the 3 required Excel files, select the period, then generate charts.")

    # ── File uploads ──
    col1, col2, col3 = st.columns(3)
    with col1:
        final_file  = st.file_uploader("📂 FINAL Approval file", type=["xls","xlsx"], key="qm_final")
    with col2:
        issued_file = st.file_uploader("📂 ISSUED file",         type=["xls","xlsx"], key="qm_issued")
    with col3:
        ncr_file    = st.file_uploader("📂 NCR file",            type=["xls","xlsx"], key="qm_ncr")

    if not final_file or not issued_file or not ncr_file:
        st.info("⬆️ Upload all 3 files above to enable the dashboard.")
    else:
        try:
            with st.spinner("Loading data..."):
                df_final_loaded  = read_excel_from_upload(final_file)
                df_issued_loaded = read_excel_from_upload(issued_file)
                df_ncr_loaded    = read_excel_from_upload(ncr_file)

                final_pkg  = build_dataset_final_issued(df_final_loaded,  date_col=FINAL_APPROVAL_COL,     dataset_name="FINAL")
                issued_pkg = build_dataset_final_issued(df_issued_loaded, date_col=CREATION_DATETIME_COL,  dataset_name="ISSUED")
                ncr_pkg    = build_dataset_ncr(df_ncr_loaded,             date_col=FINAL_APPROVAL_COL,     dataset_name="NCR")

                df  = final_pkg["cleaned_flagged"].copy()
                df_issued    = issued_pkg["cleaned_flagged"].copy()
                df_ncr_dash  = ncr_pkg["cleaned_flagged"].copy()
                df_raw_flagged = final_pkg["raw_flagged"].copy()

                for _d in (df, df_issued, df_ncr_dash):
                    _d["Year"]  = pd.to_numeric(_d["Year"],  errors="coerce").dropna().astype(int) if _d["Year"].notna().any() else _d["Year"]
                    _d["Month"] = pd.to_numeric(_d["Month"], errors="coerce").dropna().astype(int) if _d["Month"].notna().any() else _d["Month"]
                    _d.dropna(subset=["Year","Month"], inplace=True)
                    _d["Year"]  = _d["Year"].astype(int)
                    _d["Month"] = _d["Month"].astype(int)

            st.write("FINAL columns:", list(df_final_loaded.columns))
            st.success("✅ Files loaded successfully!")

            # ── Period selector ──
            df_dates = df[["Year","Month"]].drop_duplicates().sort_values(["Year","Month"])
            years = sorted(df_dates["Year"].unique().tolist())
            col1, col2 = st.columns(2)
            with col1:
                selected_year  = st.selectbox("Year",  years, index=len(years)-1, key="qm_year")
            with col2:
                months = sorted(df_dates[df_dates["Year"]==selected_year]["Month"].unique().tolist())
                selected_month = st.selectbox("Month", months, index=len(months)-1, key="qm_month")
            selected_month = int(selected_month)

            # ── Top-N controls ──
            with st.expander("⚙️ Display Controls", expanded=False):
                c1,c2,c3,c4 = st.columns(4)
                with c1:
                    TOPN_Q_DEFECT = st.selectbox("Top defects Quality (CM vs YTD)", ["All"]+list(range(3,41)), index=13, key="qm_topn1")
                    TOPN_Q_DEFECT = None if TOPN_Q_DEFECT=="All" else int(TOPN_Q_DEFECT)
                with c2:
                    TOPN_Q_CM = st.selectbox("Top defects Quality (CM)", ["All"]+list(range(3,41)), index=8, key="qm_topn2")
                    TOPN_Q_CM = None if TOPN_Q_CM=="All" else int(TOPN_Q_CM)
                with c3:
                    TOPN_SVC = st.selectbox("Top reasons Service (CM vs YTD)", ["All"]+list(range(3,41)), index=13, key="qm_topn3")
                    TOPN_SVC = None if TOPN_SVC=="All" else int(TOPN_SVC)
                with c4:
                    TOPN_NCR = st.selectbox("Top NCR defects", ["All"]+list(range(3,51)), index=18, key="qm_topn4")
                    TOPN_NCR = None if TOPN_NCR=="All" else int(TOPN_NCR)

            run = st.button("🔄 Generate Charts", type="primary", key="qm_run")

            if run:
                # ── HELPERS ──
                def fig_to_bytes(fig, dpi=150):
                    buf = io.BytesIO()
                    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight")
                    buf.seek(0)
                    return buf

                def show(fig):
                    st.image(fig_to_bytes(fig))
                    plt.close(fig)

                def add_labels(ax, bars, fmt_fn, pad):
                    for b in bars:
                        h = b.get_height()
                        if np.isnan(h) or h == 0: continue
                        ax.text(b.get_x()+b.get_width()/2, h+pad, fmt_fn(h),
                                ha="center", va="bottom", fontsize=9, color="#4D4D4D", clip_on=False)

                def donut_chart(ax, counts):
                    labels = ["Service","Quality","Invalid"]
                    values = np.array([int(counts.get(l,0)) for l in labels], dtype=float)
                    total  = values.sum()
                    if total == 0:
                        ax.text(0.5,0.5,"No data",ha="center",va="center"); ax.axis("off"); return
                    rw = 0.45; lr = 1.0 - rw + rw/2
                    wedges,_ = ax.pie(values, labels=None, colors=[COLORS[l] for l in labels],
                                      startangle=90, counterclock=False, radius=1.0,
                                      wedgeprops=dict(width=rw, edgecolor="white", linewidth=1))
                    for w,v in zip(wedges,values):
                        if v<=0: continue
                        ang=(w.theta2+w.theta1)/2.0
                        x=lr*np.cos(np.deg2rad(ang)); y=lr*np.sin(np.deg2rad(ang))
                        ax.text(x,y,f"{int(v)}",ha="center",va="center",fontsize=11,color="#4D4D4D")
                    ax.axis("equal")

                MONTH_LABELS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

                slides_for_pdf = []

                # ── FINAL Section ──
                st.header("FINAL (CRM Approved)")

                # Donuts
                st.subheader("Overview")
                df_m = df[(df["Year"]==selected_year)&(df["Month"]==selected_month)]
                df_y = df[(df["Year"]==selected_year)&(df["Month"]<=selected_month)]
                fig,axes = plt.subplots(1,2,figsize=(12,6))
                donut_chart(axes[0], df_m["Complaint_Category"].value_counts())
                donut_chart(axes[1], df_y["Complaint_Category"].value_counts())
                ml = pd.to_datetime(f"{selected_year}-{selected_month:02d}-01").strftime("%b-%y")
                axes[0].text(0.5,-0.10,ml,transform=axes[0].transAxes,ha="center",va="top",fontsize=13,color="#4D4D4D")
                axes[1].text(0.5,-0.10,f"{selected_year} YTD",transform=axes[1].transAxes,ha="center",va="top",fontsize=13,color="#4D4D4D")
                fig.legend(["Service","Quality","Invalid"],loc="lower center",ncol=3,frameon=False)
                plt.tight_layout(rect=[0,0.08,1,1])
                slides_for_pdf.append(("FINAL - Overview", fig_to_bytes(fig)))
                show(fig)

                # Valid Quality Count
                st.subheader("Quality Count (CM vs YTD)")
                prev = selected_year-1
                months_range = list(range(1,selected_month+1))
                cats = MONTH_LABELS[:selected_month]+["YTD"]
                base_q = df[(df["Is_Valid"]==True)&(df["Complaint_Category"]=="Quality")]

                def monthly_q(yr):
                    return base_q[base_q["Year"]==yr].groupby("Month").size().reindex(months_range,fill_value=0)
                def ytd_q(yr):
                    return int(base_q[(base_q["Year"]==yr)&(base_q["Month"]<=selected_month)].shape[0])

                prev_v = list(monthly_q(prev).values)+[ytd_q(prev)]
                sel_v  = list(monthly_q(selected_year).values)+[ytd_q(selected_year)]
                x = np.arange(len(cats)); w=0.28
                fig,ax=plt.subplots(figsize=(12,6))
                b1=ax.bar(x-w/2,prev_v,w,label=str(prev),color="#006394")
                b2=ax.bar(x+w/2,sel_v,w,label=str(selected_year),color="#C1A02E")
                add_labels(ax,b1,lambda v:f"{int(v)}",1); add_labels(ax,b2,lambda v:f"{int(v)}",1)
                y_arr=np.array(list(monthly_q(selected_year).values),dtype=float)
                if len(months_range)>=2 and np.any(y_arr>0):
                    coeff=np.polyfit(np.arange(len(months_range),dtype=float),y_arr,1)
                    ax.plot(x[:len(months_range)],np.polyval(coeff,np.arange(len(months_range),dtype=float)),
                            linestyle="--",linewidth=1.5,color="#C1A02E",label=f"Trend {selected_year}")
                ax.set_xticks(x); ax.set_xticklabels(cats)
                ax.legend(loc="upper center",bbox_to_anchor=(0.5,-0.08),ncol=3,frameon=False)
                ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.grid(False)
                plt.tight_layout(rect=[0,0.05,1,1])
                slides_for_pdf.append(("FINAL - Quality Count", fig_to_bytes(fig)))
                show(fig)

                # Quality Defects CM vs YTD
                st.subheader("Quality Defects (CM vs YTD)")
                base_qd = df[(df["Is_Valid"]==True)&(df["Complaint_Category"]=="Quality")].copy()
                base_qd["Defect"] = base_qd["Reason"].astype(str).str.strip()
                cm_p=base_qd[(base_qd["Year"]==prev)&(base_qd["Month"]==selected_month)]["Defect"].value_counts()
                cm_s=base_qd[(base_qd["Year"]==selected_year)&(base_qd["Month"]==selected_month)]["Defect"].value_counts()
                ytd_p=base_qd[(base_qd["Year"]==prev)&(base_qd["Month"]<=selected_month)]["Defect"].value_counts()
                ytd_s=base_qd[(base_qd["Year"]==selected_year)&(base_qd["Month"]<=selected_month)]["Defect"].value_counts()
                all_d=pd.Index(ytd_s.index).union(ytd_p.index).union(cm_s.index).union(cm_p.index)
                summ=pd.DataFrame({"CM_prev":cm_p.reindex(all_d,fill_value=0),"CM_sel":cm_s.reindex(all_d,fill_value=0),
                                   "YTD_prev":ytd_p.reindex(all_d,fill_value=0),"YTD_sel":ytd_s.reindex(all_d,fill_value=0)})
                summ=summ.sort_values(["YTD_sel","CM_sel"],ascending=False)
                if TOPN_Q_DEFECT: summ=summ.head(TOPN_Q_DEFECT)
                cmp=pd.to_datetime(f"{prev}-{selected_month:02d}-01").strftime("%b-%y")
                cms=pd.to_datetime(f"{selected_year}-{selected_month:02d}-01").strftime("%b-%y")
                x=np.arange(len(summ)); w=0.18
                fig,ax=plt.subplots(figsize=(14,7))
                b1=ax.bar(x-1.5*w,summ["CM_prev"],w,color="#0F68B9",label=f"CM {cmp}")
                b2=ax.bar(x-0.5*w,summ["CM_sel"],w,color="#D8C37D",label=f"CM {cms}")
                b3=ax.bar(x+0.5*w,summ["YTD_prev"],w,color="#006394",label=f"YTD {prev}")
                b4=ax.bar(x+1.5*w,summ["YTD_sel"],w,color="#C1A02E",label=f"YTD {selected_year}")
                for b in [b1,b2,b3,b4]: add_labels(ax,b,lambda v:f"{int(v)}",1)
                ax.set_xticks(x); ax.set_xticklabels([fill(str(d),16) for d in summ.index],rotation=35,ha="right",fontsize=9)
                ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.grid(False)
                ax.legend(loc="upper center",bbox_to_anchor=(0.5,-0.22),ncol=4,frameon=False,fontsize=10)
                fig.subplots_adjust(bottom=0.35)
                slides_for_pdf.append(("FINAL - Quality Defects CM vs YTD", fig_to_bytes(fig)))
                show(fig)

                # Quality Cost
                st.subheader("Quality Cost (Credit Note)")
                try:
                    decision_candidates=[c for c in df_raw_flagged.columns if "decision" in str(c).lower()]
                    if decision_candidates:
                        DECISION_COL=decision_candidates[0]
                        base_cost=df_raw_flagged[(df_raw_flagged["Is_Valid"]==True)&(df_raw_flagged["Complaint_Category"]=="Quality")].copy()
                        base_cost[DECISION_COL]=base_cost[DECISION_COL].astype(str).str.strip()
                        base_cost=base_cost[base_cost[DECISION_COL].str.lower()=="credit note"]
                        base_cost["Defect"]=base_cost["Reason"].astype(str).str.strip()
                        base_cost["Cost Amount"]=pd.to_numeric(base_cost["Cost Amount"],errors="coerce").fillna(0)
                        cm_pc=base_cost[(base_cost["Year"]==prev)&(base_cost["Month"]==selected_month)].groupby("Defect")["Cost Amount"].sum()
                        cm_sc=base_cost[(base_cost["Year"]==selected_year)&(base_cost["Month"]==selected_month)].groupby("Defect")["Cost Amount"].sum()
                        ytd_pc=base_cost[(base_cost["Year"]==prev)&(base_cost["Month"]<=selected_month)].groupby("Defect")["Cost Amount"].sum()
                        ytd_sc=base_cost[(base_cost["Year"]==selected_year)&(base_cost["Month"]<=selected_month)].groupby("Defect")["Cost Amount"].sum()
                        summ_cost=pd.DataFrame({"CM_prev":cm_pc,"CM_sel":cm_sc,"YTD_prev":ytd_pc,"YTD_sel":ytd_sc}).fillna(0)
                        summ_cost=summ_cost.sort_values(["YTD_sel","CM_sel"],ascending=False).head(15)
                        x=np.arange(len(summ_cost)); w=0.18
                        fig,ax=plt.subplots(figsize=(14,7))
                        b1=ax.bar(x-1.5*w,summ_cost["CM_prev"],w,label=f"CM {prev}",color="#0F68B9")
                        b2=ax.bar(x-0.5*w,summ_cost["CM_sel"],w,label=f"CM {selected_year}",color="#D8C37D")
                        b3=ax.bar(x+0.5*w,summ_cost["YTD_prev"],w,label=f"YTD {prev}",color="#006394")
                        b4=ax.bar(x+1.5*w,summ_cost["YTD_sel"],w,label=f"YTD {selected_year}",color="#C1A02E")
                        ymax=summ_cost.to_numpy().max(); pad=ymax*0.015
                        def fmt_cost(v): return f"SAR {int(v/1000)}K" if v>=1000 else f"SAR {int(v)}"
                        for bars in [b1,b2,b3,b4]:
                            for bar in bars:
                                h=bar.get_height()
                                if h>0:
                                    ax.text(bar.get_x()+bar.get_width()/2,h+pad,fmt_cost(h),
                                            ha="center",va="bottom",fontsize=9,rotation=90,color="#333333",clip_on=False)
                        ax.set_xticks(x); ax.set_xticklabels([fill(str(d),18) for d in summ_cost.index],rotation=30,ha="right",fontsize=9)
                        ax.set_ylabel("Cost Amount (SAR)")
                        ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.grid(False)
                        ax.legend(loc="upper center",bbox_to_anchor=(0.5,-0.22),ncol=4,frameon=False,fontsize=10)
                        fig.subplots_adjust(bottom=0.35)
                        slides_for_pdf.append(("FINAL - Quality Cost (Credit Note)", fig_to_bytes(fig)))
                        show(fig)
                except Exception as cost_err:
                    st.warning(f"Cost slide skipped: {cost_err}")

                # Service Count
                st.subheader("Service Count")
                base_svc=df[(df["Is_Valid"]==True)&(df["Complaint_Category"]=="Service")]
                def monthly_s(yr): return base_svc[base_svc["Year"]==yr].groupby("Month").size().reindex(months_range,fill_value=0)
                def ytd_s(yr): return int(base_svc[(base_svc["Year"]==yr)&(base_svc["Month"]<=selected_month)].shape[0])
                prev_sv=list(monthly_s(prev).values)+[ytd_s(prev)]
                sel_sv=list(monthly_s(selected_year).values)+[ytd_s(selected_year)]
                x=np.arange(len(cats)); w=0.28
                fig,ax=plt.subplots(figsize=(12,6))
                b1=ax.bar(x-w/2,prev_sv,w,label=str(prev),color="#006394")
                b2=ax.bar(x+w/2,sel_sv,w,label=str(selected_year),color="#C1A02E")
                add_labels(ax,b1,lambda v:f"{int(v)}",1); add_labels(ax,b2,lambda v:f"{int(v)}",1)
                ax.set_xticks(x); ax.set_xticklabels(cats)
                ax.legend(loc="upper center",bbox_to_anchor=(0.5,-0.08),ncol=3,frameon=False)
                ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.grid(False)
                plt.tight_layout(rect=[0,0.05,1,1])
                slides_for_pdf.append(("FINAL - Service Count", fig_to_bytes(fig)))
                show(fig)

                # ── NCR Section ──
                st.header("NCR")
                st.subheader("Quality Defects (CM vs YTD)")
                base_ncr=df_ncr_dash[(df_ncr_dash["Is_Valid"]==True)&(df_ncr_dash["Complaint_Category"]=="Quality")].copy()
                base_ncr["Defect"]=base_ncr["Reason"].astype(str).str.strip()
                ncr_cm=base_ncr[(base_ncr["Year"]==selected_year)&(base_ncr["Month"]==selected_month)]["Defect"].value_counts()
                ncr_ytd=base_ncr[(base_ncr["Year"]==selected_year)&(base_ncr["Month"]<=selected_month)]["Defect"].value_counts()
                all_ncr=pd.Index(ncr_ytd.index).union(ncr_cm.index)
                summ_ncr=pd.DataFrame({"CM":ncr_cm.reindex(all_ncr,fill_value=0),"YTD":ncr_ytd.reindex(all_ncr,fill_value=0)})
                summ_ncr=summ_ncr.sort_values(["YTD","CM"],ascending=False)
                if TOPN_NCR: summ_ncr=summ_ncr.head(TOPN_NCR)
                x=np.arange(len(summ_ncr)); w=0.35
                fig,ax=plt.subplots(figsize=(14,7))
                b1=ax.bar(x-w/2,summ_ncr["CM"],w,color="#006394",label="Current Month")
                b2=ax.bar(x+w/2,summ_ncr["YTD"],w,color="#C1A02E",label="YTD")
                add_labels(ax,b1,lambda v:f"{int(v)}",1); add_labels(ax,b2,lambda v:f"{int(v)}",1)
                ax.set_xticks(x); ax.set_xticklabels([fill(str(d),18) for d in summ_ncr.index],rotation=35,ha="right",fontsize=10)
                ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.grid(False)
                ax.legend(loc="upper center",bbox_to_anchor=(0.5,-0.08),ncol=2,frameon=False,fontsize=12)
                fig.subplots_adjust(bottom=0.22)
                slides_for_pdf.append(("NCR - Quality Defects CM vs YTD", fig_to_bytes(fig)))
                show(fig)

                # NCR vs CRM Correlation
                st.subheader("NCR vs CRM Correlation (YTD)")
                from matplotlib.ticker import FuncFormatter
                crm_c=df[(df["Is_Valid"]==True)&(df["Complaint_Category"]=="Quality")&(df["Year"]==selected_year)&(df["Month"]<=selected_month)]["Reason"].astype(str).str.strip().value_counts()
                ncr_c=df_ncr_dash[(df_ncr_dash["Is_Valid"]==True)&(df_ncr_dash["Complaint_Category"]=="Quality")&(df_ncr_dash["Year"]==selected_year)&(df_ncr_dash["Month"]<=selected_month)]["Reason"].astype(str).str.strip().value_counts()
                all_r=sorted(set(crm_c.index)|set(ncr_c.index))
                df_cmp=pd.DataFrame({"Reason":all_r,"CRM":[int(crm_c.get(r,0)) for r in all_r],"NCR":[int(ncr_c.get(r,0)) for r in all_r]})
                df_cmp=df_cmp.sort_values(["CRM","NCR"],ascending=[False,False]).head(12)
                crm_v=df_cmp["CRM"].to_numpy(dtype=float); ncr_v=df_cmp["NCR"].to_numpy(dtype=float)
                y=np.arange(len(df_cmp))
                fig,ax=plt.subplots(figsize=(14,7))
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
                fig.subplots_adjust(left=0.28,right=0.98,top=0.90,bottom=0.18)
                slides_for_pdf.append(("NCR vs CRM Correlation (YTD)", fig_to_bytes(fig)))
                show(fig)

                # ── PDF Export ──
                st.divider()
                st.subheader("📥 Export")
                def build_pdf(slides):
                    W,H=960,540
                    pdf_buf=io.BytesIO()
                    c=canvas.Canvas(pdf_buf,pagesize=(W,H))
                    for title,img_buf in slides:
                        c.setFillColor(HexColor("#006394")); c.setFont("Helvetica-Bold",26)
                        c.drawString(40,H-58,title)
                        c.setStrokeColor(HexColor("#006394")); c.setLineWidth(4)
                        c.line(40,H-78,W-40,H-78)
                        img=ImageReader(img_buf)
                        c.drawImage(img,40,40,width=W-80,height=H-92-40,preserveAspectRatio=True,anchor="c")
                        c.showPage()
                    c.save(); pdf_buf.seek(0)
                    return pdf_buf

                pdf_buf=build_pdf(slides_for_pdf)
                st.download_button(
                    label="📥 Download PDF Report",
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
        st.error(f"Could not load documents: {str(e)}")
        st.stop()

    if can_edit:
        with st.expander("➕ Add / Remove Required Document", expanded=False):
            col1, col2 = st.columns([3,1])
            with col1:
                new_doc_name = st.text_input("Document Name", placeholder="e.g. QC Council Meeting Minutes", key="qm_new_doc_name")
                new_doc_desc = st.text_input("Description (optional)", placeholder="Brief description", key="qm_new_doc_desc")
            with col2:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("➕ Add", use_container_width=True, key="qm_add_req"):
                    if not new_doc_name:
                        st.error("Please enter a document name.")
                    else:
                        supabase.table("qm_required_docs").insert({
                            "doc_name": new_doc_name, "description": new_doc_desc or "", "added_by": name
                        }).execute()
                        st.success(f"✅ '{new_doc_name}' added!")
                        st.rerun()
            if req_docs.data:
                st.divider()
                st.markdown("**Remove a requirement:**")
                del_map = {r['doc_name']: r['id'] for r in req_docs.data}
                col1,col2 = st.columns([3,1])
                with col1:
                    sel_del = st.selectbox("Select to remove", list(del_map.keys()), key="qm_del_req")
                with col2:
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("🗑️ Remove", use_container_width=True, key="qm_btn_del_req"):
                        supabase.table("qm_required_docs").delete().eq("id", del_map[sel_del]).execute()
                        st.success(f"Removed '{sel_del}'")
                        st.rerun()

    st.divider()

    if not req_docs.data:
        st.info("No required documents defined yet.")
    else:
        for req in req_docs.data:
            doc_name = req['doc_name']
            doc_desc = req.get('description','')
            uploads  = [u for u in (all_uploads.data or []) if u['doc_type'] == doc_name]
            has_up   = len(uploads) > 0
            status_icon  = "✅" if has_up else "❌"
            status_color = "#2ea043" if has_up else "#f85149"

            st.markdown(f"""
                <div style="border:1px solid {'#2ea04355' if has_up else '#f8514955'};border-radius:10px;
                    padding:16px 20px;margin-bottom:14px;background:{'#2ea04308' if has_up else '#f8514908'};">
                    <div style="display:flex;align-items:center;gap:10px;margin-bottom:4px;">
                        <span style="font-size:18px">{status_icon}</span>
                        <span style="font-weight:700;font-size:16px;">{doc_name}</span>
                        <span style="font-size:11px;color:{status_color};margin-left:auto;">
                            {len(uploads)} file{'s' if len(uploads)!=1 else ''} uploaded
                        </span>
                    </div>
                    {f'<div style="font-size:12px;color:#8B949E;">{doc_desc}</div>' if doc_desc else ''}
                </div>
            """, unsafe_allow_html=True)

            for upload in uploads:
                ext = upload['file_name'].split('.')[-1].lower()
                file_url = f"https://sjcwzbftzpfylwdqiknh.supabase.co/storage/v1/object/public/asset/{quote(upload['file_path'], safe='/')}"
                uploaded_at = pd.to_datetime(upload['created_at']).strftime("%d %b %Y %H:%M")
                col1,col2,col3 = st.columns([4,2,1])
                with col1:
                    if ext in ['png','jpg','jpeg']:
                        st.image(file_url, width=200)
                    else:
                        st.markdown(f"📄 [{upload['file_name']}]({file_url})")
                    st.caption(f"Uploaded by {upload['uploaded_by']} · {uploaded_at}")
                with col3:
                    if can_edit:
                        if st.button("🗑️", key=f"qm_del_file_{upload['id']}"):
                            supabase.table("qm_documents").delete().eq("id", upload['id']).execute()
                            st.success("Deleted!"); st.rerun()

            if can_edit:
                with st.expander(f"⬆️ Upload file for '{doc_name}'", expanded=False):
                    doc_file = st.file_uploader("Choose file", type=["xlsx","pdf","png","jpg","docx"], key=f"qm_upload_{req['id']}")
                    if doc_file and st.button("Upload", key=f"qm_btn_upload_{req['id']}", type="primary"):
                        with st.spinner("Uploading..."):
                            try:
                                file_bytes = doc_file.read()
                                file_path  = f"QM/{datetime.now().strftime('%Y%m%d_%H%M%S')}_{doc_file.name}"
                                supabase.storage.from_("asset").upload(file_path, file_bytes)
                                supabase.table("qm_documents").insert({
                                    "doc_type": doc_name, "file_name": doc_file.name,
                                    "file_path": file_path, "uploaded_by": name,
                                }).execute()
                                st.success("✅ Uploaded!"); st.rerun()
                            except Exception as e:
                                st.error(f"Upload failed: {str(e)}")
            st.markdown("---")

# ════════════════════════════════════════
# TAB 4 — ACTION PLANS
# ════════════════════════════════════════
with tab4:
    st.markdown("### 🎯 Action Plans")

    try:
        actions = supabase.table("qm_action_plans").select("*").order("created_at", desc=True).execute()
        if actions.data:
            act_df = pd.DataFrame(actions.data)
            act_df["due_date"] = pd.to_datetime(act_df["due_date"]).dt.strftime("%d %b %Y")
            st.dataframe(
                act_df[["action","owner","due_date","status","notes"]].rename(columns={
                    "action":"Action","owner":"Owner","due_date":"Due Date","status":"Status","notes":"Notes"
                }), use_container_width=True
            )
            if can_edit:
                st.markdown("#### ✏️ Update Status")
                act_map = {f"{r['action'][:40]} — {r['owner']}": r['id'] for r in actions.data}
                col1,col2 = st.columns([3,1])
                with col1: sel_act = st.selectbox("Select action", list(act_map.keys()), key="qm_sel_act")
                with col2: new_act_s = st.selectbox("Status", ["Open","In Progress","Done"], key="qm_act_status")
                if st.button("✅ Update", key="qm_update_act"):
                    supabase.table("qm_action_plans").update({"status":new_act_s}).eq("id",act_map[sel_act]).execute()
                    st.success("Updated!"); st.rerun()

                st.markdown("#### 🗑️ Delete")
                del_act_map = {f"{r['action'][:40]} — {r['owner']}": r['id'] for r in actions.data}
                sel_del_act = st.selectbox("Select to delete", list(del_act_map.keys()), key="qm_del_act")
                if st.button("🗑️ Delete Action Plan", key="qm_btn_del_act"):
                    supabase.table("qm_action_plans").delete().eq("id",del_act_map[sel_del_act]).execute()
                    st.success("Deleted!"); st.rerun()
        else:
            st.info("No action plans yet.")
    except Exception as e:
        st.error(f"Could not load action plans: {str(e)}")

    if can_edit:
        st.divider()
        st.markdown("#### ➕ Add Action Plan")
        act_text  = st.text_input("Action", key="qm_act_text")
        own_text  = st.text_input("Owner",  key="qm_own_text")
        due_d     = st.date_input("Due Date", key="qm_due_date")
        stat_ap   = st.selectbox("Status", ["Open","In Progress","Done"], key="qm_new_act_status")
        notes_ap  = st.text_area("Notes", key="qm_notes_ap")
        if st.button("➕ Add Action Plan", type="primary", key="qm_add_act"):
            if not act_text or not own_text:
                st.error("Please fill in action and owner.")
            else:
                try:
                    supabase.table("qm_action_plans").insert({
                        "action": act_text, "owner": own_text,
                        "due_date": str(due_d), "status": stat_ap, "notes": notes_ap,
                    }).execute()
                    st.success("✅ Action plan added!"); st.rerun()
                except Exception as e:
                    st.error(f"Error: {str(e)}")
