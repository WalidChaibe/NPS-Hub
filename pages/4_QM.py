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
    load_settings_from_supabase, make_classifier,
)
from utils.supabase_client import get_supabase
# ── Notifications (inlined) ──
def _get_unread_count(sb, user_id):
    try:
        res = sb.table("notifications").select("id", count="exact").eq("user_id", user_id).eq("is_read", False).execute()
        return res.count or 0
    except Exception: return 0

def _get_notifications(sb, user_id, limit=20):
    try:
        res = sb.table("notifications").select("*").eq("user_id", user_id).order("created_at", desc=True).limit(limit).execute()
        return res.data or []
    except Exception: return []

def _mark_all_read(sb, user_id):
    try: sb.table("notifications").update({"is_read": True}).eq("user_id", user_id).eq("is_read", False).execute()
    except Exception: pass

def _mark_one_read(sb, notif_id):
    try: sb.table("notifications").update({"is_read": True}).eq("id", notif_id).execute()
    except Exception: pass

def create_notification(sb, user_id, title, message, notif_type="info", action_plan_id=None):
    try:
        payload = {"user_id": user_id, "title": title, "message": message, "type": notif_type, "is_read": False}
        if action_plan_id: payload["action_plan_id"] = action_plan_id
        sb.table("notifications").insert(payload).execute()
    except Exception: pass

def check_action_plan_notifications(sb):
    from datetime import datetime as _dt
    today = _dt.now().date()
    try:
        res = sb.table("qm_action_plans").select("*").in_("status", ["Open","In Progress"]).execute()
        for plan in (res.data or []):
            due_str = plan.get("due_date",""); owner = plan.get("owner","")
            if not due_str or not owner: continue
            try: due_date = _dt.strptime(due_str[:10], "%Y-%m-%d").date()
            except Exception: continue
            days_left = (due_date - today).days
            try:
                ur = sb.table("profiles").select("id").eq("full_name", owner).execute()
                if not ur.data: continue
                uid = ur.data[0]["id"]
            except Exception: continue
            try:
                today_start = _dt.combine(today, _dt.min.time()).isoformat()
                existing = sb.table("notifications").select("id").eq("user_id", uid).eq("action_plan_id", plan["id"]).gte("created_at", today_start).execute()
                if existing.data: continue
            except Exception: pass
            if days_left < 0:
                create_notification(sb, uid, "🔴 Overdue — QM Action Plan",
                    f'"{plan.get("action","")[:60]}" was due {due_date.strftime("%d %b %Y")} and is still open.',
                    "overdue", plan["id"])
            elif days_left <= 2:
                create_notification(sb, uid, "🟡 Due Soon — QM Action Plan",
                    f'"{plan.get("action","")[:60]}" due {due_date.strftime("%d %b %Y")} ({days_left} day{"s" if days_left!=1 else ""} left).',
                    "warning", plan["id"])
    except Exception: pass

def render_bell(sb, user_id):
    from datetime import datetime as _dt
    count = _get_unread_count(sb, user_id)
    label = f"🔔 {count}" if count > 0 else "🔔"
    with st.popover(label, use_container_width=False):
        st.markdown("### Notifications")
        notifs = _get_notifications(sb, user_id)
        if not notifs:
            st.info("You are all caught up!")
        else:
            c1, c2 = st.columns([3,1])
            c1.caption(f"{count} unread" if count else "All read")
            if count > 0 and c2.button("Mark all read", key="bell_markall"):
                _mark_all_read(sb, user_id); st.rerun()
            st.divider()
            icons = {"overdue":"🔴","warning":"🟡","info":"🔵","success":"🟢"}
            for n in notifs:
                icon = icons.get(n.get("type","info"),"🔵")
                is_read = n.get("is_read", False)
                try: ts = _dt.fromisoformat(n["created_at"].replace("Z","")).strftime("%d %b %H:%M")
                except Exception: ts = ""
                border = "#555" if is_read else "#C1A02E"
                bg     = "#2a2a2a" if is_read else "#1a2a3a"
                op     = "0.5" if is_read else "1"
                fw     = "normal" if is_read else "bold"
                st.markdown(
                    f'<div style="padding:10px;margin-bottom:8px;border-radius:8px;' +
                    f'background:{"#f5f5f5" if is_read else "#ffffff"};opacity:{op};' +
                    f'border-left:4px solid {border};box-shadow:0 1px 3px rgba(0,0,0,0.1);">' +
                    f'<div style="font-weight:{fw};font-size:13px;color:#111111;">{icon} {n["title"]}</div>' +
                    f'<div style="font-size:11px;color:#444444;margin-top:3px;">{n["message"]}</div>' +
                    f'<div style="font-size:10px;color:#888888;margin-top:4px;">{ts}</div></div>',
                    unsafe_allow_html=True
                )
                if not is_read:
                    if st.button("Mark read", key=f"bell_read_{n['id']}"):
                        _mark_one_read(sb, n["id"]); st.rerun()

st.set_page_config(page_title="QM Pillar", page_icon="✅", layout="wide")

if "user" not in st.session_state:
    st.switch_page("app.py")

supabase = get_supabase()
role     = st.session_state.get("role", "member")
pillar   = st.session_state.get("pillar", "ALL")
name     = st.session_state.get("name", "User")
can_edit = (role == "plant_manager") or (role == "pillar_leader" and pillar == "QM")

# Check action plan notifications once per session
if "notif_checked" not in st.session_state:
    try:
        check_action_plan_notifications(supabase)
    except Exception:
        pass
    st.session_state["notif_checked"] = True

col1, col2, col3 = st.columns([5, 1, 1])
with col1:
    st.markdown("# ✅ Quality Maintenance")
    st.markdown(f"Logged in as **{name}** · `{role}` · {'✏️ Full Access' if can_edit else '👁️ View Only'}")
with col2:
    render_bell(supabase, st.session_state["user"])
with col3:
    if st.button("🏠 Home", use_container_width=True):
        st.switch_page("pages/1_Home.py")

st.divider()

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Maturity Level", "📈 Quality Indicators", "📁 Documents", "🎯 Action Plans", "⚙️ Settings",
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
                # Load settings from Supabase
                _crm_map, _q_set, _s_set, _i_set = load_settings_from_supabase(supabase)
                _classifier = make_classifier(_q_set, _s_set, _i_set)

                df_final_loaded  = read_excel_from_upload(final_file)
                df_issued_loaded = read_excel_from_upload(issued_file)
                df_ncr_loaded    = read_excel_from_upload(ncr_file)

                final_pkg  = build_dataset_final_issued(df_final_loaded,  date_col=FINAL_APPROVAL_COL,    dataset_name="FINAL",  crm_delete_map=_crm_map, classifier=_classifier)
                issued_pkg = build_dataset_final_issued(df_issued_loaded, date_col=CREATION_DATETIME_COL, dataset_name="ISSUED", crm_delete_map=_crm_map, classifier=_classifier)
                ncr_pkg    = build_dataset_ncr(df_ncr_loaded,             date_col=FINAL_APPROVAL_COL,    dataset_name="NCR",    classifier=_classifier)

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

            # ── Unknown reasons prompt ──
            _all_unclassified = set()
            for _pkg_name, _pkg in [("FINAL", final_pkg), ("ISSUED", issued_pkg)]:
                for r in _pkg["unclassified_counts"].index:
                    _all_unclassified.add(str(r).strip())
            _all_unclassified = {r for r in _all_unclassified if r and r.lower() not in ("", "nan", "invalid")}

            if _all_unclassified:
                st.warning(f"⚠️ {len(_all_unclassified)} unclassified reason(s) found in uploaded files. Please classify them:")
                with st.form("qm_classify_reasons_form"):
                    _classifications = {}
                    for r in sorted(_all_unclassified):
                        _col1, _col2 = st.columns([3, 1])
                        _col1.markdown(f"**{r}**")
                        _classifications[r] = _col2.radio(
                            "", ["Quality", "Service", "Invalid (exclude)"],
                            key=f"cls_{r[:40]}", horizontal=True
                        )
                    if st.form_submit_button("💾 Save Classifications", type="primary"):
                        for r, cls in _classifications.items():
                            try:
                                if cls == "Quality":
                                    supabase.table("qm_quality_reasons").upsert(
                                        {"reason": r, "added_by": name}, on_conflict="reason").execute()
                                elif cls == "Service":
                                    supabase.table("qm_service_reasons").upsert(
                                        {"reason": r, "added_by": name}, on_conflict="reason").execute()
                                else:
                                    supabase.table("qm_invalid_reasons").upsert(
                                        {"reason": r, "added_by": name}, on_conflict="reason").execute()
                            except Exception as e:
                                st.error(f"Error saving '{r}': {e}")
                        st.success("✅ Classifications saved! Regenerate charts to apply.")
                        st.rerun()

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
                c9, c10 = st.columns([1,7])
                TOPN_RC_REASONS     = topn(c9, "Root Cause Top Reasons", 4, 15)

            run = st.button("🔄 Generate Charts", type="primary", key="qm_run")

            # ── Persistent data forms (outside run block so they survive rerun) ──
            _months_range_pre = list(range(1, selected_month + 1))
            _prev_year_pre    = selected_year - 1
            _MONTH_LABELS_pre = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

            # FG Invoiced missing form
            try:
                _fg_rows = supabase.table("qm_fg_invoiced").select("*").execute()
                _fg_data = {(r["year"], r["month"]): r["fg_invoiced"] for r in (_fg_rows.data or [])}
            except Exception:
                _fg_data = {}
            _missing_fg = [m for m in _months_range_pre if _fg_data.get((selected_year, m)) is None]
            if _missing_fg:
                st.warning(f"⚠️ FG Invoiced missing for: {', '.join([_MONTH_LABELS_pre[m-1]+'-'+str(selected_year)[-2:] for m in _missing_fg])}")
                with st.form("qm_fg_form_persistent"):
                    _fg_inputs = {}
                    _cols = st.columns(min(len(_missing_fg), 6))
                    for i, m in enumerate(_missing_fg):
                        _fg_inputs[m] = _cols[i % 6].number_input(
                            f"{_MONTH_LABELS_pre[m-1]}-{str(selected_year)[-2:]}",
                            min_value=0, value=0, step=1, key=f"pfg_{selected_year}_{m}"
                        )
                    if st.form_submit_button("💾 Save FG Invoiced", type="primary"):
                        _errors = []
                        for m, val in _fg_inputs.items():
                            try:
                                supabase.table("qm_fg_invoiced").upsert({
                                    "year": int(selected_year), "month": int(m),
                                    "fg_invoiced": float(val), "updated_by": name
                                }, on_conflict="year,month").execute()
                            except Exception as e:
                                _errors.append(f"{_MONTH_LABELS_pre[m-1]}: {e}")
                        if _errors:
                            for err in _errors: st.error(err)
                        else:
                            st.success("✅ FG Invoiced saved!"); st.rerun()

            # COQ missing form
            try:
                _coq_rows = supabase.table("qm_coq_data").select("*").execute()
                _coq_data = {(r["year"], r["month"]): r for r in (_coq_rows.data or [])}
            except Exception:
                _coq_data = {}
            _coq_needed  = [(_prev_year_pre, m) for m in _months_range_pre] + [(selected_year, m) for m in _months_range_pre]
            _missing_coq = [(y, m) for y, m in _coq_needed if (y, m) not in _coq_data]
            if _missing_coq:
                st.warning(f"⚠️ COQ data missing for {len(_missing_coq)} month(s):")
                with st.form("qm_coq_form_persistent"):
                    _coq_inputs = {}
                    for y, m in _missing_coq:
                        st.markdown(f"**{_MONTH_LABELS_pre[m-1]}-{str(y)[-2:]}**")
                        c1, c2 = st.columns(2)
                        _coq_inputs[(y,m,"shredding")] = c1.number_input(f"Shredding List", min_value=0.0, value=0.0, step=100.0, key=f"pcoq_shr_{y}_{m}")
                        _coq_inputs[(y,m,"ncr")]       = c2.number_input(f"NCR Shredding",  min_value=0.0, value=0.0, step=100.0, key=f"pcoq_ncr_{y}_{m}")
                    if st.form_submit_button("💾 Save COQ Data", type="primary"):
                        _errors = []
                        for y, m in _missing_coq:
                            try:
                                supabase.table("qm_coq_data").upsert({
                                    "year": int(y), "month": int(m),
                                    "shredding_list": float(_coq_inputs[(y,m,"shredding")]),
                                    "ncr_shredding":  float(_coq_inputs[(y,m,"ncr")]),
                                    "updated_by": name
                                }, on_conflict="year,month").execute()
                            except Exception as e:
                                _errors.append(f"{_MONTH_LABELS_pre[m-1]}-{y}: {e}")
                        if _errors:
                            for err in _errors: st.error(err)
                        else:
                            st.success("✅ COQ data saved!"); st.rerun()

            # WO missing form
            try:
                _wo_rows = supabase.table("qm_wo_data").select("*").execute()
                _wo_data = {(r["year"], r["month"]): r for r in (_wo_rows.data or [])}
            except Exception:
                _wo_data = {}
            _missing_wo = [(selected_year, m) for m in _months_range_pre if (selected_year, m) not in _wo_data]
            if _missing_wo:
                st.warning(f"⚠️ Work Order data missing for {len(_missing_wo)} month(s):")
                with st.form("qm_wo_form_persistent"):
                    _wo_inputs = {}
                    _cols_wo = st.columns(min(4, len(_missing_wo)))
                    for idx, (y, m) in enumerate(_missing_wo):
                        _wo_inputs[(y, m)] = _cols_wo[idx % len(_cols_wo)].number_input(
                            f"{_MONTH_LABELS_pre[m-1]}-{str(y)[-2:]}",
                            min_value=0, value=0, step=10, key=f"pwo_{y}_{m}"
                        )
                    if st.form_submit_button("💾 Save WO Data", type="primary"):
                        _errors = []
                        for y, m in _missing_wo:
                            try:
                                supabase.table("qm_wo_data").upsert({
                                    "year": int(y), "month": int(m),
                                    "wo_count": int(_wo_inputs[(y, m)]),
                                    "updated_by": name
                                }, on_conflict="year,month").execute()
                            except Exception as e:
                                _errors.append(f"{_MONTH_LABELS_pre[m-1]}-{y}: {e}")
                        if _errors:
                            for err in _errors: st.error(err)
                        else:
                            st.success("✅ WO data saved!"); st.rerun()

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

                # ── Quality Root Cause Drill-Down ──
                st.subheader("Quality — Root Cause by Reason (Current Month)")
                def slide_rootcause_fig(category, selected_year, selected_month, top_n_reasons=4):
                    base = df_issued[
                        (df_issued["Is_Valid"] == True) &
                        (df_issued["Complaint_Category"] == category)
                    ].copy()
                    # Auto-detect root cause column
                    rc_col = next((c for c in base.columns if "root" in c.lower() and "cause" in c.lower()), None)
                    if rc_col is None:
                        rc_col = next((c for c in base.columns if "root" in c.lower()), None)
                    if rc_col is None:
                        rc_col = next((c for c in base.columns if "cause" in c.lower()), None)
                    if rc_col is None:
                        return None, f"No root cause column found. Columns: {list(base.columns)}"
                    base["_Reason"] = base["Reason"].astype(str).str.strip()
                    base["_RC"]     = base[rc_col].astype(str).str.strip().replace("nan", "Not Specified")
                    cm = base[(base["Year"] == selected_year) & (base["Month"] == selected_month)]
                    if cm.empty:
                        return None, "No data for selected period."
                    # Top N reasons by count
                    top_reasons = cm["_Reason"].value_counts().head(top_n_reasons).index.tolist()
                    n = len(top_reasons)
                    if n == 0:
                        return None, "No reasons found."
                    # Figure: one subplot per reason, horizontal bars for ALL root causes
                    row_heights = []
                    for r in top_reasons:
                        n_rc = cm[cm["_Reason"] == r]["_RC"].nunique()
                        row_heights.append(max(1.8, n_rc * 0.45 + 0.8))
                    fig_h = sum(row_heights) + 0.5
                    fig, axes = plt.subplots(n, 1, figsize=(13.33, fig_h), dpi=300,
                                             gridspec_kw={"height_ratios": row_heights})
                    if n == 1:
                        axes = [axes]
                    palette = ["#006394","#C1A02E","#0F68B9","#D8C37D","#B7910E",
                               "#4A90D9","#E8A838","#2E6DA4","#8EC6E6","#F5D07A",
                               "#1A4F72","#E09020","#5BA3C9","#C8A830","#3D7EA6"]
                    for ax, reason in zip(axes, top_reasons):
                        subset   = cm[cm["_Reason"] == reason]
                        rc_all   = subset["_RC"].value_counts()  # ALL root causes, no limit
                        reason_n = len(subset)
                        y_pos    = np.arange(len(rc_all))
                        bar_cols = [palette[i % len(palette)] for i in range(len(rc_all))]
                        bars = ax.barh(y_pos, rc_all.values, color=bar_cols, height=0.55, edgecolor="white")
                        # Value labels + % of this reason
                        for bar, val in zip(bars, rc_all.values):
                            pct = val / reason_n * 100
                            ax.text(bar.get_width() + rc_all.max() * 0.02,
                                    bar.get_y() + bar.get_height()/2,
                                    f"{int(val)}  ({pct:.0f}%)",
                                    va="center", ha="left", fontsize=9, color="#000000")
                        ax.set_yticks(y_pos)
                        ax.set_yticklabels([fill(str(r), 50) for r in rc_all.index], fontsize=9)
                        ax.invert_yaxis()
                        ax.set_xlim(0, rc_all.max() * 1.35)
                        ax.set_title(
                            f"  {fill(reason, 80)}   [n={reason_n}]",
                            loc="left", fontsize=10, fontweight="bold", color="#333333", pad=5
                        )
                        ax.spines["top"].set_visible(False)
                        ax.spines["right"].set_visible(False)
                        ax.spines["left"].set_visible(False)
                        ax.tick_params(axis="x", labelsize=8)
                        ax.set_xlabel("Count", fontsize=8)
                        ax.grid(False)
                        ax.axhline(-0.6, color="#EEEEEE", linewidth=0.8)
                    plt.tight_layout(h_pad=1.5)
                    return fig, None

                fig_rc_q, err_rc_q = slide_rootcause_fig("Quality", selected_year, selected_month, TOPN_RC_REASONS or 4)
                if err_rc_q:
                    st.warning(f"Quality root cause: {err_rc_q}")
                else:
                    show_fig(fig_rc_q)
                    slides_for_pdf.append({"title": "ISSUED - Quality Root Cause by Reason",
                        "fig": slide_rootcause_fig("Quality", selected_year, selected_month, TOPN_RC_REASONS or 4)[0]})

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

                # ── Service Root Cause Drill-Down ──
                st.subheader("Service — Root Cause by Reason (Current Month)")
                fig_rc_s, err_rc_s = slide_rootcause_fig("Service", selected_year, selected_month, TOPN_RC_REASONS or 4)
                if err_rc_s:
                    st.warning(f"Service root cause: {err_rc_s}")
                else:
                    show_fig(fig_rc_s)
                    slides_for_pdf.append({"title": "ISSUED - Service Root Cause by Reason",
                        "fig": slide_rootcause_fig("Service", selected_year, selected_month, TOPN_RC_REASONS or 4)[0]})

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
                        submitted = st.form_submit_button("💾 Save FG Invoiced", type="primary")
                        if submitted:
                            errors = []
                            for m, val in fg_inputs.items():
                                try:
                                    supabase.table("qm_fg_invoiced").upsert({
                                        "year": int(selected_year), "month": int(m),
                                        "fg_invoiced": float(val), "updated_by": name
                                    }, on_conflict="year,month").execute()
                                except Exception as e:
                                    errors.append(f"{MONTH_LABELS[m-1]}: {e}")
                            if errors:
                                for err in errors: st.error(err)
                            else:
                                st.success("✅ Saved!"); st.rerun()

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

                        # Both graphs: CC = blue, FG = gold, all text black
                        color_cc = "#006394"
                        color_fg = "#D8C37D"
                        n = len(all_labels)
                        x = np.arange(n)
                        w = 0.35

                        fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)

                        # Side-by-side bars: CC and FG
                        bars_cc = ax.bar(x - w/2, all_cc, width=w, color=color_cc, label="CC Count", alpha=0.9)
                        bars_fg = ax.bar(x + w/2, all_fg, width=w, color=color_fg, label="FG Invoiced", alpha=0.9)

                        # CC count labels above CC bars — black
                        for bar, val in zip(bars_cc, all_cc):
                            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(all_fg)*0.01,
                                    f"{int(val)}", ha="center", va="bottom", fontsize=9,
                                    color="#000000", fontweight="bold")

                        # FG labels above FG bars — black
                        for bar, val in zip(bars_fg, all_fg):
                            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(all_fg)*0.01,
                                    f"{int(val)}", ha="center", va="bottom", fontsize=9, color="#000000")

                        # ── Secondary axis for ratio line ──
                        ax2 = ax.twinx()

                        # Plot ratio line on secondary axis (monthly only, not YTD)
                        ax2.plot(x[:-1], ratios, color="#C1A02E", linewidth=2.5,
                                 marker="o", markersize=6, label="Ratio %", zorder=5)

                        # Trendline on secondary axis
                        if len(ratios) >= 2:
                            xf = np.arange(len(ratios), dtype=float)
                            coeff = np.polyfit(xf, ratios, 1)
                            ax2.plot(x[:-1], np.polyval(coeff, xf), linestyle="--",
                                     linewidth=1.5, color="#999999", label="Trend", zorder=4)

                        # Ratio % labels above each point — black
                        for i, ratio in enumerate(ratios):
                            ax2.text(x[i], ratio + max(ratios)*0.05, f"{ratio:.2f}%",
                                     ha="center", va="bottom", fontsize=10,
                                     color="#000000", fontweight="bold")

                        # YTD ratio label — black
                        ax2.text(x[-1], ytd_ratio + max(ratios)*0.05, f"{ytd_ratio:.2f}%",
                                 ha="center", va="bottom", fontsize=10,
                                 color="#000000", fontweight="bold")

                        ax2.set_ylabel("CC Ratio (%)", color="#000000")
                        ax2.tick_params(axis="y", labelcolor="#000000")
                        ax2.set_ylim(0, max(ratios + [ytd_ratio, 1]) * 1.5)
                        ax2.spines["top"].set_visible(False)

                        ax.set_xticks(x)
                        ax.set_xticklabels(all_labels, fontsize=11)
                        ax.set_ylabel("Count")
                        ax.set_ylim(0, max(all_fg) * 1.20)
                        ax.spines["top"].set_visible(False)
                        ax.spines["right"].set_visible(False)
                        ax.grid(False)

                        # Combined legend
                        lines1, labs1 = ax.get_legend_handles_labels()
                        lines2, labs2 = ax2.get_legend_handles_labels()
                        ax.legend(lines1+lines2, labs1+labs2, loc="upper center",
                                  bbox_to_anchor=(0.5, -0.08), ncol=4, frameon=False, fontsize=10)
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

                st.divider()

                # ── COST OF QUALITY ──
                st.header("Cost of Quality (COQ)")

                REWORK_FIXED = 8250.0

                # Load COQ data from Supabase
                try:
                    coq_rows = supabase.table("qm_coq_data").select("*").execute()
                    coq_data = {(r["year"], r["month"]): r for r in (coq_rows.data or [])}
                except Exception:
                    coq_data = {}

                # CN @Cost from df_final (Valid=True, Quality, Credit Note)
                def get_cn_cost(year, month):
                    try:
                        decision_col = next((c for c in df_final.columns if "decision" in c.lower()), None)
                        if not decision_col: return 0.0
                        base = df_final[
                            (df_final["Is_Valid"] == True) &
                            (df_final["Complaint_Category"] == "Quality") &
                            (df_final["Year"] == year) &
                            (df_final["Month"] == month)
                        ].copy()
                        base[decision_col] = base[decision_col].astype(str).str.strip().str.lower()
                        base = base[base[decision_col] == "credit note"]
                        base["Cost Amount"] = pd.to_numeric(base["Cost Amount"], errors="coerce").fillna(0)
                        return float(base["Cost Amount"].sum())
                    except Exception:
                        return 0.0

                # Check missing months for both years in range
                prev_year_coq = selected_year - 1
                coq_months_needed = [(prev_year_coq, m) for m in months_range] + [(selected_year, m) for m in months_range]
                missing_coq = [(y, m) for y, m in coq_months_needed if (y, m) not in coq_data]

                if missing_coq:
                    st.warning(f"⚠️ COQ data missing for {len(missing_coq)} month(s). Please fill in:")
                    with st.form("qm_coq_form"):
                        coq_inputs = {}
                        for y, m in missing_coq:
                            st.markdown(f"**{MONTH_LABELS[m-1]}-{str(y)[-2:]}**")
                            c1, c2 = st.columns(2)
                            coq_inputs[(y,m,"shredding")] = c1.number_input(f"Shredding List", min_value=0.0, value=0.0, step=100.0, key=f"coq_shr_{y}_{m}")
                            coq_inputs[(y,m,"ncr")]       = c2.number_input(f"NCR Shredding",  min_value=0.0, value=0.0, step=100.0, key=f"coq_ncr_{y}_{m}")
                        if st.form_submit_button("💾 Save COQ Data", type="primary"):
                            errors = []
                            for y, m in missing_coq:
                                try:
                                    supabase.table("qm_coq_data").upsert({
                                        "year": int(y), "month": int(m),
                                        "shredding_list": float(coq_inputs[(y,m,"shredding")]),
                                        "ncr_shredding":  float(coq_inputs[(y,m,"ncr")]),
                                        "updated_by": name
                                    }, on_conflict="year,month").execute()
                                except Exception as e:
                                    errors.append(f"{MONTH_LABELS[m-1]}-{y}: {e}")
                            if errors:
                                for err in errors: st.error(err)
                            else:
                                st.success("✅ Saved!"); st.rerun()

                # Build COQ arrays for chart
                def get_coq_month(year, month):
                    row = coq_data.get((year, month), {})
                    shred  = float(row.get("shredding_list", 0) or 0)
                    ncr    = float(row.get("ncr_shredding",  0) or 0)
                    rework = REWORK_FIXED
                    cn     = get_cn_cost(year, month)
                    total  = shred + ncr + rework + cn
                    return shred, ncr, rework, cn, total

                # Check we have data for all needed months
                coq_ready = all((y,m) in coq_data for y,m in coq_months_needed)

                if coq_ready:
                    METRIC_LABELS = ["Shredding List", "NCR Shredding", "Rework Labor", "CN @Cost", "MCOQ Total"]
                    COLOR_CM_PREV  = "#D8C37D"
                    COLOR_CM_SEL   = "#0F68B9"
                    COLOR_YTD_PREV = "#B7910E"
                    COLOR_YTD_SEL  = "#006394"

                    def fmt_sar(v):
                        if v >= 1_000_000: return f"SAR {v/1_000_000:.1f}M"
                        if v >= 1000:      return f"SAR {int(v/1000)}K"
                        return f"SAR {int(v)}"

                    def slide_coq_breakdown_fig(selected_year, selected_month):
                        prev_year = selected_year - 1

                        # CM values
                        cm_prev_vals = list(get_coq_month(prev_year, selected_month))
                        cm_sel_vals  = list(get_coq_month(selected_year, selected_month))

                        # YTD values
                        ytd_prev = [0.0]*5
                        ytd_sel  = [0.0]*5
                        for m in months_range:
                            pp = list(get_coq_month(prev_year, m))
                            ss = list(get_coq_month(selected_year, m))
                            for i in range(5):
                                ytd_prev[i] += pp[i]
                                ytd_sel[i]  += ss[i]

                        cm_prev_lbl  = pd.to_datetime(f"{prev_year}-{selected_month:02d}-01").strftime("%b-%y")
                        cm_sel_lbl   = pd.to_datetime(f"{selected_year}-{selected_month:02d}-01").strftime("%b-%y")

                        n_groups = 5
                        x = np.arange(n_groups)
                        w = 0.18

                        fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)

                        b1 = ax.bar(x - 1.5*w, cm_prev_vals,  w, color=COLOR_CM_PREV,  label=f"CM {cm_prev_lbl}")
                        b2 = ax.bar(x - 0.5*w, cm_sel_vals,   w, color=COLOR_CM_SEL,   label=f"CM {cm_sel_lbl}")
                        b3 = ax.bar(x + 0.5*w, ytd_prev,      w, color=COLOR_YTD_PREV, label=f"YTD {prev_year}")
                        b4 = ax.bar(x + 1.5*w, ytd_sel,       w, color=COLOR_YTD_SEL,  label=f"YTD {selected_year}")

                        for bars_g in [b1, b2, b3, b4]:
                            for bar in bars_g:
                                h = bar.get_height()
                                if h > 0:
                                    ax.text(bar.get_x() + bar.get_width()/2, h + ax.get_ylim()[1]*0.01,
                                            fmt_sar(h), ha="center", va="bottom", fontsize=8,
                                            rotation=90, color="#000000", clip_on=False)

                        ax.set_xticks(x)
                        ax.set_xticklabels(METRIC_LABELS, fontsize=11)
                        ax.set_ylabel("Cost (SAR)")
                        ax.spines["top"].set_visible(False)
                        ax.spines["right"].set_visible(False)
                        ax.grid(False)
                        ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.08), ncol=4, frameon=False, fontsize=10)
                        fig.subplots_adjust(bottom=0.20)
                        return fig

                    st.subheader("COQ Breakdown — CM vs YTD")
                    show_fig(slide_coq_breakdown_fig(selected_year, selected_month))
                    slides_for_pdf.append({"title": "Cost of Quality — Breakdown", "fig": slide_coq_breakdown_fig(selected_year, selected_month)})

                else:
                    st.info("Fill in missing COQ data above to generate COQ charts.")

                st.divider()

                # ── NCR RATIO ──
                st.header("NCR Ratio")

                # Load WO data from Supabase
                try:
                    wo_rows = supabase.table("qm_wo_data").select("*").execute()
                    wo_data = {(r["year"], r["month"]): r for r in (wo_rows.data or [])}
                except Exception:
                    wo_data = {}

                # Get valid NCR count per month from df_ncr
                def get_ncr_count(year, month):
                    try:
                        return int(df_ncr_dash[
                            (df_ncr_dash["Is_Valid"] == True) &
                            (df_ncr_dash["Year"] == year) &
                            (df_ncr_dash["Month"] == month)
                        ].shape[0])
                    except Exception:
                        return 0

                # Check missing WO months — current year only
                wo_months_needed = [(selected_year, m) for m in months_range]
                missing_wo = [(y, m) for y, m in wo_months_needed if (y, m) not in wo_data]

                if missing_wo:
                    st.warning(f"⚠️ Work Order data missing for {len(missing_wo)} month(s). Please fill in:")
                    with st.form("qm_wo_form"):
                        wo_inputs = {}
                        cols_wo = st.columns(min(4, len(missing_wo)))
                        for idx, (y, m) in enumerate(missing_wo):
                            col = cols_wo[idx % len(cols_wo)]
                            wo_inputs[(y, m)] = col.number_input(
                                f"{MONTH_LABELS[m-1]}-{str(y)[-2:]}",
                                min_value=0, value=0, step=10,
                                key=f"wo_{y}_{m}"
                            )
                        if st.form_submit_button("💾 Save WO Data", type="primary"):
                            errors = []
                            for y, m in missing_wo:
                                try:
                                    supabase.table("qm_wo_data").upsert({
                                        "year": int(y), "month": int(m),
                                        "wo_count": int(wo_inputs[(y, m)]),
                                        "updated_by": name
                                    }, on_conflict="year,month").execute()
                                except Exception as e:
                                    errors.append(f"{MONTH_LABELS[m-1]}-{y}: {e}")
                            if errors:
                                for err in errors: st.error(err)
                            else:
                                st.success("✅ Saved!"); st.rerun()

                wo_ready = all((y, m) in wo_data for y, m in wo_months_needed)

                if wo_ready:
                    def slide_ncr_ratio_fig(selected_year):
                        all_labels = [MONTH_LABELS[m-1] for m in months_range] + ["YTD"]

                        # Monthly NCR and WO counts — current year only
                        ncr_months = [get_ncr_count(selected_year, m) for m in months_range]
                        wo_months  = [int(wo_data.get((selected_year, m), {}).get("wo_count", 0) or 0) for m in months_range]

                        # YTD totals
                        ytd_ncr = sum(ncr_months)
                        ytd_wo  = sum(wo_months)

                        all_ncr = ncr_months + [ytd_ncr]
                        all_wo  = wo_months  + [ytd_wo]

                        # Monthly ratios
                        ratios = [ncr/wo*100 if wo > 0 else 0 for ncr, wo in zip(ncr_months, wo_months)]
                        ytd_ratio = ytd_ncr / ytd_wo * 100 if ytd_wo > 0 else 0

                        n = len(all_labels)
                        x = np.arange(n)
                        w = 0.35

                        fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)

                        # Side-by-side bars: NCR and WO
                        bars_ncr = ax.bar(x - w/2, all_ncr, w, color="#006394", label="NCR Count", alpha=0.9)
                        bars_wo  = ax.bar(x + w/2, all_wo,  w, color="#D8C37D", label="Work Orders", alpha=0.9)

                        # Value labels
                        for bar, val in zip(bars_ncr, all_ncr):
                            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(all_wo)*0.01,
                                    f"{int(val)}", ha="center", va="bottom", fontsize=9,
                                    color="#000000", fontweight="bold")
                        for bar, val in zip(bars_wo, all_wo):
                            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(all_wo)*0.01,
                                    f"{int(val)}", ha="center", va="bottom", fontsize=9, color="#000000")

                        # Secondary axis — ratio line
                        ax2 = ax.twinx()
                        ax2.plot(x[:-1], ratios, color="#C1A02E", linewidth=2.5,
                                 marker="o", markersize=6, label="NCR Ratio %", zorder=5)

                        # Trendline
                        if len(ratios) >= 2:
                            xf = np.arange(len(ratios), dtype=float)
                            ax2.plot(x[:-1], np.polyval(np.polyfit(xf, ratios, 1), xf),
                                     linestyle="--", linewidth=1.5, color="#999999", zorder=4)

                        # Ratio % labels
                        max_r = max(ratios + [ytd_ratio, 1])
                        for i, r in enumerate(ratios):
                            ax2.text(x[i], r + max_r*0.05, f"{r:.2f}%",
                                     ha="center", va="bottom", fontsize=10,
                                     color="#000000", fontweight="bold")

                        # YTD ratio label
                        ax2.text(x[-1], ytd_ratio + max_r*0.05, f"{ytd_ratio:.2f}%",
                                 ha="center", va="bottom", fontsize=10,
                                 color="#000000", fontweight="bold")

                        ax2.set_ylabel("NCR Ratio (%)", color="#000000")
                        ax2.tick_params(axis="y", labelcolor="#000000")
                        ax2.set_ylim(0, max_r * 1.5)
                        ax2.spines["top"].set_visible(False)

                        ax.set_xticks(x)
                        ax.set_xticklabels(all_labels, fontsize=10)
                        ax.set_ylabel("Count", color="#000000")
                        ax.spines["top"].set_visible(False)
                        ax.grid(False)

                        lines1, labs1 = ax.get_legend_handles_labels()
                        lines2, labs2 = ax2.get_legend_handles_labels()
                        ax.legend(lines1 + lines2, labs1 + labs2,
                                  loc="upper center", bbox_to_anchor=(0.5, -0.08),
                                  ncol=3, frameon=False, fontsize=10)
                        fig.subplots_adjust(bottom=0.15)
                        return fig

                    st.subheader("NCR Ratio")
                    show_fig(slide_ncr_ratio_fig(selected_year))
                    slides_for_pdf.append({"title": "NCR Ratio",
                                           "fig": slide_ncr_ratio_fig(selected_year)})
                else:
                    st.info("Fill in missing Work Order data above to generate NCR Ratio chart.")

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
        # Load people directory for owner dropdown
        try:
            people_res = supabase.table("profiles").select("full_name").order("full_name").execute()
            people_names = [p["full_name"] for p in (people_res.data or []) if p.get("full_name")]
        except Exception as e:
            st.warning(f"Could not load users: {e}")
            people_names = []
        act_text=st.text_input("Action",key="qm_act_text")
        if people_names:
            own_text = st.selectbox("Owner", ["— select —"] + people_names, key="qm_own_sel")
            own_text = own_text if own_text != "— select —" else ""
        else:
            st.caption("⚠️ No people found in directory — go to Main Hub → People Directory to add them.")
            own_text=st.text_input("Owner",key="qm_own_text")
        due_d=st.date_input("Due Date",key="qm_due_date")
        stat_ap=st.selectbox("Status",["Open","In Progress","Done"],key="qm_new_act_status")
        notes_ap=st.text_area("Notes",key="qm_notes_ap")
        if st.button("➕ Add Action Plan",type="primary",key="qm_add_act"):
            if not act_text or not own_text: st.error("Please fill in action and owner.")
            else:
                try:
                    supabase.table("qm_action_plans").insert({"action":act_text,"owner":own_text,"due_date":str(due_d),"status":stat_ap,"notes":notes_ap}).execute()
                    # Notify the assigned owner
                    try:
                        owner_res = supabase.table("profiles").select("id").eq("full_name", own_text).execute()
                        if owner_res.data:
                            create_notification(
                                supabase, owner_res.data[0]["id"],
                                title="📋 New Action Plan Assigned — QM",
                                message=f'"{act_text[:60]}" assigned to you. Due: {due_d.strftime("%d %b %Y")}.',
                                notif_type="info"
                            )
                    except Exception:
                        pass
                    st.success("✅ Action plan added!"); st.rerun()
                except Exception as e: st.error(f"Error: {str(e)}")


# ════════════════════════════════════════
# TAB 5 — SETTINGS
# ════════════════════════════════════════
with tab5:
    st.markdown("### ⚙️ QM Settings")
    if not can_edit:
        st.warning("🔒 Only Pillar Leader or Plant Manager can manage settings.")
        st.stop()

    set1, set2, set3 = st.tabs(["🗑️ CRM Delete Map", "📋 Reasons", "🔍 Preview"])

    # ── LOAD CURRENT SETTINGS ──
    try:
        crm_rows_s = supabase.table("qm_crm_delete_map").select("*").order("added_at", desc=True).execute()
        crm_list   = crm_rows_s.data or []
    except Exception: crm_list = []
    try:
        q_rows_s = supabase.table("qm_quality_reasons").select("*").order("reason").execute()
        q_list   = q_rows_s.data or []
    except Exception: q_list = []
    try:
        s_rows_s = supabase.table("qm_service_reasons").select("*").order("reason").execute()
        s_list   = s_rows_s.data or []
    except Exception: s_list = []
    try:
        i_rows_s = supabase.table("qm_invalid_reasons").select("*").order("reason").execute()
        i_list   = i_rows_s.data or []
    except Exception: i_list = []

    # ════════════════════════════════════
    # SETTINGS TAB 1 — CRM DELETE MAP
    # ════════════════════════════════════
    with set1:
        st.markdown("#### 🗑️ Permanent CRM Delete Log")
        st.caption("CRM entries here are permanently removed from every upload. Rows shown = how many rows are KEPT (not deleted).")

        if crm_list:
            crm_df = pd.DataFrame(crm_list)[["crm_ref","customer_name","kept_count","added_by","added_at"]].rename(columns={
                "crm_ref": "CRM #", "customer_name": "Customer",
                "kept_count": "Rows Kept", "added_by": "Added By", "added_at": "Added At"
            })
            crm_df["Added At"] = pd.to_datetime(crm_df["Added At"]).dt.strftime("%d %b %Y")
            st.dataframe(crm_df, use_container_width=True, hide_index=True)
        else:
            st.info("No CRM deletions recorded yet.")

        st.divider()
        st.markdown("#### ➕ Add / Update CRM Deletion")
        st.caption("Upload a CRM file to look up customer names, then enter the CRM # and how many rows to KEEP.")

        crm_lookup_file = st.file_uploader("Upload CRM file (for customer lookup)", type=["xlsx","xls"], key="crm_lookup")

        with st.form("qm_crm_add_form"):
            c1, c2 = st.columns(2)
            new_crm_ref   = c1.text_input("CRM Reference #", placeholder="EPAK-CRM-XXXXX").strip()
            new_kept_count = c2.number_input("Rows to KEEP (0 = delete all)", min_value=0, value=0, step=1)

            # Auto-lookup customer from uploaded file
            customer_display = ""
            if crm_lookup_file and new_crm_ref:
                try:
                    import io as _io
                    _df_lookup = pd.read_excel(_io.BytesIO(crm_lookup_file.getvalue()))
                    _match = _df_lookup[_df_lookup["Ref NB"].astype(str).str.strip() == new_crm_ref]
                    if not _match.empty and "Customer" in _df_lookup.columns:
                        customer_display = str(_match["Customer"].iloc[0]).strip()
                except Exception:
                    pass

            new_customer = st.text_input("Customer Name", value=customer_display,
                                          placeholder="Auto-filled from file or type manually")

            if st.form_submit_button("💾 Save CRM Entry", type="primary"):
                if not new_crm_ref:
                    st.error("Please enter a CRM reference number.")
                else:
                    try:
                        supabase.table("qm_crm_delete_map").upsert({
                            "crm_ref": new_crm_ref,
                            "customer_name": new_customer or "",
                            "kept_count": int(new_kept_count),
                            "added_by": name,
                        }, on_conflict="crm_ref").execute()
                        st.success(f"✅ {new_crm_ref} saved — keeping {new_kept_count} rows.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")

        if crm_list:
            st.divider()
            st.markdown("#### 🗑️ Remove from Delete Map")
            st.caption("Removing a CRM from this list means its rows will NO longer be deleted on upload.")
            del_crm_map = {r["crm_ref"]: r["id"] for r in crm_list}
            sel_del_crm = st.selectbox("Select CRM to remove", list(del_crm_map.keys()), key="del_crm_sel")
            if st.button("🗑️ Remove CRM Entry", key="btn_del_crm"):
                try:
                    supabase.table("qm_crm_delete_map").delete().eq("id", del_crm_map[sel_del_crm]).execute()
                    st.success(f"Removed {sel_del_crm}"); st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")

    # ════════════════════════════════════
    # SETTINGS TAB 2 — REASONS
    # ════════════════════════════════════
    with set2:
        col_q, col_s, col_i = st.columns(3)

        # Quality Reasons
        with col_q:
            st.markdown(f"#### ✅ Quality Reasons ({len(q_list)})")
            if q_list:
                st.dataframe(pd.DataFrame(q_list)[["reason"]].rename(columns={"reason":"Reason"}),
                             use_container_width=True, hide_index=True, height=300)
            with st.form("qm_add_q_reason"):
                new_q = st.text_input("Add Quality Reason", key="new_q_reason")
                if st.form_submit_button("➕ Add", type="primary"):
                    if new_q.strip():
                        try:
                            supabase.table("qm_quality_reasons").upsert(
                                {"reason": new_q.strip(), "added_by": name}, on_conflict="reason").execute()
                            st.success("✅ Added!"); st.rerun()
                        except Exception as e: st.error(str(e))
            if q_list:
                del_q = st.selectbox("Remove", [r["reason"] for r in q_list], key="del_q_sel")
                if st.button("🗑️ Remove", key="btn_del_q"):
                    try:
                        supabase.table("qm_quality_reasons").delete().eq("reason", del_q).execute()
                        st.rerun()
                    except Exception as e: st.error(str(e))

        # Service Reasons
        with col_s:
            st.markdown(f"#### 🚚 Service Reasons ({len(s_list)})")
            if s_list:
                st.dataframe(pd.DataFrame(s_list)[["reason"]].rename(columns={"reason":"Reason"}),
                             use_container_width=True, hide_index=True, height=300)
            with st.form("qm_add_s_reason"):
                new_s = st.text_input("Add Service Reason", key="new_s_reason")
                if st.form_submit_button("➕ Add", type="primary"):
                    if new_s.strip():
                        try:
                            supabase.table("qm_service_reasons").upsert(
                                {"reason": new_s.strip(), "added_by": name}, on_conflict="reason").execute()
                            st.success("✅ Added!"); st.rerun()
                        except Exception as e: st.error(str(e))
            if s_list:
                del_s = st.selectbox("Remove", [r["reason"] for r in s_list], key="del_s_sel")
                if st.button("🗑️ Remove", key="btn_del_s"):
                    try:
                        supabase.table("qm_service_reasons").delete().eq("reason", del_s).execute()
                        st.rerun()
                    except Exception as e: st.error(str(e))

        # Invalid Reasons
        with col_i:
            st.markdown(f"#### ❌ Invalid / Excluded Reasons ({len(i_list)})")
            if i_list:
                st.dataframe(pd.DataFrame(i_list)[["reason"]].rename(columns={"reason":"Reason"}),
                             use_container_width=True, hide_index=True, height=300)
            with st.form("qm_add_i_reason"):
                new_i = st.text_input("Add Invalid Reason", key="new_i_reason")
                if st.form_submit_button("➕ Add", type="primary"):
                    if new_i.strip():
                        try:
                            supabase.table("qm_invalid_reasons").upsert(
                                {"reason": new_i.strip(), "added_by": name}, on_conflict="reason").execute()
                            st.success("✅ Added!"); st.rerun()
                        except Exception as e: st.error(str(e))
            if i_list:
                del_i = st.selectbox("Remove", [r["reason"] for r in i_list], key="del_i_sel")
                if st.button("🗑️ Remove", key="btn_del_i"):
                    try:
                        supabase.table("qm_invalid_reasons").delete().eq("reason", del_i).execute()
                        st.rerun()
                    except Exception as e: st.error(str(e))

    # ════════════════════════════════════
    # SETTINGS TAB 3 — PREVIEW
    # ════════════════════════════════════
    with set3:
        st.markdown("#### 🔍 Settings Summary")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("CRM Deletions", len(crm_list))
        c2.metric("Quality Reasons", len(q_list))
        c3.metric("Service Reasons", len(s_list))
        c4.metric("Invalid Reasons", len(i_list))
        st.info("Changes take effect the next time you upload and generate charts in the Quality Indicators tab.")
