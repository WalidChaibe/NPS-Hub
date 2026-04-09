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

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Maturity Level", "📈 Quality Indicators", "📁 Documents", "🎯 Action Plans", "⚙️ Settings", "🔬 Quality Matrix",
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
        # Title text
        c.setFillColor(HexColor("#0E5E86"))
        c.setFont("Helvetica-Bold", 22)
        c.drawString(40, H - 52, title)
        # Red short line + Blue long line at same position as old blue bar
        c.setFillColor(HexColor("#DE201B"))
        c.rect(40, H - 68, 110, 4, fill=1, stroke=0)
        c.setFillColor(HexColor("#0C5595"))
        c.rect(155, H - 68, W - 195, 4, fill=1, stroke=0)

    # Napco logo — embedded as base64 for reliability
    _napco_logo_reader = None
    try:
        import base64 as _b64, tempfile as _tf
        _logo_b64 = "/9j/4AAQSkZJRgABAQEA3ADcAAD/2wBDAAIBAQEBAQIBAQECAgICAgQDAgICAgUEBAMEBgUGBgYFBgYGBwkIBgcJBwYGCAsICQoKCgoKBggLDAsKDAkKCgr/2wBDAQICAgICAgUDAwUKBwYHCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgr/wAARCAHoBFkDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9+KKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAoooo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"
        _logo_bytes = _b64.b64decode(_logo_b64)
        _tmp = _tf.NamedTemporaryFile(delete=False, suffix=".jpeg")
        _tmp.write(_logo_bytes); _tmp.flush()
        _napco_logo_reader = ImageReader(_tmp.name)
    except Exception:
        _napco_logo_reader = None

    def _draw_separator_lines(c, W, y):
        # Red short line on left
        c.setFillColor(HexColor("#DE201B"))
        c.rect(40, y, 110, 4, fill=1, stroke=0)
        # Blue long line continuing right
        c.setFillColor(HexColor("#0C5595"))
        c.rect(155, y, W - 195, 4, fill=1, stroke=0)

    def draw_cover_slide(c, W, H, title_text, subtitle_text, date_text):
        # White background
        c.setFillColorRGB(1, 1, 1)
        c.rect(0, 0, W, H, fill=1, stroke=0)
        # Logo top left — same position as original slide
        if _napco_logo_reader:
            try:
                c.drawImage(_napco_logo_reader, 40, H - 170, width=320, height=150,
                            preserveAspectRatio=True, mask="auto")
            except Exception:
                pass
        # Red + Blue separator lines — just below logo area
        _draw_separator_lines(c, W, H - 185)
        # Title — below lines
        c.setFillColor(HexColor("#0E5E86"))
        c.setFont("Helvetica-Bold", 44)
        title_w = c.stringWidth(title_text, "Helvetica-Bold", 44)
        title_y = H - 300
        c.drawString((W - title_w) / 2, title_y, title_text)
        # Grey horizontal line between title and subtitle
        c.setStrokeColor(HexColor("#AAAAAA"))
        c.setLineWidth(1)
        c.line(80, title_y - 18, W - 80, title_y - 18)
        # Subtitle — below grey line
        c.setFillColor(HexColor("#555555"))
        c.setFont("Helvetica-Oblique", 22)
        sub_w = c.stringWidth(subtitle_text, "Helvetica-Oblique", 22)
        c.drawString((W - sub_w) / 2, title_y - 52, subtitle_text)
        # Date bottom right
        c.setFillColor(HexColor("#888888"))
        c.setFont("Helvetica-Oblique", 14)
        date_w = c.stringWidth(date_text, "Helvetica-Oblique", 14)
        c.drawString(W - date_w - 40, 28, date_text)

    def draw_section_slide(c, W, H, section_text):
        # White background
        c.setFillColorRGB(1, 1, 1)
        c.rect(0, 0, W, H, fill=1, stroke=0)
        # Red + Blue separator lines — upper third
        _draw_separator_lines(c, W, H * 0.58)
        # Section title centered — below lines
        c.setFillColor(HexColor("#0E5E86"))
        c.setFont("Helvetica-Bold", 40)
        txt_w = c.stringWidth(section_text, "Helvetica-Bold", 40)
        c.drawString((W - txt_w) / 2, H * 0.58 - 70, section_text)

    def build_ppt_pdf(slides, dpi=300):
        W, H = 960, 540
        pdf_buf = io.BytesIO()
        c = canvas.Canvas(pdf_buf, pagesize=(W, H))
        for s in slides:
            if s.get("cover"):
                draw_cover_slide(c, W, H,
                    s.get("title", "Quality Indicators Report"),
                    s.get("subtitle", ""),
                    s.get("date", ""))
            elif s.get("section"):
                draw_section_slide(c, W, H, s.get("title", ""))
            else:
                ppt_slide_title_bar(c, s["title"], W, H)
                img = ImageReader(fig_to_png_bytes(s["fig"], dpi=dpi))
                c.drawImage(img, 40, 40, width=W-80, height=H-92-40, preserveAspectRatio=True, anchor="c")
                plt.close(s["fig"])
            c.showPage()
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
                ncr_pkg    = build_dataset_ncr(df_ncr_loaded,             date_col=CREATION_DATETIME_COL, dataset_name="NCR",    classifier=_classifier)

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
                # Store for trend explorer (persists across reruns)
                st.session_state["qm_df_final"]       = df_final
                st.session_state["qm_selected_year"]  = selected_year
                st.session_state["qm_selected_month"] = selected_month

                # ── Always re-fetch settings + rebuild datasets on Generate ──
                _crm_map, _q_set, _s_set, _i_set = load_settings_from_supabase(supabase)
                _classifier = make_classifier(_q_set, _s_set, _i_set)
                final_pkg  = build_dataset_final_issued(df_final_loaded,  date_col=FINAL_APPROVAL_COL,    dataset_name="FINAL",  crm_delete_map=_crm_map, classifier=_classifier)
                issued_pkg = build_dataset_final_issued(df_issued_loaded, date_col=CREATION_DATETIME_COL, dataset_name="ISSUED", crm_delete_map=_crm_map, classifier=_classifier)
                ncr_pkg    = build_dataset_ncr(df_ncr_loaded,             date_col=CREATION_DATETIME_COL, dataset_name="NCR",    classifier=_classifier)
                df_final   = final_pkg["cleaned_flagged"].copy()
                df_issued  = issued_pkg["cleaned_flagged"].copy()
                df_ncr     = ncr_pkg["cleaned_flagged"].copy()
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
                df             = df_final
                df_raw_flagged = df_final_raw_flagged
                df_ncr_dash    = df_ncr

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

                # ── Export processed data ──
                st.subheader("📥 Export Processed Data")
                st.caption("Download processed files with pipeline-added columns: Year, Month, Is_Valid, Type. Use to verify against manual data.")
                _exp_c1, _exp_c2 = st.columns(2)

                _final_exp_cols = [c for c in df_final.columns if c not in ["Base_Date"]]
                _final_exp_df = final_pkg["cleaned_flagged"][_final_exp_cols].copy().rename(
                    columns={"Complaint_Category":"Type_Pipeline","Is_Valid":"Valid_Pipeline"})
                _buf_final = io.BytesIO()
                _final_exp_df.to_excel(_buf_final, index=False, engine="openpyxl")
                _buf_final.seek(0)
                _exp_c1.download_button(
                    label="📥 Download Processed FINAL file",
                    data=_buf_final,
                    file_name=f"FINAL_processed_{selected_year}_{selected_month:02d}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_final_processed"
                )

                _issued_exp_cols = [c for c in df_issued.columns if c not in ["Base_Date"]]
                _issued_exp_df = issued_pkg["cleaned_flagged"][_issued_exp_cols].copy().rename(
                    columns={"Complaint_Category":"Type_Pipeline","Is_Valid":"Valid_Pipeline"})
                _buf_issued = io.BytesIO()
                _issued_exp_df.to_excel(_buf_issued, index=False, engine="openpyxl")
                _buf_issued.seek(0)
                _exp_c2.download_button(
                    label="📥 Download Processed ISSUED file",
                    data=_buf_issued,
                    file_name=f"ISSUED_processed_{selected_year}_{selected_month:02d}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_issued_processed"
                )
                st.divider()

                # ── FINAL ──
                st.header("FINAL (CRM Approved)")

                # Warn about UNCLASSIFIED valid rows
                _unc_final = final_pkg["unclassified_counts"]
                _unc_issued = issued_pkg["unclassified_counts"]
                if not _unc_final.empty:
                    _unc_total = int(_unc_final.sum())
                    with st.expander(f"⚠️ {_unc_total} valid FINAL rows are UNCLASSIFIED and excluded from charts — click to see", expanded=True):
                        st.caption("These rows pass Is_Valid=True but their reason is not in Quality or Service lists. They are NOT counted in any chart. Classify them in the prompt above.")
                        st.dataframe(_unc_final.reset_index().rename(columns={"index":"Reason", 0:"Count", "Reason":"Reason", _unc_final.name if hasattr(_unc_final,'name') else 0:"Count"}), use_container_width=True, hide_index=True)

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

                # ── Service CC Ratio (slide 6) ──
                def slide_service_cc_ratio_fig(selected_year):
                    return slide_cc_ratio_fig(selected_year, category_filter="Service", title_label="Service CC Ratio")

                # ── Service CN Cost (slide 23) ──
                def slide_service_cn_cost_fig(selected_year, selected_month, top_n=15):
                    decision_candidates = [c for c in df_raw_flagged.columns if "decision" in str(c).lower()]
                    if not decision_candidates: raise KeyError("Could not find Decision column.")
                    DECISION_COL = decision_candidates[0]
                    base = df_raw_flagged[(df_raw_flagged["Is_Valid"]==True)&(df_raw_flagged["Complaint_Category"]=="Service")].copy()
                    base[DECISION_COL] = base[DECISION_COL].astype(str).str.strip()
                    base = base[base[DECISION_COL].str.lower()=="credit note"]
                    base["Reason_S"] = base["Reason"].astype(str).str.strip()
                    base["Cost Amount"] = pd.to_numeric(base["Cost Amount"], errors="coerce").fillna(0)
                    cp = base[(base["Year"]==prev_year)&(base["Month"]==selected_month)].groupby("Reason_S")["Cost Amount"].sum()
                    cs = base[(base["Year"]==selected_year)&(base["Month"]==selected_month)].groupby("Reason_S")["Cost Amount"].sum()
                    yp = base[(base["Year"]==prev_year)&(base["Month"].between(1,selected_month))].groupby("Reason_S")["Cost Amount"].sum()
                    ys = base[(base["Year"]==selected_year)&(base["Month"].between(1,selected_month))].groupby("Reason_S")["Cost Amount"].sum()
                    summ = pd.DataFrame({"CM_prev":cp,"CM_sel":cs,"YTD_prev":yp,"YTD_sel":ys}).fillna(0)
                    summ = summ.sort_values(["YTD_sel","CM_sel"],ascending=False)
                    if top_n: summ = summ.head(int(top_n))
                    x = np.arange(len(summ)); w = 0.18
                    fig,ax = plt.subplots(figsize=(13.33,7.5),dpi=300)
                    b1=ax.bar(x-1.5*w,summ["CM_prev"],w,label=f"CM {prev_year}",color="#0F68B9")
                    b2=ax.bar(x-0.5*w,summ["CM_sel"],w,label=f"CM {selected_year}",color="#D8C37D")
                    b3=ax.bar(x+0.5*w,summ["YTD_prev"],w,label=f"YTD {prev_year}",color="#006394")
                    b4=ax.bar(x+1.5*w,summ["YTD_sel"],w,label=f"YTD {selected_year}",color="#C1A02E")
                    ymax = summ.to_numpy().max() if summ.to_numpy().max() > 0 else 1; pad = ymax*0.015
                    def fmt(v): return f"SAR {int(v/1000)}K" if v>=1000 else f"SAR {int(v)}"
                    for bars_g in [b1,b2,b3,b4]:
                        for bar in bars_g:
                            h = bar.get_height()
                            if h>0: ax.text(bar.get_x()+bar.get_width()/2,h+pad,fmt(h),ha="center",va="bottom",fontsize=9,rotation=90,color="#333333",clip_on=False)
                    ax.set_xticks(x); ax.set_xticklabels([fill(d,18) for d in summ.index],rotation=30,ha="right",fontsize=9)
                    ax.set_ylabel("Cost Amount (SAR)")
                    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.grid(False)
                    ax.legend(loc="upper center",bbox_to_anchor=(0.5,-0.22),ncol=4,frameon=False,fontsize=10)
                    fig.subplots_adjust(bottom=0.35); return fig

                # ── NCR Root Cause (slide 19) ──
                def slide_ncr_rootcause_fig(selected_year, selected_month, top_n_reasons=4):
                    base = df_ncr_dash[(df_ncr_dash["Is_Valid"]==True)].copy()
                    rc_col = next((c for c in base.columns if "root" in c.lower() and "cause" in c.lower()), None)
                    if rc_col is None: rc_col = next((c for c in base.columns if "root" in c.lower()), None)
                    if rc_col is None: rc_col = next((c for c in base.columns if "cause" in c.lower()), None)
                    if rc_col is None: return None, f"No root cause column found in NCR file."
                    base["_Reason"] = base["Reason"].astype(str).str.strip()
                    base["_RC"]     = base[rc_col].astype(str).str.strip().replace("nan","Not Specified")
                    cm = base[(base["Year"]==selected_year)&(base["Month"]==selected_month)]
                    if cm.empty: return None, "No NCR data for selected period."
                    top_reasons = cm["_Reason"].value_counts().head(top_n_reasons).index.tolist()
                    n = len(top_reasons)
                    if n==0: return None, "No reasons found."
                    row_heights = [max(1.8, cm[cm["_Reason"]==r]["_RC"].nunique()*0.45+0.8) for r in top_reasons]
                    fig_h = sum(row_heights)+0.5
                    fig,axes = plt.subplots(n,1,figsize=(13.33,fig_h),dpi=300,gridspec_kw={"height_ratios":row_heights})
                    if n==1: axes=[axes]
                    palette=["#006394","#C1A02E","#0F68B9","#D8C37D","#B7910E","#4A90D9","#E8A838","#2E6DA4","#8EC6E6","#F5D07A","#1A4F72","#E09020","#5BA3C9","#C8A830","#3D7EA6"]
                    for ax,reason in zip(axes,top_reasons):
                        subset  = cm[cm["_Reason"]==reason]
                        rc_all  = subset["_RC"].value_counts()
                        reason_n= len(subset)
                        y_pos   = np.arange(len(rc_all))
                        bars    = ax.barh(y_pos,rc_all.values,color=[palette[i%len(palette)] for i in range(len(rc_all))],height=0.55,edgecolor="white")
                        for bar,val in zip(bars,rc_all.values):
                            pct = val/reason_n*100
                            ax.text(bar.get_width()+rc_all.max()*0.02,bar.get_y()+bar.get_height()/2,f"{int(val)}  ({pct:.0f}%)",va="center",ha="left",fontsize=9,color="#000000")
                        ax.set_yticks(y_pos); ax.set_yticklabels([fill(str(r),50) for r in rc_all.index],fontsize=9)
                        ax.invert_yaxis(); ax.set_xlim(0,rc_all.max()*1.35)
                        ax.set_title(f"  {fill(reason,80)}   [n={reason_n}]",loc="left",fontsize=10,fontweight="bold",color="#333333",pad=5)
                        ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.spines["left"].set_visible(False)
                        ax.tick_params(axis="x",labelsize=8); ax.set_xlabel("Count",fontsize=8); ax.grid(False)
                        ax.axhline(-0.6,color="#EEEEEE",linewidth=0.8)
                    plt.tight_layout(h_pad=1.5)
                    return fig, None

                # ── Intro slide helper ──
                def make_intro_slide(title_text, subtitle_text=""):
                    fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)
                    fig.patch.set_facecolor("#006394")
                    ax.set_facecolor("#006394")
                    ax.text(0.5, 0.55, title_text, transform=ax.transAxes,
                            ha="center", va="center", fontsize=36, fontweight="bold",
                            color="white", wrap=True)
                    if subtitle_text:
                        ax.text(0.5, 0.35, subtitle_text, transform=ax.transAxes,
                                ha="center", va="center", fontsize=22, color="#D8C37D")
                    ax.axis("off")
                    plt.tight_layout()
                    return fig

                month_name = pd.to_datetime(f"{selected_year}-{selected_month:02d}-01").strftime("%B %Y")

                # ── Build ordered PDF ──
                pdf_slides = []

                # Cover slide
                _month_full = pd.to_datetime(f"{selected_year}-{selected_month:02d}-01").strftime("%B %Y")
                _today_str  = pd.Timestamp.now().strftime("%d \u2013 %B - %Y")
                pdf_slides.append({
                    "cover":    True,
                    "title":    "Quality Indicators Report",
                    "subtitle": f"Easternpak \u2013 {_month_full}",
                    "date":     _today_str,
                })

                # 1 - Total Complaints Issued (ISSUED donut)
                pdf_slides.append({"title": "Total Complaints Issued", "fig": slide_issued_1_donuts_fig(selected_year, selected_month)})

                # 2 - Total Complaints Approved (FINAL donut)
                pdf_slides.append({"title": "Total Complaints Approved", "fig": slide_1_final_donuts_fig(selected_year, selected_month)})

                # 3 - Complaints Lead Time
                pdf_slides.append({"title": "Complaints Lead Time — First Approval", "fig": slide_2_leadtime_fig(selected_year, selected_month)})

                # 4 - Complaints Ratio (Total)
                if months_with_fg:
                    pdf_slides.append({"title": "Complaints Ratio", "fig": slide_cc_ratio_fig(selected_year, category_filter=None)})
                    # 5 - Quality Complaints Ratio
                    pdf_slides.append({"title": "Quality Complaints Ratio", "fig": slide_cc_ratio_fig(selected_year, category_filter="Quality")})
                    # 6 - Service Complaints Ratio
                    pdf_slides.append({"title": "Service Complaints Ratio", "fig": slide_service_cc_ratio_fig(selected_year)})

                # 7 - Total Complaints Issued YoY
                pdf_slides.append({"title": "Total Complaints Issued — YoY", "fig": slide_issued_valid_count_fig(selected_year, selected_month)})

                # Intro slide — Month overview
                pdf_slides.append({"title": month_name, "section": True})

                # 8 - Breakdown Quality Complaints Issued CM
                pdf_slides.append({"title": f"Breakdown of Quality Complaints Issued — {month_name}", "fig": slide_issued_quality_reason_current_month_fig(selected_year, selected_month, TOPN_QUALITY_CM)})

                # 9 - Root Cause Quality Complaints Issued CM
                _fig_rcq, _err_rcq = slide_rootcause_fig("Quality", selected_year, selected_month, TOPN_RC_REASONS or 4)
                if _fig_rcq:
                    pdf_slides.append({"title": f"Root Cause of Quality Complaints Issued — {month_name}", "fig": _fig_rcq})

                # 10 - Breakdown Service Complaints Issued CM
                pdf_slides.append({"title": f"Breakdown of Service Complaints Issued — {month_name}", "fig": slide_issued_service_reason_current_month_fig(selected_year, selected_month, TOPN_SERVICE_CM)})

                # 11 - Root Cause Service Complaints Issued CM
                _fig_rcs, _err_rcs = slide_rootcause_fig("Service", selected_year, selected_month, TOPN_RC_REASONS or 4)
                if _fig_rcs:
                    pdf_slides.append({"title": f"Root Cause of Service Complaints Issued — {month_name}", "fig": _fig_rcs})

                # Intro slide — Quality Complaints Overview
                pdf_slides.append({"title": "Quality Complaints Overview", "section": True})

                # 12 - Valid Quality Complaints Count
                pdf_slides.append({"title": "Valid Quality Complaints Count — To Date", "fig": slide_3_valid_quality_count_fig(selected_year, selected_month)})

                # 13 - Quality Defects CM vs YTD
                pdf_slides.append({"title": "Valid Quality Complaints by Defect — CM vs To Date", "fig": slide_4_quality_defect_cm_vs_ytd_fig(selected_year, selected_month, TOPN_QUALITY_DEFECT)})

                # 14 - Quality Defects CM
                pdf_slides.append({"title": f"Valid Quality Complaints by Reason — {month_name}", "fig": slide_5_quality_defect_current_month_fig(selected_year, selected_month, TOPN_QUALITY_CM)})

                # 15 - Quality CN Cost
                try:
                    pdf_slides.append({"title": "Quality Complaints Value (CN at Cost) by Defect — To Date", "fig": slide_6_quality_cost_cm_vs_ytd_fig(selected_year, selected_month, TOPN_COST_DEFECT)})
                except Exception: pass

                # 18 - Internal Defect Ratio Trend (= NCR Ratio)
                if wo_ready:
                    pdf_slides.append({"title": "Internal Defect Ratio Trend", "fig": slide_ncr_ratio_fig(selected_year)})

                # 19 - NCR Root Cause
                _fig_ncr_rc, _err_ncr_rc = slide_ncr_rootcause_fig(selected_year, selected_month, TOPN_RC_REASONS or 4)
                if _fig_ncr_rc:
                    pdf_slides.append({"title": "NCR Breakdown by Root Cause", "fig": _fig_ncr_rc})

                # 20 - NCR and CRM Correlation
                pdf_slides.append({"title": "NCR and CRM Correlation — YTD", "fig": slide_ncr_vs_crm_correlation_ytd_fig(selected_year, selected_month, TOPN_CORRELATION)})

                # Intro slide — Service Complaints Overview
                pdf_slides.append({"title": "Service Complaints Overview", "section": True})

                # 21 - Valid Service Count
                pdf_slides.append({"title": "Valid Service Complaints Count", "fig": slide_s1_valid_service_count_fig(selected_year, selected_month)})

                # 22 - Service Reasons CM vs YTD
                pdf_slides.append({"title": "Valid Service Complaints — CM vs YTD", "fig": slide_s2_service_reason_cm_vs_ytd_fig(selected_year, selected_month, TOPN_SERVICE_REASON)})

                # 23 - Service CN Cost
                try:
                    pdf_slides.append({"title": "Service Complaints Value (CN at Cost) by Defect — To Date", "fig": slide_service_cn_cost_fig(selected_year, selected_month, TOPN_COST_DEFECT)})
                except Exception: pass

                # Intro slide — Cost of Quality
                pdf_slides.append({"title": "Cost of Quality", "section": True})

                # 24 - COQ CM (single month bar — reuse breakdown fig)
                if coq_ready:
                    pdf_slides.append({"title": f"Cost of Quality — {month_name}", "fig": slide_coq_breakdown_fig(selected_year, selected_month)})


                # PDF Export
                st.divider()
                st.header("Export")
                pdf_buf = build_ppt_pdf(pdf_slides, dpi=300)
                st.download_button(
                    label="📥 Download PPT-style PDF",
                    data=pdf_buf,
                    file_name=f"QM_Dashboard_{selected_year}-{selected_month:02d}.pdf",
                    mime="application/pdf",
                )

            # ── Reason Trend Explorer (outside run block — persists on rerun) ──
            if "qm_df_final" in st.session_state and "qm_selected_year" in st.session_state:
                _df_trend     = st.session_state["qm_df_final"]
                _year_trend   = st.session_state["qm_selected_year"]
                _month_trend  = st.session_state["qm_selected_month"]
                _prev_year_trend = _year_trend - 1

                st.divider()
                st.subheader("📈 Reason Trend Explorer")
                st.caption("Select a reason to see its monthly trend — FINAL file, Final Approval Date.")

                _all_reasons_t = sorted([
                    r for r in _df_trend[(_df_trend["Is_Valid"]==True)]["Reason"].astype(str).str.strip().unique()
                    if r and r.lower() not in ("nan","")
                ])

                _selected_reason = st.selectbox(
                    "Select Reason", ["— select a reason —"] + _all_reasons_t,
                    key="qm_trend_reason"
                )

                if _selected_reason != "— select a reason —":
                    _MONTH_LABELS_T = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
                    _months_t = list(range(1, 13))

                    def _get_monthly_trend(yr):
                        base = _df_trend[
                            (_df_trend["Is_Valid"]==True) &
                            (_df_trend["Reason"].astype(str).str.strip()==_selected_reason) &
                            (_df_trend["Year"]==yr)
                        ]
                        return base.groupby("Month").size().reindex(_months_t, fill_value=0)

                    _vals_prev = _get_monthly_trend(_prev_year_trend)
                    _vals_sel  = _get_monthly_trend(_year_trend)
                    _ytd_prev  = int(_vals_prev.iloc[:_month_trend].sum())
                    _ytd_sel   = int(_vals_sel.iloc[:_month_trend].sum())

                    _x = np.arange(12); _w = 0.35
                    _fig_t, _ax_t = plt.subplots(figsize=(13.33, 7.5), dpi=300)

                    _b1 = _ax_t.bar(_x - _w/2, _vals_prev.values, _w, color="#006394",
                                    label=f"{_prev_year_trend}  (YTD: {_ytd_prev})", alpha=0.85)
                    _b2 = _ax_t.bar(_x + _w/2, _vals_sel.values,  _w, color="#C1A02E",
                                    label=f"{_year_trend}  (YTD: {_ytd_sel})", alpha=0.85)

                    for _b in [_b1, _b2]:
                        for _bar in _b:
                            _h = _bar.get_height()
                            if _h > 0:
                                _ax_t.text(_bar.get_x()+_bar.get_width()/2, _h+0.3,
                                           f"{int(_h)}", ha="center", va="bottom",
                                           fontsize=10, color="#4D4D4D", clip_on=False)

                    _y_trend = _vals_sel.values[:_month_trend].astype(float)
                    if len(_y_trend) >= 2 and np.any(_y_trend > 0):
                        _xf = np.arange(_month_trend, dtype=float)
                        _coeff = np.polyfit(_xf, _y_trend, 1)
                        _ax_t.plot(_x[:_month_trend], np.polyval(_coeff, _xf),
                                   linestyle="--", linewidth=1.5, color="#C1A02E",
                                   label=f"Trend {_year_trend}")

                    _ax_t.axvline(_month_trend - 1, color="#DE201B", linewidth=1.2,
                                  linestyle=":", alpha=0.6, label="Current Month")
                    _ax_t.set_xticks(_x)
                    _ax_t.set_xticklabels(_MONTH_LABELS_T, fontsize=11)
                    _ax_t.set_ylabel("Count")
                    _ax_t.set_ylim(0, max(max(_vals_prev.max(), _vals_sel.max()), 1) * 1.25)
                    _ax_t.spines["top"].set_visible(False)
                    _ax_t.spines["right"].set_visible(False)
                    _ax_t.grid(False)
                    _ax_t.legend(loc="upper center", bbox_to_anchor=(0.5, -0.08),
                                 ncol=4, frameon=False, fontsize=11)
                    plt.tight_layout(rect=[0, 0.05, 1, 1])
                    st.image(fig_to_png_bytes(_fig_t))
                    plt.close(_fig_t)

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
            crm_df = pd.DataFrame(crm_list)[["crm_ref","customer_name","delete_count","added_by","added_at"]].rename(columns={
                "crm_ref": "CRM #", "customer_name": "Customer",
                "delete_count": "Rows Deleted", "added_by": "Added By", "added_at": "Added At"
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
            new_delete_count = c2.number_input("Rows to DELETE", min_value=0, value=0, step=1)

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
                            "delete_count": int(new_delete_count),
                            "added_by": name,
                        }, on_conflict="crm_ref").execute()
                        st.success(f"✅ {new_crm_ref} saved — keeping {new_delete_count} rows.")
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


# ════════════════════════════════════════
# TAB 6 — QUALITY MATRIX (PPM TABLE)
# ════════════════════════════════════════
with tab6:
    st.markdown("### 🔬 Quality Matrix — PPM Analysis")

    # ── File uploaders ──
    _c1, _c2 = st.columns(2)
    _qm_final_file   = _c1.file_uploader("📂 FINAL Approved file",      type=["xls","xlsx"], key="qm6_final")
    _qm_ncr_app_file = _c2.file_uploader("📂 NCR Fully Approved file",  type=["xls","xlsx"], key="qm6_ncr_app")

    if not _qm_final_file or not _qm_ncr_app_file:
        st.info("⬆️ Upload both files above to enable PPM analysis.")
    else:
        try:
            with st.spinner("Loading files..."):
                _qm6_crm_map, _qm6_q_set, _qm6_s_set, _qm6_i_set = load_settings_from_supabase(supabase)
                _qm6_classifier = make_classifier(_qm6_q_set, _qm6_s_set, _qm6_i_set)

                _qm6_final_loaded = read_excel_from_upload(_qm_final_file)
                _qm6_final_pkg    = build_dataset_final_issued(
                    _qm6_final_loaded, date_col=FINAL_APPROVAL_COL,
                    dataset_name="QM6_FINAL", crm_delete_map={}, classifier=_qm6_classifier
                )
                _qm6_df_final = _qm6_final_pkg["cleaned_flagged"].copy()
                for _col in ["Year","Month"]:
                    _qm6_df_final[_col] = pd.to_numeric(_qm6_df_final[_col], errors="coerce").astype("Int64")
                _qm6_df_final = _qm6_df_final.dropna(subset=["Year","Month"])
                _qm6_df_final["Year"]  = _qm6_df_final["Year"].astype(int)
                _qm6_df_final["Month"] = _qm6_df_final["Month"].astype(int)

                _qm6_ncr_loaded = read_excel_from_upload(_qm_ncr_app_file)
                _qm6_ncr_loaded["_Base_Date"] = pd.to_datetime(_qm6_ncr_loaded[FINAL_APPROVAL_COL], errors="coerce")
                _qm6_ncr_loaded["Year"]  = _qm6_ncr_loaded["_Base_Date"].dt.year.astype("Int64")
                _qm6_ncr_loaded["Month"] = _qm6_ncr_loaded["_Base_Date"].dt.month.astype("Int64")
                _qm6_ncr_loaded = _qm6_ncr_loaded.dropna(subset=["Year","Month"])
                _qm6_ncr_loaded["Year"]  = _qm6_ncr_loaded["Year"].astype(int)
                _qm6_ncr_loaded["Month"] = _qm6_ncr_loaded["Month"].astype(int)

            st.success("✅ Files loaded!")

            # ── Unknown reasons prompt ──
            _qm6_unc = set()
            for r in _qm6_final_pkg["unclassified_counts"].index:
                _qm6_unc.add(str(r).strip())
            _qm6_unc = {r for r in _qm6_unc if r and r.lower() not in ("","nan","invalid")}
            if _qm6_unc:
                st.warning(f"⚠️ {len(_qm6_unc)} unclassified reason(s) found. Please classify:")
                with st.form("qm6_classify_form"):
                    _qm6_cls = {}
                    for r in sorted(_qm6_unc):
                        _cc1, _cc2 = st.columns([3,1])
                        _cc1.markdown(f"**{r}**")
                        _qm6_cls[r] = _cc2.radio("", ["Quality","Service","Invalid (exclude)"],
                                                   key=f"qm6_cls_{r[:40]}", horizontal=True)
                    if st.form_submit_button("💾 Save Classifications", type="primary"):
                        for r, cls in _qm6_cls.items():
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
                        st.success("✅ Saved! Re-upload files to apply.")
                        st.rerun()

            # ── Period selector ──
            _qm6_dates = _qm6_df_final[["Year","Month"]].drop_duplicates().sort_values(["Year","Month"])
            _qm6_years = sorted(_qm6_dates["Year"].unique().tolist())
            _pc1, _pc2, _pc3 = st.columns(3)
            _qm6_year       = _pc1.selectbox("Year", _qm6_years, index=len(_qm6_years)-1, key="qm6_year")
            _qm6_months_avail = sorted(_qm6_dates[_qm6_dates["Year"]==_qm6_year]["Month"].unique().tolist())
            _MONTH_LABELS_6   = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
            _month_opts       = [_MONTH_LABELS_6[m-1] for m in _qm6_months_avail]
            _qm6_from_lbl = _pc2.selectbox("From Month", _month_opts, index=0, key="qm6_from_month")
            _qm6_to_lbl   = _pc3.selectbox("To Month",   _month_opts, index=len(_month_opts)-1, key="qm6_to_month")
            _qm6_from_month = _MONTH_LABELS_6.index(_qm6_from_lbl) + 1
            _qm6_to_month   = _MONTH_LABELS_6.index(_qm6_to_lbl)   + 1

            if _qm6_from_month > _qm6_to_month:
                st.warning("⚠️ 'From Month' must be before or equal to 'To Month'.")
                st.stop()

            # ── Load production data from Supabase ──
            try:
                _prod_rows = supabase.table("qm_production_data").select("*").execute()
                _prod_data = {(r["year"], r["month"]): r["produced_qty"] for r in (_prod_rows.data or [])}
            except Exception:
                _prod_data = {}

            # Check missing production months in selected range
            _needed_months   = list(range(_qm6_from_month, _qm6_to_month + 1))
            _missing_prod    = [(int(_qm6_year), m) for m in _needed_months
                                if _prod_data.get((int(_qm6_year), m)) is None]

            if _missing_prod:
                st.warning(f"⚠️ Production data missing for {len(_missing_prod)} month(s):")
                with st.form("qm6_prod_form"):
                    _prod_inputs = {}
                    _pcols = st.columns(min(6, len(_missing_prod)))
                    for i, (y, m) in enumerate(_missing_prod):
                        _prod_inputs[(y,m)] = _pcols[i % len(_pcols)].number_input(
                            f"{_MONTH_LABELS_6[m-1]}-{str(y)[-2:]}",
                            min_value=0.0, value=0.0, step=100000.0,
                            key=f"qm6_prod_{y}_{m}"
                        )
                    if st.form_submit_button("💾 Save Production Data", type="primary"):
                        _perrors = []
                        for (y, m), val in _prod_inputs.items():
                            try:
                                supabase.table("qm_production_data").upsert({
                                    "year": int(y), "month": int(m),
                                    "produced_qty": float(val), "updated_by": name
                                }, on_conflict="year,month").execute()
                            except Exception as e:
                                _perrors.append(f"{_MONTH_LABELS_6[m-1]}: {e}")
                        if _perrors:
                            for err in _perrors: st.error(err)
                        else:
                            st.success("✅ Production data saved!"); st.rerun()

            # Total produced qty for selected range
            _total_produced = sum(
                float(_prod_data.get((int(_qm6_year), m), 0) or 0)
                for m in _needed_months
            )

            _qm6_topn_col, _qm6_btn_col = st.columns([1, 3])
            _qm6_topn = _qm6_topn_col.number_input("Top N Defects (chart)", min_value=3, max_value=50, value=15, step=1, key="qm6_topn")

            if _qm6_btn_col.button("📊 Generate PPM Table", type="primary", key="qm6_run"):
                if _total_produced == 0:
                    st.warning("⚠️ Production data missing or zero for selected period. Please fill in above.")
                else:
                    with st.spinner("Computing PPM..."):
                        _dec_col = next((c for c in _qm6_df_final.columns if "decision" in c.lower()), None)
                        if not _dec_col:
                            st.error("Could not find Decision column in FINAL file."); st.stop()

                        _qm6_final_filt = _qm6_df_final[
                            (_qm6_df_final["Is_Valid"] == True) &
                            (_qm6_df_final["Complaint_Category"] == "Quality") &
                            (_qm6_df_final["Year"] == int(_qm6_year)) &
                            (_qm6_df_final["Month"] >= _qm6_from_month) &
                            (_qm6_df_final["Month"] <= _qm6_to_month) &
                            (_qm6_df_final[_dec_col].astype(str).str.strip().str.lower() == "credit note")
                        ].copy()
                        _qm6_final_filt["_Reason"] = _qm6_final_filt["Reason"].astype(str).str.strip()
                        _qm6_final_filt["_Qty"]    = pd.to_numeric(_qm6_final_filt["Dec Qty"], errors="coerce").fillna(0)
                        _qm6_final_qty = _qm6_final_filt.groupby("_Reason")["_Qty"].sum()

                        _qm6_dec_col_ncr = next((c for c in _qm6_ncr_loaded.columns if "decision" in c.lower()), None)
                        if not _qm6_dec_col_ncr:
                            st.error("Could not find Decision column in NCR file."); st.stop()

                        _qm6_ncr_filt = _qm6_ncr_loaded[
                            (_qm6_ncr_loaded["Year"] == int(_qm6_year)) &
                            (_qm6_ncr_loaded["Month"] >= _qm6_from_month) &
                            (_qm6_ncr_loaded["Month"] <= _qm6_to_month) &
                            (_qm6_ncr_loaded[_qm6_dec_col_ncr].astype(str).str.strip().str.lower() == "shredding")
                        ].copy()
                        _qm6_ncr_filt["_Reason"] = _qm6_ncr_filt["Reason"].astype(str).str.strip()
                        _qm6_ncr_filt["_Qty"]    = pd.to_numeric(_qm6_ncr_filt["Dec Qty"], errors="coerce").fillna(0)
                        _qm6_ncr_qty = _qm6_ncr_filt.groupby("_Reason")["_Qty"].sum()

                        _qm6_all_reasons = set(_qm6_final_qty.index) | set(_qm6_ncr_qty.index)
                        _qm6_rows = []
                        for r in sorted(_qm6_all_reasons):
                            _f_qty = float(_qm6_final_qty.get(r, 0))
                            _n_qty = float(_qm6_ncr_qty.get(r, 0))
                            _total_r = _f_qty + _n_qty
                            _ppm_r   = round((_total_r / _total_produced) * 1_000_000, 1) if _total_produced > 0 else 0
                            _qm6_rows.append({
                                "Defect":                 r,
                                "NCR Shredding (Qty)":    int(_n_qty),
                                "Credit Note (Qty)":      int(_f_qty),
                                "Total Defective Qty":    int(_total_r),
                                "PPM":                    _ppm_r,
                            })

                        _qm6_ppm_df = pd.DataFrame(_qm6_rows)
                        _qm6_ppm_df = _qm6_ppm_df.sort_values("PPM", ascending=False).reset_index(drop=True)
                        _qm6_total_def = int(_qm6_ppm_df["Total Defective Qty"].sum())
                        _qm6_total_ppm = round((_qm6_total_def / _total_produced) * 1_000_000, 1) if _total_produced > 0 else 0

                        st.session_state["qm6_ppm_df"]       = _qm6_ppm_df
                        st.session_state["qm6_total_def"]    = _qm6_total_def
                        st.session_state["qm6_total_ppm"]    = _qm6_total_ppm
                        st.session_state["qm6_total_prod"]   = _total_produced
                        st.session_state["qm6_ppm_year"]     = _qm6_year
                        st.session_state["qm6_ppm_topn"]     = int(_qm6_topn)
                        st.session_state["qm6_ppm_from"]     = _qm6_from_month
                        st.session_state["qm6_ppm_to"]       = _qm6_to_month

            # ── Display ──
            if "qm6_ppm_df" in st.session_state:
                _ppm_df      = st.session_state["qm6_ppm_df"]
                _total_def   = st.session_state["qm6_total_def"]
                _total_ppm   = st.session_state["qm6_total_ppm"]
                _total_prod  = st.session_state["qm6_total_prod"]
                _yr          = st.session_state["qm6_ppm_year"]
                _fm          = st.session_state["qm6_ppm_from"]
                _tm          = st.session_state["qm6_ppm_to"]
                _MONTH_L     = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
                _period_lbl  = f"{_MONTH_L[_fm-1]} to {_MONTH_L[_tm-1]} {_yr}"

                st.divider()
                st.markdown(f"#### PPM Table — {_period_lbl}")

                _mc1, _mc2, _mc3, _mc4 = st.columns(4)
                _mc1.metric("Sheets Produced",      f"{int(_total_prod):,}")
                _mc2.metric("NCR Shredding Qty",     f"{int(_ppm_df['NCR Shredding (Qty)'].sum()):,}")
                _mc3.metric("Credit Note Qty",       f"{int(_ppm_df['Credit Note (Qty)'].sum()):,}")
                _mc4.metric("Total PPM",             f"{_total_ppm:,.1f}")

                st.dataframe(
                    _ppm_df.style.apply(
                        lambda row: ["background-color:#EAF4FB" if row["Total Defective Qty"]>0 else "" for _ in row],
                        axis=1
                    ),
                    use_container_width=True, hide_index=True
                )

                # ── Pareto ──
                _qm6_topn_disp = st.session_state.get("qm6_ppm_topn", 15)
                _pareto = _ppm_df[_ppm_df["PPM"]>0].head(int(_qm6_topn_disp)).copy()
                if not _pareto.empty:
                    st.divider()
                    st.markdown(f"#### Pareto — PPM by Defect ({_period_lbl})")
                    _pareto["Cumulative %"] = (_pareto["PPM"].cumsum() / _pareto["PPM"].sum() * 100)
                    _fig_p, _ax_p = plt.subplots(figsize=(13.33, 7.5), dpi=150)
                    _xp = np.arange(len(_pareto))
                    _bars_p = _ax_p.bar(_xp, _pareto["PPM"].values, color="#006394", width=0.6, alpha=0.9)
                    for _bar, _val in zip(_bars_p, _pareto["PPM"].values):
                        _ax_p.text(_bar.get_x()+_bar.get_width()/2,
                                   _bar.get_height()+_pareto["PPM"].max()*0.01,
                                   f"{_val:.1f}", ha="center", va="bottom", fontsize=9, color="#000000")
                    _ax_p2 = _ax_p.twinx()
                    _ax_p2.plot(_xp, _pareto["Cumulative %"].values, color="#C1A02E",
                                linewidth=2.5, marker="o", markersize=5, label="Cumulative %")
                    _ax_p2.axhline(80, color="#DE201B", linewidth=1, linestyle="--", alpha=0.6)
                    _ax_p2.set_ylabel("Cumulative %"); _ax_p2.set_ylim(0, 115)
                    _ax_p2.tick_params(axis="y"); _ax_p2.spines["top"].set_visible(False)
                    _ax_p.set_xticks(_xp)
                    _ax_p.set_xticklabels([fill(d,16) for d in _pareto["Defect"]],
                                          rotation=40, ha="right", fontsize=8)
                    _ax_p.set_ylabel("PPM")
                    _ax_p.spines["top"].set_visible(False); _ax_p.spines["right"].set_visible(False)
                    _ax_p.grid(False)
                    _l2, _lb2 = _ax_p2.get_legend_handles_labels()
                    _ax_p.legend(_l2, _lb2, loc="upper right", frameon=False, fontsize=10)
                    plt.tight_layout()
                    st.image(fig_to_png_bytes(_fig_p))
                    plt.close(_fig_p)

                # Download
                _buf6 = io.BytesIO()
                _ppm_df.to_excel(_buf6, index=False, engine="openpyxl")
                _buf6.seek(0)
                st.download_button(
                    label="📥 Download PPM Table",
                    data=_buf6,
                    file_name=f"PPM_{_yr}_{_MONTH_L[_fm-1]}_to_{_MONTH_L[_tm-1]}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_ppm_table"
                )

        except Exception as e:
            st.error(f"Error: {e}")
