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
    COLORS = {"Service": "#C1A02E", "Quality": "#006394", "Invalid": "#D8C37D", "Commercial": "#2E8449"}

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
        _logo_b64 = "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAIBAQEBAQIBAQECAgICAgQDAgICAgUEBAMEBgUGBgYFBgYGBwkIBgcJBwYGCAsICQoKCgoKBggLDAsKDAkKCgr/2wBDAQICAgICAgUDAwUKBwYHCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgr/wAARCAF/BVYDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9+KKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigBB7jFLnsazvEXiPSvCuhX3ijXr+O1sdNtJLm+upThYoY1LO59goJ/CvxZ/bG/4OE/2nPGnxFu9I/ZI1Oz8G+FrG4aOy1GfSbe7vtQUHHmuLlJI41PUIq7gOrZ4HsZPkWYZ5VlDDJWju27JX26M+fz/iXLOHaMZ4pu8top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iRo3yPStf7Wy/CK+Bw6Uv5pvma9FZJW6PUy/sTNMc/8AhRxLcf5aacE/V3bd+qTPkP8AZZ/4IrfsW/s1SQeItU8GHxx4ijbe2seLUW4VZP7yQY8tTnnJDEHnNfXVtClvCsUcQRVGFVegHoKkGDz/AFpa8bF43F46pz15uT83+Xb5Hu4LLsFl1Pkw8FFeS/Pq35sKKKK5jtCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigVkgooooGFFFFABgelFFFArIKKKKBhRRRQAYHpRRRQCSQUUUUBZBXwv/wcOf8AKPK4/wCxz0v+ctfdFfC//Bw5/wAo8rj/ALHPS/5y17HD3/I8w/8Ajj+Z8/xR/wAk5iv+vcvyZ+Dh+6K/Zj/g2X/5N/8AiN/2N0H/AKTivxnP3RX7Mf8ABsv/AMm//Eb/ALG6D/0nFfr/AB3/AMk5P1j/AOlI/CPDL/kqYf4ZfkfpnRRRX4Qf0vZBRRRQAUUUUAFFFFABRRRQAYHpRRRQKyDA9KKKKB2QUUUUAFFFFAWQUUUUAFFFFABRRRQAUUUUAFFFFABgelFFFArIKKKKBhRRRQAYA6CjA9KKKBWQUUUUDCiiigAooooCyCiiigLIKKKKACiiigApCiHqg/KlooE0nuFBAPUUUUDshNq+gpaKKBWSCiiigYUUUUAGAOgooooCyQUUUUAFFFFABgelFFFAWSCiiigAooooAKKKKAsgooooAKKKKADA9KKKKAskFGB1xRRQKyCiiigYUUUUAIVU9QD9aAqrwoA+lLRQKyuFFFFAwooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAQ98da+F/8Ag4ebH/BPG4GP+Zz0z/2rX3RjJ5HTpXkv7Zv7IHw8/be+DMnwQ+Jus6nY6ZJqdvfNPpMiLL5kW7aMurDHzHPFehlOJp4PM6Ver8MZJv0TPKzvCVcwyivh6XxTi0r6K7VkfzHjBUtiv2X/AODZdgv7PvxGz/0N0H/pNXRj/g20/Y4xg/Ezx1j/AK/rf/41X01+wr+wB8Kv2BPB2teCvhR4h1rULbW9SW9upNYljZlkVNgC7FXjFfoHFHFmU5tlMsNh3LmbT1Vtnc/L+DOCc6yPO44rEqPIk1o7vVaaWPd6KKK/Lz9lCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA//9k="
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
        # Logo — pushed higher
        if _napco_logo_reader:
            try:
                c.drawImage(_napco_logo_reader, 18, H - 115, width=280, height=110,
                            preserveAspectRatio=True, mask="auto")
            except Exception:
                pass
        # Blue horizontal bar — from LEFT edge of logo to right edge of slide
        # Logo starts at x=18
        line_y = H - 125
        logo_left = 18
        c.setFillColor(HexColor("#0C5595"))
        c.rect(logo_left, line_y, W - logo_left, 4, fill=1, stroke=0)
        # Thin vertical red bar — flush with left edge, top to blue line
        c.setFillColor(HexColor("#DE201B"))
        c.rect(0, line_y, 4, H - line_y, fill=1, stroke=0)
        # Layout: title block centered in remaining space (below lines to above date)
        # line is at H-125, remaining space down to 50
        remaining_center = (H - 125 + 50) / 2  # midpoint between blue line and bottom
        title_font_size = 44
        subtitle_font_size = 20
        gap = 28  # gap between title bottom and red line, and red line and subtitle top
        # Title y — above center
        title_y = remaining_center + gap + 10
        c.setFillColor(HexColor("#0E5E86"))
        c.setFont("Helvetica-Bold", title_font_size)
        title_w = c.stringWidth(title_text, "Helvetica-Bold", title_font_size)
        c.drawString((W - title_w) / 2, title_y, title_text)
        # Red line — centered between title baseline and subtitle top
        line_y = title_y - gap
        c.setFillColor(HexColor("#DE201B"))
        c.rect(80, line_y, W - 160, 2, fill=1, stroke=0)
        # Subtitle — same gap below red line
        sub_y = line_y - gap - 4
        c.setFillColor(HexColor("#555555"))
        c.setFont("Helvetica-Oblique", subtitle_font_size)
        sub_w = c.stringWidth(subtitle_text, "Helvetica-Oblique", subtitle_font_size)
        c.drawString((W - sub_w) / 2, sub_y, subtitle_text)
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

            run = st.button("📄 Generate & Download PDF", type="primary", key="qm_run")

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
                    labels = ["Service","Quality","Invalid","Commercial"]
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
                                f"{int(v)}",ha="center",va="center",fontsize=11,color="white")
                    ax.axis("equal")
                    ax.legend(handles=[plt.Rectangle((0,0),1,1,fc=COLORS[l]) for l in labels if values[labels.index(l)]>0], labels=[l for l in labels if values[labels.index(l)]>0], loc="upper center", bbox_to_anchor=(0.5, -0.05), ncol=len(labels),frameon=False, fontsize=10)

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
                    plt.tight_layout(rect=[0,0.14,1,1]); return fig

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
                    plt.tight_layout(rect=[0,0.14,1,1]); return fig

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
                    pass


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

                # ── Service Root Cause Drill-Down ──
                st.subheader("Service — Root Cause by Reason (Current Month)")
                fig_rc_s, err_rc_s = slide_rootcause_fig("Service", selected_year, selected_month, TOPN_RC_REASONS or 4)
                if err_rc_s:
                    st.warning(f"Service root cause: {err_rc_s}")
                else:
                    pass


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
                    pass
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

                    st.subheader("Quality Customer Complaints Ratio")

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
                pdf_slides.append({"title": f"{month_name} Overview", "section": True})

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

# ── Commercial Reasons slides ──────────────────────────────────
                def slide_commercial_count_fig(year, month):
                    comm = df_final[
                        (df_final["Complaint_Category"] == "Commercial") &
                        (df_final["Year"] == year) &
                        (df_final["Month"].between(1, month))
                    ].copy()
                    if comm.empty:
                        fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)
                        ax.text(0.5, 0.5, "No Commercial CRMs found", ha="center", va="center", fontsize=20)
                        ax.axis("off"); return fig
                    comm["_Reason"] = comm["Reason"].astype(str).str.strip()
                    cm_counts  = comm[comm["Month"] == month]["_Reason"].value_counts()
                    ytd_counts = comm["_Reason"].value_counts()
                    all_r      = pd.Index(ytd_counts.index).union(cm_counts.index)
                    summ = pd.DataFrame({
                        "CM":  cm_counts.reindex(all_r, fill_value=0),
                        "YTD": ytd_counts.reindex(all_r, fill_value=0),
                    }).sort_values("YTD", ascending=False)
                    x = np.arange(len(summ)); w = 0.35
                    fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)
                    b1 = ax.bar(x - w/2, summ["CM"],  w, color="#006394", label="Current Month")
                    b2 = ax.bar(x + w/2, summ["YTD"], w, color="#2E8449", label="YTD")
                    for bars_g in [b1, b2]:
                        for bar in bars_g:
                            h = bar.get_height()
                            if h > 0:
                                ax.text(bar.get_x() + bar.get_width()/2, h + 0.3,
                                        f"{int(h)}", ha="center", va="bottom",
                                        fontsize=10, color="#4D4D4D")
                    ax.set_xticks(x)
                    ax.set_xticklabels(summ.index.tolist(), fontsize=12)
                    ax.set_ylabel("Count")
                    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
                    ax.grid(False)
                    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.08), ncol=2, frameon=False)
                    plt.tight_layout(rect=[0, 0.05, 1, 1])
                    return fig

                def slide_commercial_cost_fig(year, month):
                    dec_col = next((c for c in df_final.columns if "decision" in c.lower()), None)
                    if dec_col is None:
                        dec_col = next((c for c in df_final.columns if "dec" in c.lower() and "approv" not in c.lower()), None)
                    comm = df_final[
                        (df_final["Complaint_Category"] == "Commercial") &
                        (df_final["Year"] == year) &
                        (df_final["Month"].between(1, month))
                    ].copy()
                    if dec_col and dec_col in comm.columns:
                        comm = comm[comm[dec_col].astype(str).str.strip().str.lower() == "credit note"]
                    if comm.empty:
                        fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)
                        ax.text(0.5, 0.5, "No Commercial Credit Notes found", ha="center", va="center", fontsize=20)
                        ax.axis("off"); return fig
                    comm["_Reason"] = comm["Reason"].astype(str).str.strip()
                    comm["_Cost"]   = pd.to_numeric(comm["Cost Amount"], errors="coerce").fillna(0)
                    cm_cost  = comm[comm["Month"] == month].groupby("_Reason")["_Cost"].sum()
                    ytd_cost = comm.groupby("_Reason")["_Cost"].sum()
                    all_r    = pd.Index(ytd_cost.index).union(cm_cost.index)
                    summ = pd.DataFrame({
                        "CM":  cm_cost.reindex(all_r, fill_value=0),
                        "YTD": ytd_cost.reindex(all_r, fill_value=0),
                    }).sort_values("YTD", ascending=False)
                    def fmt_sar(v):
                        if v >= 1_000_000: return f"SAR {v/1_000_000:.1f}M"
                        if v >= 1000:      return f"SAR {int(v/1000)}K"
                        return f"SAR {int(v)}"
                    x = np.arange(len(summ)); w = 0.35
                    fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)
                    b1 = ax.bar(x - w/2, summ["CM"],  w, color="#006394", label="Current Month (Credit Notes)")
                    b2 = ax.bar(x + w/2, summ["YTD"], w, color="#2E8449", label="YTD (Credit Notes)")
                    ymax = summ.to_numpy().max() if summ.to_numpy().max() > 0 else 1
                    for bars_g in [b1, b2]:
                        for bar in bars_g:
                            h = bar.get_height()
                            if h > 0:
                                ax.text(bar.get_x() + bar.get_width()/2, h + ymax * 0.015,
                                        fmt_sar(h), ha="center", va="bottom",
                                        fontsize=9, rotation=90, color="#333333")
                    ax.set_xticks(x)
                    ax.set_xticklabels(summ.index.tolist(), fontsize=12)
                    ax.set_ylabel("Cost Amount (SAR)")
                    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
                    ax.grid(False)
                    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.08), ncol=2, frameon=False)
                    plt.tight_layout(rect=[0, 0.05, 1, 1])
                    return fig

                def slide_customer_return_cost_fig(year, month, top_n=10):
                    COST_PER_CLAIM = 3918.0
                    dec_col = next((c for c in df_final.columns if "decision" in c.lower()), None)
                    if dec_col is None:
                        dec_col = next((c for c in df_final.columns if "dec" in c.lower() and "approv" not in c.lower()), None)
                    if dec_col is None:
                        fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)
                        ax.text(0.5, 0.5, "Decision column not found", ha="center", va="center", fontsize=20)
                        ax.axis("off"); return fig
                    base = df_final[
                        (df_final["Year"] == year) &
                        (df_final["Month"].between(1, month)) &
                        (df_final[dec_col].astype(str).str.strip().str.lower() == "return material from customer")
                    ].copy()
                    if base.empty:
                        fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)
                        ax.text(0.5, 0.5, "No Customer Returns found", ha="center", va="center", fontsize=20)
                        ax.axis("off"); return fig
                    base["_Reason"] = base["Reason"].astype(str).str.strip()
                    cm_counts  = base[base["Month"] == month]["_Reason"].value_counts()
                    ytd_counts = base["_Reason"].value_counts()
                    all_r      = pd.Index(ytd_counts.index).union(cm_counts.index)
                    summ = pd.DataFrame({
                        "CM_count":  cm_counts.reindex(all_r, fill_value=0),
                        "YTD_count": ytd_counts.reindex(all_r, fill_value=0),
                    }).sort_values("YTD_count", ascending=False)
                    if top_n:
                        summ = summ.head(int(top_n))
                    summ["CM_cost"]  = summ["CM_count"]  * COST_PER_CLAIM
                    summ["YTD_cost"] = summ["YTD_count"] * COST_PER_CLAIM
                    def fmt_sar(v):
                        if v >= 1_000_000: return f"SAR {v/1_000_000:.1f}M"
                        if v >= 1000:      return f"SAR {int(v/1000)}K"
                        return f"SAR {int(v)}"
                    x = np.arange(len(summ)); w = 0.35
                    fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)
                    b1 = ax.bar(x - w/2, summ["CM_cost"],  w, color="#006394", label=f"Current Month  [n={int(summ['CM_count'].sum())}]")
                    b2 = ax.bar(x + w/2, summ["YTD_cost"], w, color="#C1A02E", label=f"YTD  [n={int(summ['YTD_count'].sum())}]")
                    ymax = summ[["CM_cost","YTD_cost"]].to_numpy().max() if summ[["CM_cost","YTD_cost"]].to_numpy().max() > 0 else 1
                    for bars_g in [b1, b2]:
                        for bar in bars_g:
                            h = bar.get_height()
                            if h > 0:
                                ax.text(bar.get_x() + bar.get_width()/2, h + ymax * 0.015,
                                        fmt_sar(h), ha="center", va="bottom",
                                        fontsize=9, rotation=90, color="#333333")
                    ax.set_xticks(x)
                    ax.set_xticklabels([fill(r, 20) for r in summ.index.tolist()], fontsize=10, rotation=25, ha="right")
                    ax.set_ylabel("Estimated Cost (SAR)  [@ SAR 3,918 / claim]")
                    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
                    ax.grid(False)
                    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.18), ncol=2, frameon=False)
                    fig.subplots_adjust(bottom=0.25)
                    return fig

                def slide_classification_guide_fig():
                    import textwrap as _tw
                    categories = [
                        {
                            "label": "Quality",
                            "color": "#006394",
                            "text_color": "white",
                            "lines": [
                                "Physical or functional",
                                "defects from production.",
                                "",
                                "Score Cracking, Ink",
                                "rubbing, Delamination,",
                                "Warped Sheets, Poor Die",
                                "Cutting, Wet boards,",
                                "GSM Downgrade.",
                            ],
                        },
                        {
                            "label": "Service",
                            "color": "#C1A02E",
                            "text_color": "white",
                            "lines": [
                                "Failures in delivery,",
                                "invoicing, or admin.",
                                "",
                                "Wrong Item Delivered,",
                                "Deviation from delivery",
                                "Schedule, Incorrect",
                                "Sales Contract Pricing,",
                                "Wrong Unit Price.",
                            ],
                        },
                        {
                            "label": "Commercial",
                            "color": "#2E8449",
                            "text_color": "white",
                            "lines": [
                                "Pre-agreed commercial",
                                "arrangements — not a",
                                "product or service",
                                "failure.",
                                "",
                                "Root Cause:",
                                "  Pre-Agreement",
                                "Reasons: Sales Discount,",
                                "  FOC, Tools reimb.",
                            ],
                        },
                        {
                            "label": "Invalid",
                            "color": "#B8A040",
                            "text_color": "white",
                            "lines": [
                                "Excluded from KPIs.",
                                "",
                                "• Reason Type = 0",
                                "  or blank",
                                "• Physical Status:",
                                "  Baled / Plastic Waste",
                                "• Gen. Category not",
                                "  in allowed list",
                                "• Reason = Invalid",
                            ],
                        },
                    ]
                    fig, ax = plt.subplots(figsize=(13.33, 7.5), dpi=300)
                    ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")
                    fig.patch.set_facecolor("white")
                    n = len(categories)
                    pad = 0.02
                    card_w = (1 - pad * (n + 1)) / n
                    header_h = 0.13
                    body_top = 0.88
                    body_bot = 0.04
                    for i, cat in enumerate(categories):
                        x0 = pad + i * (card_w + pad)
                        x1 = x0 + card_w
                        # Header rectangle
                        ax.add_patch(plt.Rectangle(
                            (x0, body_top), card_w, header_h,
                            transform=ax.transAxes,
                            facecolor=cat["color"], edgecolor="none", zorder=2,
                            clip_on=False
                        ))
                        # Header label
                        ax.text(
                            (x0 + x1) / 2, body_top + header_h / 2,
                            cat["label"],
                            transform=ax.transAxes,
                            ha="center", va="center",
                            fontsize=16, fontweight="bold",
                            color=cat["text_color"], zorder=3
                        )
                        # Body rectangle
                        ax.add_patch(plt.Rectangle(
                            (x0, body_bot), card_w, body_top - body_bot,
                            transform=ax.transAxes,
                            facecolor="#F5F5F5",
                            edgecolor=cat["color"], linewidth=2,
                            zorder=1, clip_on=False
                        ))
                        # Body text — line by line, evenly spaced inside the box
                        lines = cat["lines"]
                        n_lines = len(lines)
                        line_h = (body_top - body_bot - 0.04) / max(n_lines, 1)
                        for j, line in enumerate(lines):
                            y_pos = body_top - 0.02 - (j + 0.5) * line_h
                            ax.text(
                                (x0 + x1) / 2, y_pos,
                                line,
                                transform=ax.transAxes,
                                ha="center", va="center",
                                fontsize=9, color="#333333", zorder=4
                            )
                    plt.tight_layout()
                    return fig

                pdf_slides.append({"title": "Commercial Reasons", "section": True})
                pdf_slides.append({
                    "title": "Commercial CRMs — Count by Reason (CM vs YTD)",
                    "fig": slide_commercial_count_fig(selected_year, selected_month),
                })
                pdf_slides.append({
                    "title": "Commercial CRMs — Cost Amount by Reason (CM vs YTD)",
                    "fig": slide_commercial_cost_fig(selected_year, selected_month),
                })
                pdf_slides.append({
                    "title": "Customer Returns — Estimated Cost by Reason",
                    "fig": slide_customer_return_cost_fig(selected_year, selected_month, TOPN_SERVICE_CM),
                })
                pdf_slides.append({
                    "title": "CRM Classification Guide",
                    "fig": slide_classification_guide_fig(),
                })

                # Intro slide — Cost of Quality
                pdf_slides.append({"title": "Cost of Quality", "section": True})
                # 24 - COQ CM (single month bar — reuse breakdown fig)
                if coq_ready:
                    pdf_slides.append({"title": f"Cost of Quality — {month_name}", "fig": slide_coq_breakdown_fig(selected_year, selected_month)})


                # PDF Export
                with st.spinner("⚙️ Building PDF..."):
                    pdf_buf = build_ppt_pdf(pdf_slides, dpi=300)
                st.success("✅ PDF ready!")
                st.download_button(
                    label="📥 Download PPT-style PDF",
                    data=pdf_buf,
                    file_name=f"QM_Dashboard_{selected_year}-{selected_month:02d}.pdf",
                    mime="application/pdf",
                )
                plt.close("all")

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
                    plt.close(_fig_t)
                    plt.close("all")

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
                _qm6_ncr_loaded["_Base_Date"] = pd.to_datetime(_qm6_ncr_loaded[CREATION_DATETIME_COL], errors="coerce")
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
            
                # Convert to DataFrame
                _prod_df = pd.DataFrame(_prod_rows.data or [])
            
                # Build lookup dict (used later for PPM)
                _prod_data = {
                    (r["year"], r["month"]): r["produced_qty"]
                    for r in (_prod_rows.data or [])
                }
            
                # ── Display production data ──
                if not _prod_df.empty:
                    _prod_df = _prod_df.sort_values(["year", "month"])
            
                    # Filter to selected year + range
                    _prod_filtered = _prod_df[
                        (_prod_df["year"] == int(_qm6_year)) &
                        (_prod_df["month"] >= _qm6_from_month) &
                        (_prod_df["month"] <= _qm6_to_month)
                    ]
            
                    st.markdown("### 🏭 Production Data for Selected Period")
                    st.dataframe(_prod_filtered, use_container_width=True)
            
            except Exception as e:
                _prod_data = {}
                st.warning(f"Could not load production data: {e}")

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
                            (_qm6_ncr_loaded[_qm6_dec_col_ncr].astype(str).str.strip().str.lower() == "shredding") &
                            (_qm6_ncr_loaded["Gen Categories"].astype(str).str.strip().str.lower() == "work in progress")    
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
                st.divider()
                st.markdown("### 💰 Cost Analysis by Defect")
                with st.spinner("Computing Cost Analysis..."):
            
                    # FINAL
                    _cost_final = _qm6_df_final[
                        (_qm6_df_final["Is_Valid"] == True) &
                        (_qm6_df_final["Complaint_Category"] == "Quality") &
                        (_qm6_df_final["Year"] == int(_qm6_year)) &
                        (_qm6_df_final["Month"] >= _qm6_from_month) &
                        (_qm6_df_final["Month"] <= _qm6_to_month) &
                        (_qm6_df_final[_dec_col].astype(str).str.strip().str.lower() == "credit note")
                    ].copy()
            
                    _cost_final["_Reason"] = _cost_final["Reason"].astype(str).str.strip()
                    _cost_final["_Cost"]   = pd.to_numeric(_cost_final["Cost Amount"], errors="coerce").fillna(0)
            
                    _cost_final_grp = _cost_final.groupby("_Reason")["_Cost"].sum()
            
                    # NCR
                    _cost_ncr = _qm6_ncr_loaded[
                        (_qm6_ncr_loaded["Year"] == int(_qm6_year)) &
                        (_qm6_ncr_loaded["Month"] >= _qm6_from_month) &
                        (_qm6_ncr_loaded["Month"] <= _qm6_to_month) &
                        (_qm6_ncr_loaded[_qm6_dec_col_ncr].astype(str).str.strip().str.lower() == "shredding") &
                        (_qm6_ncr_loaded["Gen Categories"].astype(str).str.strip().str.lower() == "work in progress")
                    ].copy()
            
                    _cost_ncr["_Reason"] = _cost_ncr["Reason"].astype(str).str.strip()
                    _cost_ncr["_Cost"]   = pd.to_numeric(_cost_ncr["Cost Amount"], errors="coerce").fillna(0)
            
                    _cost_ncr_grp = _cost_ncr.groupby("_Reason")["_Cost"].sum()
            
                    # Merge
                    _all_reasons_cost = set(_cost_final_grp.index) | set(_cost_ncr_grp.index)
            
                    _cost_rows = []
                    for r in sorted(_all_reasons_cost):
                        _f_cost = float(_cost_final_grp.get(r, 0))
                        _n_cost = float(_cost_ncr_grp.get(r, 0))
                        _total  = _f_cost + _n_cost
            
                        _cost_rows.append({
                            "Defect": r,
                            "NCR Cost": round(_n_cost, 2),
                            "Credit Note Cost": round(_f_cost, 2),
                            "Total Cost": round(_total, 2),
                        })
            
                    _cost_df = pd.DataFrame(_cost_rows).sort_values("Total Cost", ascending=False)
            
                    st.dataframe(_cost_df, use_container_width=True)
            
                                        # Pareto
                    # ── Pareto (Cost) ──
                    _cost_df_plot = _cost_df[_cost_df["Total Cost"] > 0].copy()
                    
                    if not _cost_df_plot.empty:
                        st.divider()
                        st.markdown(f"#### Pareto — Cost by Defect ({_period_lbl})")
                    
                        # Sort like PPM
                        _cost_df_plot = _cost_df_plot.sort_values("Total Cost", ascending=False).reset_index(drop=True)
                    
                        # Cumulative %
                        _cost_df_plot["Cumulative %"] = (
                            _cost_df_plot["Total Cost"].cumsum() / _cost_df_plot["Total Cost"].sum() * 100
                        )
                    
                        _fig_c, _ax_c = plt.subplots(figsize=(13.33, 7.5), dpi=150)
                        _xp = np.arange(len(_cost_df_plot))
                    
                        # Bars (same styling as PPM)
                        _bars_c = _ax_c.bar(
                            _xp,
                            _cost_df_plot["Total Cost"].values,
                            color="#006394",
                            width=0.6,
                            alpha=0.9
                        )
                    
                        # Value labels
                        for _bar, _val in zip(_bars_c, _cost_df_plot["Total Cost"].values):
                            _ax_c.text(
                                _bar.get_x() + _bar.get_width()/2,
                                _bar.get_height() + _cost_df_plot["Total Cost"].max()*0.01,
                                f"{_val:,.0f}",
                                ha="center",
                                va="bottom",
                                fontsize=9,
                                color="#000000"
                            )
                    
                        # Cumulative line
                        _ax_c2 = _ax_c.twinx()
                        _ax_c2.plot(
                            _xp,
                            _cost_df_plot["Cumulative %"].values,
                            color="#C1A02E",
                            linewidth=2.5,
                            marker="o",
                            markersize=5,
                            label="Cumulative %"
                        )
                    
                        # 80% line
                        _ax_c2.axhline(80, color="#DE201B", linewidth=1, linestyle="--", alpha=0.6)
                    
                        _ax_c2.set_ylabel("Cumulative %")
                        _ax_c2.set_ylim(0, 115)
                        _ax_c2.tick_params(axis="y")
                        _ax_c2.spines["top"].set_visible(False)
                    
                        # X-axis labels (same formatting!)
                        _ax_c.set_xticks(_xp)
                        _ax_c.set_xticklabels(
                            [fill(d, 16) for d in _cost_df_plot["Defect"]],
                            rotation=40,
                            ha="right",
                            fontsize=8
                        )
                    
                        _ax_c.set_ylabel("Cost")
                        _ax_c.spines["top"].set_visible(False)
                        _ax_c.spines["right"].set_visible(False)
                        _ax_c.grid(False)
                    
                        # Legend (same as PPM)
                        _l2, _lb2 = _ax_c2.get_legend_handles_labels()
                        _ax_c.legend(_l2, _lb2, loc="upper right", frameon=False, fontsize=10)
                    
                        plt.tight_layout()
                        st.image(fig_to_png_bytes(_fig_c))
                        plt.close(_fig_c)
        except Exception as e:
            st.error(f"Error: {e}")
