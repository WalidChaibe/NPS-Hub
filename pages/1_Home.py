# pages/1_Home.py - NPS Hub Home Screen
import streamlit as st
from utils.supabase_client import get_supabase

st.set_page_config(
    page_title="NPS Hub - Home",
    page_icon="🏭",
    layout="wide"
)

# ── Auth guard ──
if "user" not in st.session_state:
    st.switch_page("app.py")

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

# ── Supabase ──
supabase = get_supabase()

# ── Notification check (once per session) ──
if "notif_checked" not in st.session_state:
    try:
        check_action_plan_notifications(supabase)
    except Exception:
        pass
    st.session_state["notif_checked"] = True

# ── Pillar definitions ──
PILLARS = [
    {"id": "AM",  "name": "Autonomous Maintenance",      "color": "#E67E22", "icon": "🔧", "image": "https://sjcwzbftzpfylwdqiknh.supabase.co/storage/v1/object/public/asset/AM%20Pillar%20Image.jpg"},
    {"id": "PM",  "name": "Planned Maintenance",          "color": "#3498DB", "icon": "📅", "image": "https://sjcwzbftzpfylwdqiknh.supabase.co/storage/v1/object/public/asset/PM%20Pillar%20Image.png"},
    {"id": "QM",  "name": "Quality Maintenance",          "color": "#E74C3C", "icon": "✅", "image": "https://sjcwzbftzpfylwdqiknh.supabase.co/storage/v1/object/public/asset/QM%20Pillar%20Image.jpg"},
    {"id": "HSE", "name": "Health, Safety & Environment", "color": "#27AE60", "icon": "🦺", "image": "https://sjcwzbftzpfylwdqiknh.supabase.co/storage/v1/object/public/asset/HSE%20Pillar%20Image.jpg"},
    {"id": "FI",  "name": "Focused Improvement",          "color": "#9B59B6", "icon": "💡", "image": "https://sjcwzbftzpfylwdqiknh.supabase.co/storage/v1/object/public/asset/FI%20Pillar%20Image.jpg"},
    {"id": "ET",  "name": "Education & Training",         "color": "#1ABC9C", "icon": "🎓", "image": "https://sjcwzbftzpfylwdqiknh.supabase.co/storage/v1/object/public/asset/E%26T%20Pillar%20Image.png"},
]

role   = st.session_state.get("role", "member")
pillar = st.session_state.get("pillar", "ALL")
name   = st.session_state.get("name", "User")

# ── Header ──
col1, col2, col3 = st.columns([5, 1, 1])
with col1:
    st.markdown("# 🏭 NPS Hub")
    st.markdown(f"Welcome back, **{name}** &nbsp;|&nbsp; Role: `{role}` &nbsp;|&nbsp; Pillar: `{pillar}`")
with col2:
    render_bell(supabase, st.session_state["user"])
with col3:
    if st.button("🚪 Logout", use_container_width=True):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.switch_page("app.py")

st.divider()
st.markdown("### Select a Pillar")

# ── Pillar cards ──
cols = st.columns(3)
for i, p in enumerate(PILLARS):
    can_access = (role == "plant_manager") or (pillar == "ALL") or (pillar == p["id"])
    with cols[i % 3]:
        st.markdown(f"""
            <div style="
                border: 1px solid {p['color']}55;
                border-radius: 10px;
                padding: 16px;
                margin-bottom: 12px;
                background: {p['color']}11;
                display: flex;
                align-items: center;
                justify-content: space-between;
            ">
                <div>
                    <div style="font-size: 28px">{p['icon']}</div>
                    <div style="font-weight: 700; font-size: 15px; margin: 8px 0 4px">{p['name']}</div>
                </div>
                <img src="{p['image']}" style="
                    height: 80px; width: 80px;
                    object-fit: contain; background: transparent;
                ">
            </div>
        """, unsafe_allow_html=True)
        if st.button(f"Open {p['id']}", key=f"btn_{p['id']}", use_container_width=True):
            st.session_state["active_pillar"] = p["id"]
            st.switch_page(f"pages/{i+2}_{p['id']}.py")

st.divider()
st.caption("Napco National · NPS Hub · TPM Program")
