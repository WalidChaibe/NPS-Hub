# pages/1_Home.py - NPS Hub Home Screen
import streamlit as st

st.set_page_config(
    page_title="NPS Hub - Home",
    page_icon="🏭",
    layout="wide"
)

# ── Auth guard - redirect to login if not logged in ──
if "user" not in st.session_state:
    st.switch_page("app.py")

# ── Pillar definitions ──
PILLARS = [
    {"id": "AM",  "name": "Autonomous Maintenance",     "icon": "🔧", "color": "#E67E22"},
    {"id": "PM",  "name": "Planned Maintenance",         "icon": "📅", "color": "#3498DB"},
    {"id": "QM",  "name": "Quality Maintenance",         "icon": "✅", "color": "#E74C3C"},
    {"id": "HSE", "name": "Health, Safety & Environment","icon": "🦺", "color": "#27AE60"},
    {"id": "FI",  "name": "Focused Improvement",         "icon": "💡", "color": "#9B59B6"},
    {"id": "ET",  "name": "Education & Training",        "icon": "🎓", "color": "#1ABC9C"},
]

role   = st.session_state.get("role", "member")
pillar = st.session_state.get("pillar", "ALL")
name   = st.session_state.get("name", "User")

# ── Header ──
col1, col2 = st.columns([5, 1])
with col1:
    st.markdown(f"# 🏭 NPS Hub")
    st.markdown(f"Welcome back, **{name}** &nbsp;|&nbsp; Role: `{role}` &nbsp;|&nbsp; Pillar: `{pillar}`")
with col2:
    if st.button("🚪 Logout", use_container_width=True):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.switch_page("app.py")

st.divider()
st.markdown("### Select a Pillar")

# ── Pillar cards ──
cols = st.columns(3)
for i, p in enumerate(PILLARS):
    can_edit = (role == "plant_manager") or (role == "pillar_leader" and pillar == p["id"])

    with cols[i % 3]:
        badge = "✏️ Edit" if can_edit else "👁️ View only"
        st.markdown(f"""
            <div style="
                border: 1px solid {p['color']}55;
                border-radius: 10px;
                padding: 20px;
                margin-bottom: 12px;
                background: {p['color']}11;
            ">
                <div style="font-size: 28px">{p['icon']}</div>
                <div style="font-weight: 700; font-size: 15px; margin: 8px 0 4px">{p['name']}</div>
                <div style="font-size: 11px; color: {p['color']}; font-weight: 600">{badge}</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button(f"Open {p['id']}", key=f"btn_{p['id']}", use_container_width=True):
            st.session_state["active_pillar"] = p["id"]
            st.switch_page(f"pages/{i+2}_{p['id']}.py")

st.divider()
st.caption("Napco National · NPS Hub · TPM Program")
