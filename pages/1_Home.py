# pages/1_Home.py - NPS Hub Home Screen
import streamlit as st

st.set_page_config(
    page_title="NPS Hub - Home",
    page_icon="🏭",
    layout="wide"
)

# ── Auth guard ──
if "user" not in st.session_state:
    st.switch_page("app.py")

# ── Pillar definitions ──
PILLARS = [
    {"id": "AM",  "name": "Autonomous Maintenance",     "color": "#E67E22", "icon": "🔧", "image": "https://sjcwzbftzpfylwdqiknh.supabase.co/storage/v1/object/public/asset/AM%20Pillar%20Image.jpg"},
    {"id": "PM",  "name": "Planned Maintenance",         "color": "#3498DB", "icon": "📅", "image": "https://sjcwzbftzpfylwdqiknh.supabase.co/storage/v1/object/public/asset/PM%20Pillar%20Image.png"},
    {"id": "QM",  "name": "Quality Maintenance",         "color": "#E74C3C", "icon": "✅", "image": "https://sjcwzbftzpfylwdqiknh.supabase.co/storage/v1/object/public/asset/QM%20Pillar%20Image.jpg"},
    {"id": "HSE", "name": "Health, Safety & Environment","color": "#27AE60", "icon": "🦺", "image": "https://sjcwzbftzpfylwdqiknh.supabase.co/storage/v1/object/public/asset/HSE%20Pillar%20Image.jpg"},
    {"id": "FI",  "name": "Focused Improvement",         "color": "#9B59B6", "icon": "💡", "image": "https://sjcwzbftzpfylwdqiknh.supabase.co/storage/v1/object/public/asset/FI%20Pillar%20Image.jpg"},
    {"id": "ET",  "name": "Education & Training",        "color": "#1ABC9C", "icon": "🎓", "image": "https://sjcwzbftzpfylwdqiknh.supabase.co/storage/v1/object/public/asset/E%26T%20Pillar%20Image.png"},
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
                    height: 80px;
                    width: 80px;
                    object-fit: contain;
                    background: transparent;
                ">
            </div>
        """, unsafe_allow_html=True)
        if st.button(f"Open {p['id']}", key=f"btn_{p['id']}", use_container_width=True):
            st.session_state["active_pillar"] = p["id"]
            st.switch_page(f"pages/{i+2}_{p['id']}.py")

st.divider()
st.caption("Napco National · NPS Hub · TPM Program")
