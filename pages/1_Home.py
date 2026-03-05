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
    {"id": "AM",  "name": "Autonomous Maintenance",     "color": "#E67E22", "image": "https://sjcwzbftzpfylwdqiknh.supabase.co/storage/v1/object/public/asset/AM%20Pillar%20Image.jpg"},
    {"id": "PM",  "name": "Planned Maintenance",         "color": "#3498DB", "image": "https://sjcwzbftzpfylwdqiknh.supabase.co/storage/v1/object/public/asset/PM%20Pillar%20Image.png"},
    {"id": "QM",  "name": "Quality Maintenance",         "color": "#E74C3C", "image": "https://sjcwzbftzpfylwdqiknh.supabase.co/storage/v1/object/public/asset/QM%20Pillar%20Image.jpg"},
    {"id": "HSE", "name": "Health, Safety & Environment","color": "#27AE60", "image": "https://sjcwzbftzpfylwdqiknh.supabase.co/storage/v1/object/public/asset/HSE%20Pillar%20Image.jpg"},
    {"id": "FI",  "name": "Focused Improvement",         "color": "#9B59B6", "image": "https://sjcwzbftzpfylwdqiknh.supabase.co/storage/v1/object/public/asset/FI%20Pillar%20Image.jpg"},
    {"id": "ET",  "name": "Education & Training",        "color": "#1ABC9C", "image": "https://sjcwzbftzpfylwdqiknh.supabase.co/storage/v1/object/public/asset/E%26T%20Pillar%20Image.png"},
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
    access_label = "✏️ Full Access" if can_edit else "👁️ View Only"
    access_color = "#0d68a3" if can_edit else "#888888"

    with cols[i % 3]:
        st.markdown(f"""
            <div style="
                border: 1px solid #30363D;
                border-radius: 10px;
                overflow: hidden;
                margin-bottom: 12px;
                background: #1a1a1a;
            ">
                <div style="
                    width: 100%;
                    height: 180px;
                    overflow: hidden;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    background: #111;
                ">
                    <img src="{p['image']}" style="width:100%; height:100%; object-fit:contain;">
                </div>
                <div style="padding: 10px 14px;">
                    <div style="font-size: 11px; color: {access_color}; font-weight: 600;">{access_label}</div>
                </div>
            </div>
        """, unsafe_allow_html=True)
        if st.button(f"Open", key=f"btn_{p['id']}", use_container_width=True):
            st.session_state["active_pillar"] = p["id"]
            st.switch_page(f"pages/{i+2}_{p['id']}.py")

st.divider()
st.caption("Napco National · NPS Hub · TPM Program")
