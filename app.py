# app.py - NPS Hub Login Page
import streamlit as st
from utils.supabase_client import get_supabase

st.set_page_config(
    page_title="NPS Hub",
    page_icon="🏭",
    layout="centered"
)

# ── Hide the default streamlit sidebar nav until logged in ──
if "user" not in st.session_state:
    st.markdown("""
        <style>
        [data-testid="stSidebarNav"] { display: none; }
        </style>
    """, unsafe_allow_html=True)

# ── If already logged in, go to home ──
if "user" in st.session_state:
    st.switch_page("pages/1_Home.py")

# ── Login UI ──
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    st.image("https://i.imgur.com/placeholder.png", width=80)  # replace with your logo later
    st.markdown("## 🏭 NPS Hub")
    st.markdown("**Napco Production System — TPM Program**")
    st.divider()

    st.markdown("### Sign In")
    email    = st.text_input("Email", placeholder="you@napconational.com")
    password = st.text_input("Password", type="password", placeholder="••••••••")

    if st.button("Login", use_container_width=True, type="primary"):
        if not email or not password:
            st.error("Please enter your email and password.")
        else:
            try:
                supabase = get_supabase()
                res = supabase.auth.sign_in_with_password({
                    "email": email,
                    "password": password
                })

                # Fetch the user's profile (role + pillar)
                profile = supabase.table("profiles") \
                    .select("*") \
                    .eq("id", res.user.id) \
                    .single() \
                    .execute()

                # Save to session
                st.session_state["user"]     = res.user
                st.session_state["profile"]  = profile.data
                st.session_state["role"]     = profile.data["role"]
                st.session_state["pillar"]   = profile.data["pillar"]
                st.session_state["name"]     = profile.data["full_name"]

                st.success(f"Welcome back, {profile.data['full_name']}!")
                st.switch_page("pages/1_Home.py")

            except Exception as e:
                st.error(f"Error: {str(e)}")

    st.divider()
    st.caption("Access issues? Contact your NPS Administrator.")
