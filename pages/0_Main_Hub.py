# pages/0_Main_Hub.py - Plant Manager User Management
import os, sys
import streamlit as st
import pandas as pd

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils.supabase_client import get_supabase, get_supabase_admin

st.set_page_config(page_title="Main Hub", page_icon="⚙️", layout="wide")

# ── Auth guard: plant manager only ──
if "user" not in st.session_state:
    st.switch_page("app.py")

role = st.session_state.get("role", "member")
name = st.session_state.get("name", "User")

if role != "plant_manager":
    st.error("🔒 Access restricted to Plant Manager only.")
    st.stop()

supabase       = get_supabase()        # anon key  — for profiles/directory
supabase_admin = get_supabase_admin()  # service key — for auth user creation

col1, col2 = st.columns([5, 1])
with col1:
    st.markdown("# ⚙️ Main Hub")
    st.markdown(f"Logged in as **{name}** · `plant_manager`")
with col2:
    if st.button("🏠 Home", use_container_width=True):
        st.switch_page("pages/1_Home.py")

st.divider()

tab1, tab2 = st.tabs(["👥 User Management", "🏭 People Directory"])

PILLARS = ["ALL", "AM", "PM", "QM", "HSE", "FI", "ET"]
ROLES   = ["plant_manager", "pillar_leader", "coordinator", "member"]

# ════════════════════════════════════════
# TAB 1 — USER MANAGEMENT
# ════════════════════════════════════════
with tab1:

    # ── Load existing profiles ──
    try:
        profiles_res = supabase.table("profiles").select("*").order("full_name").execute()
        profiles = profiles_res.data or []
    except Exception as e:
        st.error(f"Could not load profiles: {e}")
        profiles = []

    # ── Current Users Table ──
    st.markdown("### 👥 Current Users")
    if profiles:
        df_profiles = pd.DataFrame(profiles)[["full_name","email","pillar","role"]].rename(columns={
            "full_name": "Name", "email": "Email", "pillar": "Pillar", "role": "Role"
        })
        st.dataframe(df_profiles, use_container_width=True, hide_index=True)
    else:
        st.info("No users found.")

    st.divider()

    # ── Add New User ──
    st.markdown("### ➕ Add New User")
    st.caption("This creates a Supabase auth account and profile in one step.")

    with st.form("hub_add_user_form"):
        c1, c2 = st.columns(2)
        new_name     = c1.text_input("Full Name *")
        new_email    = c2.text_input("Email *")
        new_password = c1.text_input("Temporary Password *", type="password")
        new_pillar   = c2.selectbox("Pillar / Department", PILLARS)
        new_role     = c1.selectbox("Role", ROLES, index=3)

        submitted = st.form_submit_button("➕ Create User", type="primary")
        if submitted:
            if not new_name or not new_email or not new_password:
                st.error("Please fill in Name, Email and Password.")
            else:
                try:
                    # Create auth user via service role (admin) client
                    res = supabase_admin.auth.admin.create_user({
                        "email": new_email,
                        "password": new_password,
                        "email_confirm": True,
                    })
                    uid = res.user.id
                    # Insert profile using anon client
                    supabase.table("profiles").insert({
                        "id": uid,
                        "full_name": new_name,
                        "email": new_email,
                        "pillar": new_pillar,
                        "role": new_role,
                    }).execute()
                    st.success(f"✅ User **{new_name}** created! They can log in with `{new_email}`.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Failed to create user: {str(e)}")

    st.divider()

    # ── Edit Existing User ──
    st.markdown("### ✏️ Edit User Access")
    if profiles:
        user_map = {f"{p['full_name']} ({p.get('email','')})": p for p in profiles}
        sel_user_key = st.selectbox("Select user to edit", list(user_map.keys()), key="hub_edit_sel")
        sel_user = user_map[sel_user_key]

        with st.form("hub_edit_user_form"):
            c1, c2 = st.columns(2)
            edit_name   = c1.text_input("Full Name",  value=sel_user.get("full_name",""))
            edit_email  = c2.text_input("Email",      value=sel_user.get("email",""), disabled=True)
            edit_pillar = c1.selectbox("Pillar", PILLARS,
                index=PILLARS.index(sel_user.get("pillar","ALL")) if sel_user.get("pillar") in PILLARS else 0)
            edit_role   = c2.selectbox("Role", ROLES,
                index=ROLES.index(sel_user.get("role","member")) if sel_user.get("role") in ROLES else 3)

            col_save, col_del = st.columns([3,1])
            save_clicked   = col_save.form_submit_button("💾 Save Changes", type="primary", use_container_width=True)
            delete_clicked = col_del.form_submit_button("🗑️ Delete User",  use_container_width=True)

        if save_clicked:
            try:
                supabase.table("profiles").update({
                    "full_name": edit_name,
                    "pillar":    edit_pillar,
                    "role":      edit_role,
                }).eq("id", sel_user["id"]).execute()
                st.success(f"✅ Updated **{edit_name}**!")
                st.rerun()
            except Exception as e:
                st.error(f"Update failed: {e}")

        if delete_clicked:
            try:
                supabase_admin.auth.admin.delete_user(sel_user["id"])
                supabase.table("profiles").delete().eq("id", sel_user["id"]).execute()
                st.success(f"🗑️ Deleted user **{sel_user['full_name']}**.")
                st.rerun()
            except Exception as e:
                st.error(f"Delete failed: {e}")

# ════════════════════════════════════════
# TAB 2 — PEOPLE DIRECTORY
# ════════════════════════════════════════
with tab2:
    st.markdown("### 🏭 People Directory")
    st.caption("This list is used as the owner dropdown in Action Plans and Projects across all pillars.")

    try:
        people_res = supabase.table("people_directory").select("*").order("name").execute()
        people = people_res.data or []
    except Exception as e:
        st.error(f"Could not load people directory: {e}")
        people = []

    if people:
        df_people = pd.DataFrame(people)[["name","department","email"]].rename(columns={
            "name":"Name","department":"Department","email":"Email"
        })
        st.dataframe(df_people, use_container_width=True, hide_index=True)
    else:
        st.info("No people in directory yet. Add them below.")

    st.divider()
    st.markdown("#### ➕ Add Person")
    with st.form("hub_add_person_form"):
        c1, c2, c3 = st.columns(3)
        p_name  = c1.text_input("Name *")
        p_dept  = c2.selectbox("Department / Pillar", PILLARS + ["Other"])
        p_email = c3.text_input("Email (optional)")
        if st.form_submit_button("➕ Add", type="primary"):
            if not p_name:
                st.error("Name is required.")
            else:
                try:
                    supabase.table("people_directory").insert({
                        "name": p_name, "department": p_dept, "email": p_email or ""
                    }).execute()
                    st.success(f"✅ Added {p_name}!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Failed: {e}")

    if people:
        st.markdown("#### 🗑️ Remove Person")
        del_map = {p["name"]: p["id"] for p in people}
        sel_del = st.selectbox("Select to remove", list(del_map.keys()), key="hub_del_person")
        if st.button("🗑️ Remove", key="hub_btn_del_person"):
            try:
                supabase.table("people_directory").delete().eq("id", del_map[sel_del]).execute()
                st.success(f"Removed {sel_del}"); st.rerun()
            except Exception as e:
                st.error(f"Failed: {e}")
