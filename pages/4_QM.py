# pages/4_QM.py - Quality Maintenance Pillar
import os, sys
from datetime import datetime
from urllib.parse import quote

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils.supabase_client import get_supabase

st.set_page_config(page_title="QM Pillar", page_icon="✅", layout="wide")

# ── Auth guard ──
if "user" not in st.session_state:
    st.switch_page("app.py")

supabase = get_supabase()
role     = st.session_state.get("role", "member")
pillar   = st.session_state.get("pillar", "ALL")
name     = st.session_state.get("name", "User")
can_edit = (role == "plant_manager") or (role == "pillar_leader" and pillar == "QM")

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
        {"name": "Entry", "icon": "⚪", "focus": "QM by basic methods. Identification of Q factors.",
         "criteria": [
             "QM Pillar and targets defined (Cost of non-Quality, CC, Internal Rejections)",
             "Basic Root Cause Analysis implemented",
             "Action plans generated to prevent defect re-occurrence",
             "Identification of Q factors (Machine, Material, Manpower, Method)",
             "QA Matrix tool developed",
             "KPIs monitored and suggestions for improvement available",
             "Customer Complaints trend is decreasing"]},
        {"name": "Bronze", "icon": "🥉", "focus": "Reduce Waste and Complaints through a Quality system.",
         "criteria": [
             "Analysis implemented for all Q factors",
             "QA Matrix used to tackle high risk defects",
             "Action plans effective (IR and CC data reduced by 10%)",
             "Accurate and updated MPCC for all FGs",
             "Quality skills assessed and improved for all shopfloor employees",
             "QC council meetings implemented",
             "Shopfloor/QC/MT employees have basic knowledge on Q factors"]},
        {"name": "Silver", "icon": "🥈", "focus": "Quality Awareness is clear.",
         "criteria": [
             "Reduction of NCR & CC ratio by 10% vs last year",
             "Evidence of 2 major Quality improvement projects",
             "QA matrix showing decrease in RPN of 50% of defects",
             "Cost of Poor Quality actual equals target COPQ",
             "Evidence of shopfloor involvement in NCR and CC RCAs",
             "Quality circles used to transfer knowledge",
             "Clear evidence of Q Points for machines, materials, methods, men and measurement"]},
        {"name": "Gold", "icon": "🥇", "focus": "Quality is embedded in the BU culture.",
         "criteria": [
             "Decreasing monthly trend of NCR & CC ratio",
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
# TAB 2 — QUALITY INDICATORS (embedded)
# ════════════════════════════════════════
with tab2:
    st.markdown("### 📈 CRM Quality Dashboard")
    st.caption("Full CRM dashboard — upload your files and generate charts directly below.")
    components.iframe(
        "https://crm-app-dashboard-vdv2rw68ah2gxabwxnofrr.streamlit.app/?embed=true",
        height=900,
        scrolling=True
    )

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
            col1, col2 = st.columns([3, 1])
            with col1:
                new_doc_name = st.text_input("Document Name", key="qm_new_doc_name")
                new_doc_desc = st.text_input("Description (optional)", key="qm_new_doc_desc")
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
                del_map = {r['doc_name']: r['id'] for r in req_docs.data}
                col1, col2 = st.columns([3, 1])
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
            doc_desc = req.get('description', '')
            uploads  = [u for u in (all_uploads.data or []) if u['doc_type'] == doc_name]
            has_up   = len(uploads) > 0

            st.markdown(f"""
                <div style="border:1px solid {'#2ea04355' if has_up else '#f8514955'};border-radius:10px;
                    padding:16px 20px;margin-bottom:14px;background:{'#2ea04308' if has_up else '#f8514908'};">
                    <div style="display:flex;align-items:center;gap:10px;">
                        <span style="font-size:18px">{'✅' if has_up else '❌'}</span>
                        <span style="font-weight:700;font-size:16px;">{doc_name}</span>
                        <span style="font-size:11px;color:{'#2ea043' if has_up else '#f85149'};margin-left:auto;">
                            {len(uploads)} file{'s' if len(uploads) != 1 else ''} uploaded</span>
                    </div>
                    {f'<div style="font-size:12px;color:#8B949E;margin-top:4px;">{doc_desc}</div>' if doc_desc else ''}
                </div>
            """, unsafe_allow_html=True)

            for upload in uploads:
                ext = upload['file_name'].split('.')[-1].lower()
                file_url = f"https://sjcwzbftzpfylwdqiknh.supabase.co/storage/v1/object/public/asset/{quote(upload['file_path'], safe='/')}"
                col1, col2, col3 = st.columns([4, 2, 1])
                with col1:
                    if ext in ['png', 'jpg', 'jpeg']:
                        st.image(file_url, width=200)
                    else:
                        st.markdown(f"📄 [{upload['file_name']}]({file_url})")
                    st.caption(f"Uploaded by {upload['uploaded_by']} · {pd.to_datetime(upload['created_at']).strftime('%d %b %Y %H:%M')}")
                with col3:
                    if can_edit and st.button("🗑️", key=f"qm_del_file_{upload['id']}"):
                        supabase.table("qm_documents").delete().eq("id", upload['id']).execute()
                        st.success("Deleted!")
                        st.rerun()

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
                                st.success("✅ Uploaded!")
                                st.rerun()
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
                act_map = {f"{r['action'][:40]} — {r['owner']}": r['id'] for r in actions.data}
                col1, col2 = st.columns([3, 1])
                with col1:
                    sel_act = st.selectbox("Select action", list(act_map.keys()), key="qm_sel_act")
                with col2:
                    new_act_s = st.selectbox("Status", ["Open","In Progress","Done"], key="qm_act_status")
                if st.button("✅ Update Status", key="qm_update_act"):
                    supabase.table("qm_action_plans").update({"status": new_act_s}).eq("id", act_map[sel_act]).execute()
                    st.success("Updated!")
                    st.rerun()

                st.markdown("#### 🗑️ Delete")
                del_act_map = {f"{r['action'][:40]} — {r['owner']}": r['id'] for r in actions.data}
                sel_del_act = st.selectbox("Select to delete", list(del_act_map.keys()), key="qm_del_act")
                if st.button("🗑️ Delete Action Plan", key="qm_btn_del_act"):
                    supabase.table("qm_action_plans").delete().eq("id", del_act_map[sel_del_act]).execute()
                    st.success("Deleted!")
                    st.rerun()
        else:
            st.info("No action plans yet.")
    except Exception as e:
        st.error(f"Could not load action plans: {str(e)}")

    if can_edit:
        st.divider()
        st.markdown("#### ➕ Add Action Plan")
        act_text = st.text_input("Action", key="qm_act_text")
        own_text = st.text_input("Owner",  key="qm_own_text")
        due_d    = st.date_input("Due Date", key="qm_due_date")
        stat_ap  = st.selectbox("Status", ["Open","In Progress","Done"], key="qm_new_act_status")
        notes_ap = st.text_area("Notes", key="qm_notes_ap")
        if st.button("➕ Add Action Plan", type="primary", key="qm_add_act"):
            if not act_text or not own_text:
                st.error("Please fill in action and owner.")
            else:
                try:
                    supabase.table("qm_action_plans").insert({
                        "action": act_text, "owner": own_text,
                        "due_date": str(due_d), "status": stat_ap, "notes": notes_ap,
                    }).execute()
                    st.success("✅ Action plan added!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {str(e)}")
