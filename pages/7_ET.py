# pages/7_ET.py - Education & Training Pillar
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import io
import tempfile
import os
from fpdf import FPDF
from utils.supabase_client import get_supabase
from urllib.parse import quote

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
        res = sb.table("et_action_plans").select("*").in_("status", ["Open","In Progress"]).execute()
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
                create_notification(sb, uid, "🔴 Overdue — ET Action Plan",
                    f'"{plan.get("action","")[:60]}" was due {due_date.strftime("%d %b %Y")} and is still open.',
                    "overdue", plan["id"])
            elif days_left <= 2:
                create_notification(sb, uid, "🟡 Due Soon — ET Action Plan",
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

st.set_page_config(page_title="E&T Pillar", page_icon="🎓", layout="wide")

# ── Auth guard ──
if "user" not in st.session_state:
    st.switch_page("app.py")

supabase = get_supabase()
role     = st.session_state.get("role", "member")
pillar   = st.session_state.get("pillar", "ALL")
name     = st.session_state.get("name", "User")

can_edit = (role == "plant_manager") or (role == "pillar_leader" and pillar == "ET")

# ── Notification check ──
if "notif_checked" not in st.session_state:
    try:
        check_action_plan_notifications(supabase)
    except Exception:
        pass
    st.session_state["notif_checked"] = True

# ── Header ──
col1, col2, col3 = st.columns([5, 1, 1])
with col1:
    st.markdown("# 🎓 Education & Training")
    st.markdown(f"Logged in as **{name}** · `{role}` · {'✏️ Full Access' if can_edit else '👁️ View Only'}")
with col2:
    render_bell(supabase, st.session_state["user"])
with col3:
    if st.button("🏠 Home", use_container_width=True):
        st.switch_page("pages/1_Home.py")

st.divider()

# ── Tabs ──
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Maturity Level",
    "🔬 Competency Analysis",
    "📁 Documents",
    "📋 Request Training",
    "🎯 Action Plans",
    "📝 One Point Lessons"
])

# ════════════════════════════════════════
# TAB 1 — MATURITY LEVEL
# ════════════════════════════════════════
with tab1:
    st.markdown("### TPM E&T Maturity Level Tracker")
    st.markdown("Current Level: 🥇 **Gold**")
    st.progress(1.0)

    levels = [
        {
            "name": "Entry", "icon": "⚪",
            "focus": "E&T by the basic methods and tools.",
            "criteria": [
                "Shop floor employees assessed based on competency matrix",
                "Flexibility charts designed and used for operators",
                "Action plans developed for employee gaps",
                "OPLs developed to cover gaps",
                "Employees started taking courses in PLJ",
                "Toolbox talks targeting shopfloor gaps",
                "Basic E&T concept exists in BU management",
            ]
        },
        {
            "name": "Bronze", "icon": "🥉",
            "focus": "Focus E&T on NPS Gaps.",
            "criteria": [
                "10% increase in overall competency scores",
                "Flexibility charts covering all plant operators and machines",
                "Training split between NPS team and BU pillar teams",
                "E&T pillar team engaging in different training topics",
                "BU internal and cross-functional training culture exists",
                "50% of employees working on PLJ",
                "Technical competencies developed in more detail",
            ]
        },
        {
            "name": "Silver", "icon": "🥈",
            "focus": "Continuous E&T — Training culture spreading.",
            "criteria": [
                "10% increase in overall competency scores vs last year",
                "Flexibility charts optimized to cover gaps in any section",
                "TWI used to pass on knowledge between employees",
                "Skills Dojo (Training Hall / Gemba Room) in use",
                "BU internal trainings > external trainings",
                "At least 2 engineers passed the NPS Test",
                "50% of plant employees passing PLJ scores",
            ]
        },
        {
            "name": "Gold", "icon": "🥇",
            "focus": "Continuous Improvement — Kaizen Culture evident.",
            "criteria": [
                "10% increase in overall competency scores vs last year",
                "Detailed schedule for employee training covering all gaps",
                "100% of Grade 12+ employees passed NPS Test",
                "Process specialists developed",
                "Training effectiveness evaluated (linked to BD, NCR, LTI)",
            ]
        },
    ]

    for level in levels:
        is_current = level["name"] == "Gold"
        with st.expander(f"{level['icon']} {level['name']} — {level['focus']}", expanded=is_current):
            if is_current:
                st.success("✅ Current Level")
            for criterion in level["criteria"]:
                st.checkbox(criterion, value=is_current, key=f"lvl_{level['name']}_{criterion[:30]}")

# ════════════════════════════════════════
# TAB 2 — COMPETENCY ANALYSIS
# ════════════════════════════════════════
with tab2:
    st.markdown("### Competency Matrix Analysis")

    st.markdown("#### 📅 Analysis History")
    try:
        history = supabase.table("et_analysis_runs").select("*").order("created_at", desc=True).execute()
        if history.data:
            hist_df = pd.DataFrame(history.data)
            hist_df["created_at"] = pd.to_datetime(hist_df["created_at"]).dt.strftime("%d %b %Y %H:%M")
            st.dataframe(
                hist_df[["created_at", "run_by", "total_failures", "notes"]].rename(columns={
                    "created_at": "Date", "run_by": "Run By",
                    "total_failures": "Total Failures", "notes": "Notes"
                }),
                use_container_width=True
            )
            if can_edit:
                run_map = {f"{r['created_at']} — {r['notes'] or 'No notes'} ({r['total_failures']} failures)": r['id'] for r in history.data}
                selected_run = st.selectbox("Select run to delete", list(run_map.keys()), key="del_run")
                if st.button("🗑️ Delete Run", key="btn_del_run"):
                    supabase.table("et_analysis_runs").delete().eq("id", run_map[selected_run]).execute()
                    st.success("Run deleted!")
                    st.rerun()
        else:
            st.info("No analysis runs yet. Upload a file below to run the first analysis.")
    except Exception as e:
        st.error(f"Could not load history: {str(e)}")

    st.divider()

    if not can_edit:
        st.warning("👁️ View Only — you do not have permission to run analysis.")
    else:
        st.markdown("#### ▶️ Run New Analysis")
        notes = st.text_input("Notes for this run (optional)", placeholder="e.g. Monthly analysis - March 2026")
        uploaded_file = st.file_uploader("Upload Competency Matrix Excel", type=["xlsx"])

        if uploaded_file and st.button("🔬 Run Analysis", type="primary"):
            with st.spinner("Analysing competency matrix..."):
                try:
                    xls = pd.ExcelFile(uploaded_file)
                    requirement_df = pd.read_excel(xls, sheet_name='Requirement', dtype=str)
                    requirement_df.columns = requirement_df.columns.str.strip()
                    requirement_df['Grade'] = requirement_df['Grade'].astype(str).str.strip()
                    topics = list(requirement_df.columns[1:])
                    requirement_dict = requirement_df.set_index('Grade')[topics].to_dict(orient='index')
                    employee_list_df = pd.read_excel(xls, sheet_name='Employee List', dtype=str)
                    employee_list_df.columns = employee_list_df.columns.str.strip()
                    employee_list_df['Grade'] = employee_list_df['Grade'].astype(str).str.strip()
                    skill_sheets = [s for s in xls.sheet_names if s not in ['Requirement', 'Employee List']]
                    failures = []
                    for _, emp_row in employee_list_df.iterrows():
                        emp_name = str(emp_row['Employees']).strip()
                        emp_grade = emp_row['Grade']
                        if emp_grade not in requirement_dict:
                            continue
                        required_scores = requirement_dict[emp_grade]
                        for sheet in skill_sheets:
                            df = pd.read_excel(xls, sheet_name=sheet, header=None, dtype=str)
                            df[0] = df[0].astype(str).str.strip()
                            match_rows = df[df[0] == emp_name]
                            if not match_rows.empty:
                                row_index = match_rows.index[0]
                                for i, topic in enumerate(topics):
                                    required_value = required_scores.get(topic)
                                    if pd.isna(required_value) or required_value == '':
                                        continue
                                    try:
                                        required_value = float(required_value)
                                    except:
                                        continue
                                    employee_value = df.iloc[row_index, i+1]
                                    if pd.isna(employee_value) or str(employee_value) in ['', 'N/A', 'nan']:
                                        continue
                                    try:
                                        employee_value = float(employee_value)
                                    except:
                                        continue
                                    if employee_value < required_value:
                                        failures.append({
                                            'Employee': emp_name,
                                            'Grade': emp_grade,
                                            'Department': sheet,
                                            'Topic': topic,
                                            'Employee Score': employee_value,
                                            'Required Score': required_value
                                        })

                    failures_df = pd.DataFrame(failures)
                    total_failures = len(failures_df)
                    st.success(f"✅ Analysis complete — {total_failures} failures detected")

                    if not failures_df.empty:
                        failures_df['Section'] = failures_df['Department'].str.replace(" Employees", "", regex=False)
                        failures_df['Remarks'] = "Needs Training"
                        topic_counts       = failures_df['Topic'].value_counts()
                        section_counts     = failures_df['Section'].value_counts()
                        topic_grade_counts = failures_df.groupby(['Topic', 'Grade']).size().reset_index(name='Count')
                        heatmap_data       = failures_df.groupby(['Grade', 'Topic']).size().unstack(fill_value=0)
                        sns.set(style="whitegrid")
                        fig1, ax1 = plt.subplots(figsize=(9, 5))
                        topic_counts.plot(kind='bar', color="#0D68A3", ax=ax1)
                        ax1.set_title("Total Failures per Topic")
                        ax1.set_ylabel("Number of Employees")
                        plt.xticks(rotation=45)
                        for p in ax1.patches:
                            ax1.annotate(f'{int(p.get_height())}', (p.get_x() + p.get_width()/2., p.get_height()), ha='center', va='bottom', fontsize=10)
                        plt.tight_layout()
                        fig2, ax2 = plt.subplots(figsize=(12, 6))
                        tgs = topic_grade_counts.copy()
                        tgs['Grade'] = tgs['Grade'].astype(str)
                        sns.barplot(data=tgs, x='Topic', y='Count', hue='Grade', palette=["#0D68A3","#5DA9E9","#DE201B","#F06A6A","#6C757D"], ax=ax2)
                        ax2.set_title("Failures per Topic by Grade")
                        plt.xticks(rotation=45)
                        plt.tight_layout()
                        fig3, ax3 = plt.subplots(figsize=(9, 5))
                        section_counts.plot(kind='bar', color="#DE201B", ax=ax3)
                        ax3.set_title("Failures by Section")
                        ax3.set_ylabel("Number of Employees")
                        plt.xticks(rotation=45)
                        plt.tight_layout()
                        fig4, ax4 = plt.subplots(figsize=(12, 6))
                        hm = heatmap_data.copy()
                        hm.index = hm.index.astype(str)
                        sns.heatmap(hm, annot=True, fmt="d", cmap=sns.light_palette("#0D68A3", as_cmap=True), linewidths=0.5, ax=ax4)
                        ax4.set_title("Failure Matrix (Grade vs Topic)")
                        plt.tight_layout()
                        st.markdown("#### 📊 Results")
                        st.dataframe(failures_df, use_container_width=True)
                        c1, c2 = st.columns(2)
                        with c1:
                            st.pyplot(fig1)
                            st.pyplot(fig3)
                        with c2:
                            st.pyplot(fig2)
                            st.pyplot(fig4)

                        def fig_to_tmp(fig):
                            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
                            fig.savefig(tmp.name, format='png', dpi=100)
                            tmp.close()
                            return tmp.name

                        tmp1 = fig_to_tmp(fig1)
                        tmp2 = fig_to_tmp(fig2)
                        tmp3 = fig_to_tmp(fig3)
                        tmp4 = fig_to_tmp(fig4)
                        pdf = FPDF()
                        pdf.set_auto_page_break(auto=True, margin=10)
                        pdf.add_page()
                        pdf.set_font("Arial", "B", 18)
                        pdf.set_text_color(13, 104, 163)
                        pdf.cell(0, 10, "TRAINING NEEDS REPORT", ln=True, align="C")
                        pdf.ln(5)
                        pdf.set_font("Arial", "", 11)
                        pdf.set_text_color(0, 0, 0)
                        pdf.cell(0, 8, f"Generated: {datetime.now().strftime('%d %b %Y %H:%M')}   |   Run by: {name}", ln=True)
                        if notes:
                            pdf.cell(0, 8, f"Notes: {notes}", ln=True)
                        pdf.ln(5)
                        grouped = failures_df.groupby(['Grade', 'Topic'])['Employee'].apply(list).reset_index()
                        pdf.set_font("Arial", "B", 14)
                        pdf.set_text_color(222, 32, 27)
                        pdf.cell(0, 10, "Training Schedule", ln=True)
                        pdf.set_text_color(0, 0, 0)
                        current_grade = None
                        for _, row in grouped.iterrows():
                            if row['Grade'] != current_grade:
                                pdf.ln(4)
                                pdf.set_font("Arial", "B", 13)
                                pdf.cell(0, 8, f"Grade {row['Grade']}", ln=True)
                                current_grade = row['Grade']
                            pdf.set_font("Arial", "B", 11)
                            pdf.cell(0, 6, f"Topic: {row['Topic']}", ln=True)
                            pdf.set_font("Arial", "", 10)
                            for emp in row['Employee']:
                                pdf.cell(0, 5, f"- {emp} (Needs Training)", ln=True)
                            pdf.ln(2)
                        for tmp_path, label in [(tmp4, "Failure Matrix"), (tmp1, "Failures per Topic"), (tmp2, "Failures by Grade"), (tmp3, "Failures by Section")]:
                            pdf.add_page()
                            pdf.set_font("Arial", "B", 14)
                            pdf.set_text_color(13, 104, 163)
                            pdf.cell(0, 10, label, ln=True)
                            pdf.image(tmp_path, w=180)
                        pdf_bytes = bytes(pdf.output())
                        for f in [tmp1, tmp2, tmp3, tmp4]:
                            os.unlink(f)
                        st.download_button(
                            label="📥 Download PDF Report",
                            data=pdf_bytes,
                            file_name=f"Training_Needs_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                            mime="application/pdf"
                        )
                        supabase.table("et_analysis_runs").insert({
                            "run_by": name,
                            "total_failures": total_failures,
                            "notes": notes or "",
                        }).execute()
                    else:
                        st.success("🎉 No failures detected — all employees meet requirements!")
                        supabase.table("et_analysis_runs").insert({
                            "run_by": name,
                            "total_failures": 0,
                            "notes": notes or "",
                        }).execute()
                except Exception as e:
                    st.error(f"Error during analysis: {str(e)}")

# ════════════════════════════════════════
# TAB 3 — DOCUMENTS
# ════════════════════════════════════════
with tab3:
    st.markdown("### 📁 Required Documents")
    try:
        req_docs = supabase.table("et_required_docs").select("*").order("created_at").execute()
        all_uploads = supabase.table("et_documents").select("*").order("created_at", desc=True).execute()
    except Exception as e:
        st.error(f"Could not load documents: {str(e)}")
        st.stop()

    if can_edit:
        with st.expander("➕ Add / Remove Required Document", expanded=False):
            col1, col2 = st.columns([3, 1])
            with col1:
                new_doc_name = st.text_input("Document Name", placeholder="e.g. NPS Test Results")
                new_doc_desc = st.text_input("Description (optional)", placeholder="Brief description of what this document is")
            with col2:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("➕ Add Requirement", use_container_width=True):
                    if not new_doc_name:
                        st.error("Please enter a document name.")
                    else:
                        supabase.table("et_required_docs").insert({
                            "doc_name": new_doc_name,
                            "description": new_doc_desc or "",
                            "added_by": name,
                        }).execute()
                        st.success(f"✅ '{new_doc_name}' added to requirements!")
                        st.rerun()
            if req_docs.data:
                st.divider()
                st.markdown("**Remove a requirement:**")
                del_doc_map = {f"{r['doc_name']}": r['id'] for r in req_docs.data}
                col1, col2 = st.columns([3, 1])
                with col1:
                    selected_req_del = st.selectbox("Select requirement to remove", list(del_doc_map.keys()), key="del_req_doc")
                with col2:
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("🗑️ Remove", use_container_width=True, key="btn_del_req_doc"):
                        supabase.table("et_required_docs").delete().eq("id", del_doc_map[selected_req_del]).execute()
                        st.success(f"Removed '{selected_req_del}' from requirements.")
                        st.rerun()

    st.divider()
    if not req_docs.data:
        st.info("No required documents defined yet. Add some using the button above.")
    else:
        for req in req_docs.data:
            doc_name = req['doc_name']
            doc_desc = req.get('description', '')
            uploads = [u for u in (all_uploads.data or []) if u['doc_type'] == doc_name]
            has_uploads = len(uploads) > 0
            status_icon = "✅" if has_uploads else "❌"
            status_color = "#2ea043" if has_uploads else "#f85149"
            with st.container():
                st.markdown(f"""
                    <div style="border: 1px solid {'#2ea04355' if has_uploads else '#f8514955'};
                        border-radius: 10px; padding: 16px 20px; margin-bottom: 14px;
                        background: {'#2ea04308' if has_uploads else '#f8514908'};">
                        <div style="display:flex; align-items:center; gap:10px; margin-bottom:4px;">
                            <span style="font-size:18px">{status_icon}</span>
                            <span style="font-weight:700; font-size:16px;">{doc_name}</span>
                            <span style="font-size:11px; color:{status_color}; margin-left:auto;">
                                {len(uploads)} file{'s' if len(uploads) != 1 else ''} uploaded
                            </span>
                        </div>
                        {f'<div style="font-size:12px; color:#8B949E; margin-bottom:8px;">{doc_desc}</div>' if doc_desc else ''}
                    </div>
                """, unsafe_allow_html=True)
                if has_uploads:
                    for upload in uploads:
                        file_name = upload['file_name']
                        file_path = upload['file_path']
                        uploaded_by = upload['uploaded_by']
                        uploaded_at = pd.to_datetime(upload['created_at']).strftime("%d %b %Y %H:%M")
                        ext = file_name.split('.')[-1].lower()
                        file_url = f"https://sjcwzbftzpfylwdqiknh.supabase.co/storage/v1/object/public/asset/{quote(file_path, safe='/')}"
                        st.caption(f"Debug URL: {file_url}")
                        col1, col2, col3 = st.columns([4, 2, 1])
                        with col1:
                            if ext in ['png', 'jpg', 'jpeg']:
                                st.image(file_url, width=200)
                            else:
                                st.markdown(f"📄 [{file_name}]({file_url})")
                            st.caption(f"Uploaded by {uploaded_by} · {uploaded_at}")
                        with col3:
                            if can_edit:
                                if st.button("🗑️", key=f"del_file_{upload['id']}", help="Delete this file"):
                                    supabase.table("et_documents").delete().eq("id", upload['id']).execute()
                                    st.success("Deleted!")
                                    st.rerun()
                if can_edit:
                    with st.expander(f"⬆️ Upload file for '{doc_name}'", expanded=False):
                        doc_file = st.file_uploader("Choose file", type=["xlsx", "pdf", "png", "jpg", "docx"], key=f"upload_{req['id']}")
                        if doc_file and st.button("Upload", key=f"btn_upload_{req['id']}", type="primary"):
                            with st.spinner("Uploading..."):
                                try:
                                    file_bytes = doc_file.read()
                                    file_path  = f"ET/{datetime.now().strftime('%Y%m%d_%H%M%S')}_{doc_file.name}"
                                    supabase.storage.from_("asset").upload(file_path, file_bytes)
                                    supabase.table("et_documents").insert({
                                        "doc_type": doc_name,
                                        "file_name": doc_file.name,
                                        "file_path": file_path,
                                        "uploaded_by": name,
                                    }).execute()
                                    st.success(f"✅ Uploaded successfully!")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Upload failed: {str(e)}")
                st.markdown("---")

# ════════════════════════════════════════
# TAB 4 — REQUEST TRAINING
# ════════════════════════════════════════
with tab4:
    st.markdown("### 📋 Training Requests")
    st.caption("Any pillar leader or coordinator can submit a training request. E&T leader manages the status.")
    try:
        requests = supabase.table("et_training_requests").select("*").order("created_at", desc=True).execute()
        if requests.data:
            req_df = pd.DataFrame(requests.data)
            req_df["created_at"] = pd.to_datetime(req_df["created_at"]).dt.strftime("%d %b %Y")
            st.dataframe(
                req_df[["created_at", "requested_by", "employee_name", "topic", "reason", "urgency", "status"]].rename(columns={
                    "created_at": "Date", "requested_by": "Requested By",
                    "employee_name": "Employee", "topic": "Topic",
                    "reason": "Reason", "urgency": "Urgency", "status": "Status"
                }),
                use_container_width=True
            )
            if can_edit:
                st.markdown("#### ✏️ Update Request Status")
                req_map = {f"{r['employee_name']} — {r['topic']} ({r['created_at']})": r['id'] for r in requests.data}
                col1, col2 = st.columns([3, 1])
                with col1:
                    selected_req = st.selectbox("Select request", list(req_map.keys()))
                with col2:
                    new_status = st.selectbox("Status", ["Pending", "Scheduled", "Done"])
                if st.button("✅ Update Status"):
                    supabase.table("et_training_requests").update({"status": new_status}).eq("id", req_map[selected_req]).execute()
                    st.success("Status updated!")
                    st.rerun()
                st.markdown("#### 🗑️ Delete Request")
                del_req_map = {f"{r['employee_name']} — {r['topic']} ({r['created_at']})": r['id'] for r in requests.data}
                selected_del_req = st.selectbox("Select request to delete", list(del_req_map.keys()), key="del_req")
                if st.button("🗑️ Delete Request", key="btn_del_req"):
                    supabase.table("et_training_requests").delete().eq("id", del_req_map[selected_del_req]).execute()
                    st.success("Request deleted!")
                    st.rerun()
        else:
            st.info("No training requests yet.")
    except Exception as e:
        st.error(f"Could not load requests: {str(e)}")

    st.divider()
    st.markdown("#### ➕ Submit Training Request")
    emp_name_req = st.text_input("Employee Name")
    topic_req    = st.text_input("Training Topic")
    reason_req   = st.text_area("Reason / Justification", placeholder="Why does this employee need this training?")
    urgency_req  = st.selectbox("Urgency", ["Low", "Medium", "High"])
    if st.button("📨 Submit Request", type="primary"):
        if not emp_name_req or not topic_req:
            st.error("Please fill in employee name and topic.")
        else:
            try:
                supabase.table("et_training_requests").insert({
                    "requested_by": name,
                    "employee_name": emp_name_req,
                    "topic": topic_req,
                    "reason": reason_req,
                    "urgency": urgency_req,
                    "status": "Pending",
                }).execute()
                st.success("✅ Training request submitted!")
                st.rerun()
            except Exception as e:
                st.error(f"Error: {str(e)}")

# ════════════════════════════════════════
# TAB 5 — ACTION PLANS
# ════════════════════════════════════════
with tab5:
    st.markdown("### 🎯 Action Plans")
    try:
        actions = supabase.table("et_action_plans").select("*").order("created_at", desc=True).execute()
        if actions.data:
            act_df = pd.DataFrame(actions.data)
            act_df["due_date"] = pd.to_datetime(act_df["due_date"]).dt.strftime("%d %b %Y")
            st.dataframe(
                act_df[["action", "owner", "due_date", "status", "notes"]].rename(columns={
                    "action": "Action", "owner": "Owner",
                    "due_date": "Due Date", "status": "Status", "notes": "Notes"
                }),
                use_container_width=True
            )
            if can_edit:
                st.markdown("#### ✏️ Update Action Status")
                act_map = {f"{r['action'][:40]} — {r['owner']}": r['id'] for r in actions.data}
                col1, col2 = st.columns([3, 1])
                with col1:
                    selected_act = st.selectbox("Select action", list(act_map.keys()))
                with col2:
                    new_act_status = st.selectbox("Status", ["Open", "In Progress", "Done"], key="update_act_status")
                if st.button("✅ Update Action Status"):
                    supabase.table("et_action_plans").update({"status": new_act_status}).eq("id", act_map[selected_act]).execute()
                    st.success("Status updated!")
                    st.rerun()
                st.markdown("#### 🗑️ Delete Action Plan")
                del_act_map = {f"{r['action'][:40]} — {r['owner']}": r['id'] for r in actions.data}
                selected_del_act = st.selectbox("Select action to delete", list(del_act_map.keys()), key="del_act")
                if st.button("🗑️ Delete Action Plan", key="btn_del_act"):
                    supabase.table("et_action_plans").delete().eq("id", del_act_map[selected_del_act]).execute()
                    st.success("Action plan deleted!")
                    st.rerun()
        else:
            st.info("No action plans yet.")
    except Exception as e:
        st.error(f"Could not load action plans: {str(e)}")

    if can_edit:
        st.divider()
        st.markdown("#### ➕ Add Action Plan")
        action_text = st.text_input("Action")
        try:
            ppl_res = supabase.table("profiles").select("full_name").order("full_name").execute()
            ppl_names = [p["full_name"] for p in (ppl_res.data or []) if p.get("full_name")]
        except Exception:
            ppl_names = []
        if ppl_names:
            own_sel = st.selectbox("Owner", ["— select —"] + ppl_names, key="et_own_sel")
            owner_text = own_sel if own_sel != "— select —" else ""
        else:
            owner_text = st.text_input("Owner")
        due_date    = st.date_input("Due Date")
        status_ap   = st.selectbox("Status", ["Open", "In Progress", "Done"], key="new_act_status")
        notes_ap    = st.text_area("Notes", placeholder="Additional context...")
        if st.button("➕ Add Action Plan", type="primary"):
            if not action_text or not owner_text:
                st.error("Please fill in action and owner.")
            else:
                try:
                    supabase.table("et_action_plans").insert({
                        "action": action_text,
                        "owner": owner_text,
                        "due_date": str(due_date),
                        "status": status_ap,
                        "notes": notes_ap,
                    }).execute()
                    try:
                        owner_res = supabase.table("profiles").select("id").eq("full_name", owner_text).execute()
                        if owner_res.data:
                            create_notification(
                                supabase, owner_res.data[0]["id"],
                                title="📋 New Action Plan Assigned — ET",
                                message=f'"{action_text[:60]}" assigned to you. Due: {due_date.strftime("%d %b %Y")}.',
                                notif_type="info"
                            )
                    except Exception:
                        pass
                    st.success("✅ Action plan added!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {str(e)}")

# ════════════════════════════════════════
# TAB 6 — ONE POINT LESSONS (OPL)
# ════════════════════════════════════════
with tab6:
    import io as _opl_io
    import base64 as _opl_b64
    import json as _opl_json
    from reportlab.pdfgen import canvas as _opl_canvas
    from reportlab.lib.utils import ImageReader as _opl_ir
    from reportlab.lib.colors import HexColor as _opl_hc, black, white, red, green

    st.markdown("### 📝 One Point Lessons")

    OPL_TYPES = [
        "Equipment Handling", "Safety", "Quality", "Maintenance",
        "Environment", "Process", "Cleaning & Inspection", "Other"
    ]

    # ── Annotation canvas component ──
    def annotation_canvas(image_bytes, key):
        """Returns annotated image bytes or None"""
        if not image_bytes:
            return None
        img_b64 = _opl_b64.b64encode(image_bytes).decode()
        canvas_html = f"""
        <div style="position:relative;display:inline-block;">
          <canvas id="canvas_{key}" style="border:2px solid #006394;cursor:crosshair;max-width:100%;"></canvas>
          <br/>
          <div style="margin:6px 0;">
            <button onclick="setTool('circle')" style="margin:2px;padding:4px 10px;background:#006394;color:white;border:none;border-radius:4px;cursor:pointer;">⭕ Circle</button>
            <button onclick="setTool('arrow')" style="margin:2px;padding:4px 10px;background:#C1A02E;color:white;border:none;border-radius:4px;cursor:pointer;">➡️ Arrow</button>
            <button onclick="setTool('rect')" style="margin:2px;padding:4px 10px;background:#DE201B;color:white;border:none;border-radius:4px;cursor:pointer;">🟥 Rectangle</button>
            <button onclick="undoLast()" style="margin:2px;padding:4px 10px;background:#888;color:white;border:none;border-radius:4px;cursor:pointer;">↩ Undo</button>
            <button onclick="clearAll()" style="margin:2px;padding:4px 10px;background:#333;color:white;border:none;border-radius:4px;cursor:pointer;">🗑 Clear</button>
            <button onclick="saveCanvas()" style="margin:2px;padding:4px 10px;background:#27AE60;color:white;border:none;border-radius:4px;cursor:pointer;">💾 Save Annotation</button>
          </div>
          <input type="hidden" id="result_{key}" value=""/>
        </div>
        <script>
        (function() {{
          const canvas = document.getElementById('canvas_{key}');
          const ctx = canvas.getContext('2d');
          const img = new Image();
          img.onload = function() {{
            canvas.width = img.width;
            canvas.height = img.height;
            ctx.drawImage(img, 0, 0);
          }};
          img.src = 'data:image/jpeg;base64,{img_b64}';

          let tool = 'circle';
          let drawing = false;
          let startX, startY;
          let shapes = [];
          let snapshot;

          function setTool(t) {{ tool = t; }}
          window.setTool = setTool;

          canvas.addEventListener('mousedown', e => {{
            const r = canvas.getBoundingClientRect();
            const scaleX = canvas.width / r.width;
            const scaleY = canvas.height / r.height;
            startX = (e.clientX - r.left) * scaleX;
            startY = (e.clientY - r.top) * scaleY;
            drawing = true;
            snapshot = ctx.getImageData(0, 0, canvas.width, canvas.height);
          }});

          canvas.addEventListener('mousemove', e => {{
            if (!drawing) return;
            const r = canvas.getBoundingClientRect();
            const scaleX = canvas.width / r.width;
            const scaleY = canvas.height / r.height;
            const x = (e.clientX - r.left) * scaleX;
            const y = (e.clientY - r.top) * scaleY;
            ctx.putImageData(snapshot, 0, 0);
            ctx.strokeStyle = '#DE201B';
            ctx.lineWidth = 3;
            ctx.beginPath();
            if (tool === 'circle') {{
              const rx = Math.abs(x - startX) / 2;
              const ry = Math.abs(y - startY) / 2;
              const cx = startX + (x - startX) / 2;
              const cy = startY + (y - startY) / 2;
              ctx.ellipse(cx, cy, rx, ry, 0, 0, 2 * Math.PI);
              ctx.stroke();
            }} else if (tool === 'rect') {{
              ctx.strokeRect(startX, startY, x - startX, y - startY);
            }} else if (tool === 'arrow') {{
              ctx.moveTo(startX, startY);
              ctx.lineTo(x, y);
              ctx.stroke();
              const angle = Math.atan2(y - startY, x - startX);
              const len = 15;
              ctx.beginPath();
              ctx.moveTo(x, y);
              ctx.lineTo(x - len * Math.cos(angle - 0.4), y - len * Math.sin(angle - 0.4));
              ctx.lineTo(x - len * Math.cos(angle + 0.4), y - len * Math.sin(angle + 0.4));
              ctx.closePath();
              ctx.fillStyle = '#DE201B';
              ctx.fill();
            }}
          }});

          canvas.addEventListener('mouseup', e => {{
            if (!drawing) return;
            drawing = false;
            const r = canvas.getBoundingClientRect();
            const scaleX = canvas.width / r.width;
            const scaleY = canvas.height / r.height;
            const x = (e.clientX - r.left) * scaleX;
            const y = (e.clientY - r.top) * scaleY;
            shapes.push({{tool, startX, startY, x, y}});
          }});

          window.undoLast = function() {{
            shapes.pop();
            redraw();
          }};

          window.clearAll = function() {{
            shapes = [];
            redraw();
          }};

          function redraw() {{
            ctx.putImageData(snapshot, 0, 0);
            shapes.forEach(s => {{
              ctx.strokeStyle = '#DE201B';
              ctx.lineWidth = 3;
              ctx.beginPath();
              if (s.tool === 'circle') {{
                const rx = Math.abs(s.x - s.startX) / 2;
                const ry = Math.abs(s.y - s.startY) / 2;
                const cx = s.startX + (s.x - s.startX) / 2;
                const cy = s.startY + (s.y - s.startY) / 2;
                ctx.ellipse(cx, cy, rx, ry, 0, 0, 2 * Math.PI);
                ctx.stroke();
              }} else if (s.tool === 'rect') {{
                ctx.strokeRect(s.startX, s.startY, s.x - s.startX, s.y - s.startY);
              }} else if (s.tool === 'arrow') {{
                ctx.moveTo(s.startX, s.startY);
                ctx.lineTo(s.x, s.y);
                ctx.stroke();
                const angle = Math.atan2(s.y - s.startY, s.x - s.startX);
                const len = 15;
                ctx.beginPath();
                ctx.moveTo(s.x, s.y);
                ctx.lineTo(s.x - len * Math.cos(angle - 0.4), s.y - len * Math.sin(angle - 0.4));
                ctx.lineTo(s.x - len * Math.cos(angle + 0.4), s.y - len * Math.sin(angle + 0.4));
                ctx.closePath();
                ctx.fillStyle = '#DE201B';
                ctx.fill();
              }}
            }});
          }}

          window.saveCanvas = function() {{
            const data = canvas.toDataURL('image/jpeg', 0.92).split(',')[1];
            document.getElementById('result_{key}').value = data;
            alert('✅ Annotation saved!');
          }};
        }})();
        </script>
        """
        st.components.v1.html(canvas_html, height=500, scrolling=True)
        return None  # annotation saved via JS — user clicks Save then we read from session

    # ── PDF Generator ──
    def generate_opl_pdf(opl):
        W, H = 792, 612  # landscape A4 in points
        buf = _opl_io.BytesIO()
        c = _opl_canvas.Canvas(buf, pagesize=(W, H))

        # ── Header ──
        # Company logo box
        c.setFillColor(_opl_hc("#DDDDDD"))
        c.rect(20, H-60, 120, 50, fill=1, stroke=1)
        c.setFillColor(black)
        c.setFont("Helvetica-Bold", 8)
        c.drawCentredString(80, H-40, "Company Logo")

        # Title
        c.setFillColor(_opl_hc("#DDDDDD"))
        c.rect(140, H-60, 512, 50, fill=1, stroke=1)
        c.setFillColor(black)
        c.setFont("Helvetica-Bold", 22)
        c.drawCentredString(396, H-32, "ONE POINT LESSON")

        # OPL label box
        c.setFillColor(_opl_hc("#1F4E79"))
        c.rect(20, H-95, 120, 35, fill=1, stroke=0)
        c.setFillColor(white)
        c.setFont("Helvetica-Bold", 16)
        c.drawCentredString(80, H-78, "OPL")

        # Subject / Type row
        c.setFillColor(white)
        c.rect(140, H-95, 512, 18, fill=0, stroke=1)
        c.setFillColor(black)
        c.setFont("Helvetica", 7)
        c.drawString(145, H-84, "Subject:")
        c.drawString(530, H-84, "Type")

        c.rect(140, H-113, 512, 18, fill=0, stroke=1)
        c.setFont("Helvetica-Bold", 9)
        c.drawString(145, H-106, opl.get("subject",""))
        c.drawString(530, H-106, opl.get("opl_type",""))

        # ── Body divider ──
        body_top = H - 115
        body_bot = 120
        body_h = body_top - body_bot
        mid_x = W / 2

        # Bad side header (❌)
        c.setFillColor(_opl_hc("#FFCCCC"))
        c.rect(20, body_top - 22, mid_x - 25, 22, fill=1, stroke=1)
        c.setFillColor(red)
        c.setFont("Helvetica-Bold", 18)
        c.drawCentredString((20 + mid_x - 25) / 2, body_top - 15, "✗")

        # Good side header (✓)
        c.setFillColor(_opl_hc("#CCFFCC"))
        c.rect(mid_x + 5, body_top - 22, mid_x - 25, 22, fill=1, stroke=1)
        c.setFillColor(green)
        c.setFont("Helvetica-Bold", 18)
        c.drawCentredString(mid_x + 5 + (mid_x - 25) / 2, body_top - 15, "✓")

        # Body frames
        c.setStrokeColor(black)
        c.rect(20, body_bot, mid_x - 45, body_h - 22, fill=0, stroke=1)
        c.rect(mid_x + 5, body_bot, mid_x - 25, body_h - 22, fill=0, stroke=1)

        def draw_side(x_start, w, texts, image_paths):
            """Draw bad or good side content"""
            y_cursor = body_top - 30
            slot_h = (body_h - 30) / 2
            for i, (txt, img_path) in enumerate(zip(texts, image_paths)):
                slot_top = body_top - 30 - i * slot_h
                slot_bot = slot_top - slot_h + 10
                # Text on left ~30% of slot
                if txt:
                    c.setFillColor(black)
                    c.setFont("Helvetica", 8)
                    txt_x = x_start + 4
                    txt_w = w * 0.28
                    words = txt.split()
                    lines = []
                    line = ""
                    for word in words:
                        test = (line + " " + word).strip()
                        if c.stringWidth(test, "Helvetica", 8) < txt_w:
                            line = test
                        else:
                            if line: lines.append(line)
                            line = word
                    if line: lines.append(line)
                    ty = slot_top - 12
                    for ln in lines[:8]:
                        c.drawString(txt_x, ty, ln)
                        ty -= 10
                # Image on right ~68%
                if img_path:
                    try:
                        img_x = x_start + w * 0.30
                        img_w = w * 0.66
                        img_h = slot_h - 16
                        c.drawImage(_opl_ir(img_path), img_x, slot_bot + 4,
                                    width=img_w, height=img_h,
                                    preserveAspectRatio=True, mask="auto")
                    except Exception:
                        pass

        # Prepare image temp files
        import tempfile as _tmp
        def _save_img(b64_str):
            if not b64_str: return None
            try:
                data = _opl_b64.b64decode(b64_str)
                tf = _tmp.NamedTemporaryFile(delete=False, suffix=".jpg")
                tf.write(data); tf.flush()
                return tf.name
            except Exception:
                return None

        bad_imgs = [_save_img(opl.get("bad_image_1")), _save_img(opl.get("bad_image_2"))]
        good_imgs = [_save_img(opl.get("good_image_1")), _save_img(opl.get("good_image_2"))]
        bad_texts = [opl.get("bad_text_1",""), opl.get("bad_text_2","")]
        good_texts = [opl.get("good_text_1",""), opl.get("good_text_2","")]

        draw_side(20, mid_x - 45, bad_texts, bad_imgs)
        draw_side(mid_x + 5, mid_x - 25, good_texts, good_imgs)

        # ── Footer ──
        fy = 118
        c.setFont("Helvetica-Bold", 7)
        c.setFillColor(black)
        col_w = (W - 40) / 3
        for i, (label, val_name, date_name) in enumerate([
            ("Prepared By", "prepared_by", "prepared_date"),
            ("Approved by", "approved_by", "approved_date"),
            ("Administered by", "administered_by", "administered_date"),
        ]):
            fx = 20 + i * col_w
            c.rect(fx, fy - 12, col_w, 12, fill=0, stroke=1)
            c.drawString(fx + 3, fy - 9, label)
            c.rect(fx, fy - 24, col_w, 12, fill=0, stroke=1)
            c.setFont("Helvetica", 7)
            c.drawString(fx + 3, fy - 21, f"Date: {opl.get(date_name,'')}")
            c.setFont("Helvetica-Bold", 7)

        # Operators row
        c.setFont("Helvetica-Bold", 7)
        c.rect(20, fy - 38, 80, 14, fill=0, stroke=1)
        c.drawString(23, fy - 30, "Seen & understood by:")
        ops = opl.get("operators","").split(",")
        op_w = (W - 100) / max(len(ops), 4)
        for j, op in enumerate(ops[:6]):
            ox = 100 + j * op_w
            c.rect(ox, fy - 38, op_w, 14, fill=0, stroke=1)
            c.setFont("Helvetica", 6)
            c.drawString(ox + 2, fy - 28, "Operator")
            c.drawString(ox + 2, fy - 36, op.strip()[:20])

        c.save()
        buf.seek(0)
        return buf

    # ── UI ──
    opl_view, opl_create = st.tabs(["📋 All OPLs", "➕ Create OPL"])

    # ── LIST VIEW ──
    with opl_view:
        try:
            opls = supabase.table("et_opls").select("*").order("created_at", desc=True).execute()
            opl_list = opls.data or []
        except Exception as e:
            st.error(f"Could not load OPLs: {e}"); opl_list = []

        if not opl_list:
            st.info("No OPLs created yet.")
        else:
            for opl in opl_list:
                with st.expander(f"📝 {opl['subject']} — {opl['opl_type']} | {opl.get('status','Draft')}", expanded=False):
                    c1, c2, c3 = st.columns([3,1,1])
                    c1.caption(f"Created by {opl.get('created_by','')} on {str(opl.get('created_at',''))[:10]}")

                    # Download PDF
                    try:
                        pdf_buf = generate_opl_pdf(opl)
                        c2.download_button("📥 PDF", data=pdf_buf,
                            file_name=f"OPL_{opl['subject'].replace(' ','_')}.pdf",
                            mime="application/pdf", key=f"opl_dl_{opl['id']}")
                    except Exception as e:
                        c2.warning(f"PDF error: {e}")

                    # Delete
                    if can_edit:
                        if c3.button("🗑️ Delete", key=f"opl_del_{opl['id']}"):
                            supabase.table("et_opls").delete().eq("id", opl["id"]).execute()
                            st.rerun()

                    # Preview
                    st.markdown(f"**Subject:** {opl['subject']} | **Type:** {opl['opl_type']}")
                    col_b, col_g = st.columns(2)
                    with col_b:
                        st.markdown("❌ **Bad**")
                        if opl.get("bad_text_1"): st.caption(opl["bad_text_1"])
                        if opl.get("bad_image_1"):
                            try: st.image(_opl_b64.b64decode(opl["bad_image_1"]), use_container_width=True)
                            except: pass
                        if opl.get("bad_text_2"): st.caption(opl["bad_text_2"])
                        if opl.get("bad_image_2"):
                            try: st.image(_opl_b64.b64decode(opl["bad_image_2"]), use_container_width=True)
                            except: pass
                    with col_g:
                        st.markdown("✅ **Good**")
                        if opl.get("good_text_1"): st.caption(opl["good_text_1"])
                        if opl.get("good_image_1"):
                            try: st.image(_opl_b64.b64decode(opl["good_image_1"]), use_container_width=True)
                            except: pass
                        if opl.get("good_text_2"): st.caption(opl["good_text_2"])
                        if opl.get("good_image_2"):
                            try: st.image(_opl_b64.b64decode(opl["good_image_2"]), use_container_width=True)
                            except: pass

    # ── CREATE / EDIT VIEW ──
    with opl_create:
        if not can_edit:
            st.warning("🔒 Only Pillar Leader or Plant Manager can create OPLs.")
        else:
            st.markdown("#### ➕ New One Point Lesson")

            with st.form("opl_create_form"):
                fc1, fc2 = st.columns(2)
                opl_subject = fc1.text_input("Subject *", placeholder="e.g. Motor Shafts")
                opl_type    = fc2.selectbox("Type *", OPL_TYPES)

                st.divider()
                st.markdown("### ❌ Bad Side")
                b1c1, b1c2 = st.columns([1,2])
                bad_text_1  = b1c1.text_area("Description 1", key="bad_t1", height=100)
                bad_file_1  = b1c2.file_uploader("Image 1 (Bad)", type=["jpg","jpeg","png"], key="bad_f1")
                b2c1, b2c2  = st.columns([1,2])
                bad_text_2  = b2c1.text_area("Description 2", key="bad_t2", height=100)
                bad_file_2  = b2c2.file_uploader("Image 2 (Bad)", type=["jpg","jpeg","png"], key="bad_f2")

                st.divider()
                st.markdown("### ✅ Good Side")
                g1c1, g1c2 = st.columns([1,2])
                good_text_1 = g1c1.text_area("Description 1", key="good_t1", height=100)
                good_file_1 = g1c2.file_uploader("Image 1 (Good)", type=["jpg","jpeg","png"], key="good_f1")
                g2c1, g2c2  = st.columns([1,2])
                good_text_2 = g2c1.text_area("Description 2", key="good_t2", height=100)
                good_file_2 = g2c2.file_uploader("Image 2 (Good)", type=["jpg","jpeg","png"], key="good_f2")

                st.divider()
                st.markdown("### 📋 Footer")
                ff1, ff2, ff3 = st.columns(3)
                prepared_by     = ff1.text_input("Prepared By")
                prepared_date   = ff1.text_input("Date", key="prep_date")
                approved_by     = ff2.text_input("Approved By")
                approved_date   = ff2.text_input("Date", key="appr_date")
                administered_by = ff3.text_input("Administered By")
                administered_date = ff3.text_input("Date", key="adm_date")
                operators = st.text_input("Operators (comma separated)", placeholder="John, Sarah, Mike")

                submitted = st.form_submit_button("💾 Save OPL", type="primary")
                if submitted:
                    if not opl_subject:
                        st.error("Please enter a subject.")
                    else:
                        def _to_b64(f):
                            if f is None: return None
                            return _opl_b64.b64encode(f.getvalue()).decode()
                        try:
                            supabase.table("et_opls").insert({
                                "subject":          opl_subject,
                                "opl_type":         opl_type,
                                "bad_text_1":       bad_text_1,
                                "bad_image_1":      _to_b64(bad_file_1),
                                "bad_text_2":       bad_text_2,
                                "bad_image_2":      _to_b64(bad_file_2),
                                "good_text_1":      good_text_1,
                                "good_image_1":     _to_b64(good_file_1),
                                "good_text_2":      good_text_2,
                                "good_image_2":     _to_b64(good_file_2),
                                "prepared_by":      prepared_by,
                                "prepared_date":    prepared_date,
                                "approved_by":      approved_by,
                                "approved_date":    approved_date,
                                "administered_by":  administered_by,
                                "administered_date":administered_date,
                                "operators":        operators,
                                "created_by":       name,
                                "status":           "Draft",
                            }).execute()
                            st.success("✅ OPL saved!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error saving OPL: {e}")

            # ── Annotation tool (outside form) ──
            st.divider()
            st.markdown("#### 🎨 Annotate Images")
            st.caption("Upload an image below, draw circles/arrows, then click Save Annotation. Copy the annotated image and re-upload it in the form above.")

            ann_file = st.file_uploader("Upload image to annotate", type=["jpg","jpeg","png"], key="ann_upload")
            if ann_file:
                annotation_canvas(ann_file.getvalue(), key="ann_main")
