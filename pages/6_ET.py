# pages/6_ET.py - Education & Training Pillar
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import io
import base64
from fpdf import FPDF
from utils.supabase_client import get_supabase

st.set_page_config(page_title="E&T Pillar", page_icon="🎓", layout="wide")

# ── Auth guard ──
if "user" not in st.session_state:
    st.switch_page("app.py")

supabase = get_supabase()
role     = st.session_state.get("role", "member")
pillar   = st.session_state.get("pillar", "ALL")
name     = st.session_state.get("name", "User")
user_id  = st.session_state.get("user", "")

can_edit = (role == "plant_manager") or (role == "pillar_leader" and pillar == "ET")

# ── Header ──
col1, col2 = st.columns([5, 1])
with col1:
    st.markdown("# 🎓 Education & Training")
    st.markdown(f"Logged in as **{name}** · `{role}` · {'✏️ Full Access' if can_edit else '👁️ View Only'}")
with col2:
    if st.button("🏠 Home", use_container_width=True):
        st.switch_page("pages/1_Home.py")

st.divider()

# ── Tabs ──
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Maturity Level",
    "🔬 Competency Analysis",
    "📁 Documents",
    "📋 Request Training",
    "🎯 Action Plans"
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
            "name": "Entry", "color": "#888888", "icon": "⚪",
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
            "name": "Bronze", "color": "#CD7F32", "icon": "🥉",
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
            "name": "Silver", "color": "#C0C0C0", "icon": "🥈",
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
            "name": "Gold", "color": "#FFD700", "icon": "🥇",
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
                st.checkbox(criterion, value=is_current, key=f"lvl_{level['name']}_{criterion[:20]}")

# ════════════════════════════════════════
# TAB 2 — COMPETENCY ANALYSIS
# ════════════════════════════════════════
with tab2:
    st.markdown("### Competency Matrix Analysis")

    # ── Past runs ──
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
        else:
            st.info("No analysis runs yet. Upload a file below to run the first analysis.")
    except:
        st.info("Analysis history table not set up yet. Run the SQL setup first.")

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

                    # Read requirements
                    requirement_df = pd.read_excel(xls, sheet_name='Requirement', dtype=str)
                    requirement_df.columns = requirement_df.columns.str.strip()
                    requirement_df['Grade'] = requirement_df['Grade'].astype(str).str.strip()
                    topics = list(requirement_df.columns[1:])
                    requirement_dict = requirement_df.set_index('Grade')[topics].to_dict(orient='index')

                    # Read employee list
                    employee_list_df = pd.read_excel(xls, sheet_name='Employee List', dtype=str)
                    employee_list_df.columns = employee_list_df.columns.str.strip()
                    employee_list_df['Grade'] = employee_list_df['Grade'].astype(str).str.strip()

                    # Department sheets
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

                        # Charts
                        topic_counts = failures_df['Topic'].value_counts()
                        section_counts = failures_df['Section'].value_counts()
                        topic_grade_counts = failures_df.groupby(['Topic', 'Grade']).size().reset_index(name='Count')
                        heatmap_data = failures_df.groupby(['Grade', 'Topic']).size().unstack(fill_value=0)

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

                        # Display charts
                        st.markdown("#### 📊 Results")
                        st.dataframe(failures_df, use_container_width=True)
                        c1, c2 = st.columns(2)
                        with c1:
                            st.pyplot(fig1)
                            st.pyplot(fig3)
                        with c2:
                            st.pyplot(fig2)
                            st.pyplot(fig4)

                        # Save charts to bytes for PDF
                        def fig_to_bytes(fig):
                            buf = io.BytesIO()
                            fig.savefig(buf, format='png', dpi=100)
                            buf.seek(0)
                            return buf

                        buf1 = fig_to_bytes(fig1)
                        buf2 = fig_to_bytes(fig2)
                        buf3 = fig_to_bytes(fig3)
                        buf4 = fig_to_bytes(fig4)

                        # Generate PDF
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

                        # Save chart images temporarily and add to PDF
                        import tempfile, os
                        tmp_files = []
                        for buf, label in [(buf4, "Failure Matrix"), (buf1, "Failures per Topic"), (buf2, "Failures by Grade"), (buf3, "Failures by Section")]:
                            pdf.add_page()
                            pdf.set_font("Arial", "B", 14)
                            pdf.set_text_color(13, 104, 163)
                            pdf.cell(0, 10, label, ln=True)
                            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
                            tmp.write(buf.read())
                            tmp.close()
                            tmp_files.append(tmp.name)
                            pdf.image(tmp.name, w=180)

                        pdf_bytes = pdf.output(dest='S').encode('latin-1')
                        for f in tmp_files:
                            os.unlink(f)

                        st.download_button(
                            label="📥 Download PDF Report",
                            data=pdf_bytes,
                            file_name=f"Training_Needs_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                            mime="application/pdf"
                        )

                        # Save run to Supabase
                        supabase.table("et_analysis_runs").insert({
                            "run_by": name,
                            "total_failures": total_failures,
                            "notes": notes or "",
                        }).execute()

                    else:
                        st.success("🎉 No failures detected — all employees meet requirements!")

                except Exception as e:
                    st.error(f"Error during analysis: {str(e)}")

# ════════════════════════════════════════
# TAB 3 — DOCUMENTS
# ════════════════════════════════════════
with tab3:
    st.markdown("### 📁 E&T Required Documents")

    doc_types = [
        "Flexibility Chart",
        "Training Schedule",
        "Training Records",
        "Employee Improvement Log",
        "OPL Documents",
    ]

    # Show existing uploads
    try:
        docs = supabase.table("et_documents").select("*").order("created_at", desc=True).execute()
        if docs.data:
            docs_df = pd.DataFrame(docs.data)
            docs_df["created_at"] = pd.to_datetime(docs_df["created_at"]).dt.strftime("%d %b %Y")
            st.dataframe(
                docs_df[["doc_type", "file_name", "uploaded_by", "created_at"]].rename(columns={
                    "doc_type": "Document Type", "file_name": "File Name",
                    "uploaded_by": "Uploaded By", "created_at": "Date"
                }),
                use_container_width=True
            )
        else:
            st.info("No documents uploaded yet.")
    except:
        st.info("Documents table not set up yet.")

    st.divider()

    if not can_edit:
        st.warning("👁️ View Only — you do not have permission to upload documents.")
    else:
        st.markdown("#### ⬆️ Upload New Document")
        doc_type = st.selectbox("Document Type", doc_types)
        doc_file = st.file_uploader("Choose file", type=["xlsx", "pdf", "png", "jpg", "docx"])

        if doc_file and st.button("Upload Document", type="primary"):
            with st.spinner("Uploading..."):
                try:
                    file_bytes = doc_file.read()
                    file_path  = f"ET/{datetime.now().strftime('%Y%m%d_%H%M%S')}_{doc_file.name}"
                    supabase.storage.from_("assets").upload(file_path, file_bytes)
                    supabase.table("et_documents").insert({
                        "doc_type": doc_type,
                        "file_name": doc_file.name,
                        "file_path": file_path,
                        "uploaded_by": name,
                    }).execute()
                    st.success(f"✅ {doc_file.name} uploaded successfully!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Upload failed: {str(e)}")

# ════════════════════════════════════════
# TAB 4 — REQUEST TRAINING
# ════════════════════════════════════════
with tab4:
    st.markdown("### 📋 Training Requests")

    # Show all requests
    try:
        requests = supabase.table("et_training_requests").select("*").order("created_at", desc=True).execute()
        if requests.data:
            req_df = pd.DataFrame(requests.data)
            req_df["created_at"] = pd.to_datetime(req_df["created_at"]).dt.strftime("%d %b %Y")

            # Color status
            def status_color(s):
                if s == "Done": return "background-color: #d4edda"
                if s == "Scheduled": return "background-color: #fff3cd"
                return "background-color: #f8d7da"

            st.dataframe(
                req_df[["created_at", "requested_by", "employee_name", "topic", "reason", "urgency", "status"]].rename(columns={
                    "created_at": "Date", "requested_by": "Requested By",
                    "employee_name": "Employee", "topic": "Topic",
                    "reason": "Reason", "urgency": "Urgency", "status": "Status"
                }),
                use_container_width=True
            )

            # ET pillar leader can update status
            if can_edit:
                st.markdown("#### Update Request Status")
                req_ids = {f"{r['employee_name']} — {r['topic']} ({r['created_at']})": r['id'] for r in requests.data}
                selected_req = st.selectbox("Select request", list(req_ids.keys()))
                new_status = st.selectbox("New Status", ["Pending", "Scheduled", "Done"])
                if st.button("Update Status"):
                    supabase.table("et_training_requests").update({"status": new_status}).eq("id", req_ids[selected_req]).execute()
                    st.success("Status updated!")
                    st.rerun()
        else:
            st.info("No training requests yet.")
    except:
        st.info("Training requests table not set up yet.")

    st.divider()
    st.markdown("#### ➕ Submit Training Request")
    emp_name_req  = st.text_input("Employee Name")
    topic_req     = st.text_input("Training Topic")
    reason_req    = st.text_area("Reason / Justification", placeholder="Why does this employee need this training?")
    urgency_req   = st.selectbox("Urgency", ["Low", "Medium", "High"])

    if st.button("Submit Request", type="primary"):
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
        else:
            st.info("No action plans yet.")
    except:
        st.info("Action plans table not set up yet.")

    if can_edit:
        st.divider()
        st.markdown("#### ➕ Add Action Plan")
        action_text = st.text_input("Action")
        owner_text  = st.text_input("Owner")
        due_date    = st.date_input("Due Date")
        status_ap   = st.selectbox("Status", ["Open", "In Progress", "Done"])
        notes_ap    = st.text_area("Notes", placeholder="Additional context...")

        if st.button("Add Action Plan", type="primary"):
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
                    st.success("✅ Action plan added!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {str(e)}")
