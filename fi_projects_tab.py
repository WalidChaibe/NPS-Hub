# fi_projects_tab.py
# Injected into 5_FI.py as a tab
# Call: render_fi_projects_tab(supabase, role, pillar, name)

import io, json, base64, math
from datetime import date, datetime, timedelta
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib as mpl
import streamlit as st
from reportlab.pdfgen import canvas as _rl_canvas
from reportlab.lib.colors import HexColor, black, white
from reportlab.lib.utils import ImageReader
from reportlab.lib.pagesizes import A4, landscape as _rl_landscape

mpl.rcParams["font.family"] = "DejaVu Sans"

# ════════════════════════════════════════
# CONSTANTS
# ════════════════════════════════════════
QUESTIONS = {
    1:  {"text": "Are team members listed?", "score": 1, "week": 1},
    2:  {"text": "Have all team members been assigned clear roles?", "score": 1, "week": 1},
    3:  {"text": "Is it clear why this problem was targeted? Link to company KPI?", "score": 4, "week": 1},
    4:  {"text": "Has a Cost/Benefit chart been introduced? Is it up-to-date?", "score": 4, "week": 5},
    5:  {"text": "Is the historical data (timeframe and actual value) clearly shown?", "score": 1, "week": 1},
    6:  {"text": "Is the performance indicator's target clearly shown?", "score": 1, "week": 2},
    7:  {"text": "Is the KPI subdivided into definable components?", "score": 1, "week": 1},
    8:  {"text": "Are Route and Master Plan clearly visible and up-to-date?", "score": 2, "week": 1},
    9:  {"text": "Is the target of each step clear?", "score": 1, "week": 2},
    10: {"text": "Have step targets been subdivided into specific activities?", "score": 1, "week": 3},
    11: {"text": "Has root-cause analysis been used and well documented?", "score": 1, "week": 4},
    12: {"text": "Have suspected causes been verified and quantified with data?", "score": 1, "week": 5},
    13: {"text": "Is data collection consistent across all shifts?", "score": 1, "week": 2},
    14: {"text": "Has the team used route methods/tools to attack problems?", "score": 2, "week": 5},
    15: {"text": "Is reoccurrence analysis present and updated?", "score": 1, "week": 7},
    16: {"text": "Is the single problem analysis applied and followed up?", "score": 1, "week": 7},
    17: {"text": "Have logical countermeasures been defined with sound logic?", "score": 2, "week": 6},
    18: {"text": "Have critical areas been restored to basic conditions?", "score": 3, "week": 4},
    19: {"text": "Are planned actions clearly visible with target completion dates?", "score": 5, "week": 4},
    20: {"text": "Is there an owner for each action?", "score": 2, "week": 5},
    21: {"text": "Is the action plan up-to-date?", "score": 2, "week": 5},
    22: {"text": "Is the majority of actions completed on time?", "score": 3, "week": 5},
    23: {"text": "Is there evidence of implemented actions?", "score": 4, "week": 5},
    24: {"text": "Is the trend of the performance indicator positive?", "score": 15, "week": 7},
    25: {"text": "Has the team achieved its goal or made substantial progress?", "score": 15, "week": 11},
    26: {"text": "Are procedures in place to hold the gains achieved?", "score": 2, "week": 10},
    27: {"text": "Are monitoring systems for key actions in place and visible?", "score": 2, "week": 7},
    28: {"text": "Are monitoring system devices used and up-to-date?", "score": 2, "week": 8},
    29: {"text": "Have OPLs/SOPs been created for every significant improvement?", "score": 2, "week": 10},
    30: {"text": "Is there a training matrix for OPLs/SOPs with a training plan?", "score": 2, "week": 10},
    31: {"text": "Are CIL standards clear? Do CIL audits achieve ≥90%?", "score": 3, "week": 5},
    32: {"text": "Is the workplace well organised (5S)?", "score": 1, "week": 5},
    33: {"text": "Are improvements on the machine/area evident?", "score": 1, "week": 5},
    34: {"text": "Is the methodology route well understood by all team members?", "score": 5, "week": 3},
    35: {"text": "Can a randomly picked team member explain the activity board?", "score": 3, "week": 3},
    36: {"text": "Are meetings organised and attendance at expected levels?", "score": 2, "week": 1},
}

TARGET_RAMP = {
    1:12, 2:15, 3:20, 4:27, 5:45, 6:47,
    7:62, 8:64, 9:64, 10:70, 11:85, 12:100
}

TEAM_ROLES = ["Team Leader","Analyst","Operator","Maintenance","Quality","Other"]
RCA_METHODS = ["5-Why","Fishbone","Pareto","FMEA","Other"]
ACTION_STATUSES = ["Open","In Progress","Completed","Overdue"]
MONITORING_TYPES = ["Checklist","Audit","Form","Visual Board","Other"]

# ── Plant sections & machines ──
PLANT_SECTIONS = {
    "Corrugator":     ["BHS","Fosber"],
    "Die-Cut":        ["BOBST 160-II","BOBST 203","BOBST MASTERCUT 1","BOBST MASTERCUT 2"],
    "FFG":            ["LMC FFG","MARTIN 616","924","SATURN"],
    "Folder Gluers":  ["Bahmüller TURBOX","VEGA 2"],
    "Stitcher":       ["BAHMULLER STITCHER"],
    "Printer":        ["IPACK"],
    "Pre-Print":      ["CI4","CI6"],
    "QuickSet":       ["QuickSet"],
    "Jumbo":          ["JUMBO"],
    "RM Warehouse":   [],
    "FG Warehouse":   [],
    "Maintenance":    [],
}

# ── KPI → KAI tree ──
# KPI = Key Performance Indicator (the big goal)
# KAI = Key Activity Indicator (measurable sub-actions that drive the KPI)
KPI_TREE = {
    "OEE Improvement": {
        "unit": "%",
        "kais": [
            "Reduce Breakdown Time",
            "Reduce Minor Stoppages",
            "Reduce Setup & Changeover Time",
            "Reduce Planned Maintenance Time",
            "Reduce Speed Losses",
            "Increase Availability Rate",
            "Increase Performance Rate",
            "Increase Quality Rate",
        ]
    },
    "Quality Defect Reduction": {
        "unit": "%",
        "kais": [
            "Reduce Customer Complaints Count",
            "Reduce Internal Defects (PPM)",
            "Reduce Rework Rate",
            "Reduce Scrap Rate",
            "Improve First Pass Yield",
            "Reduce NCR Count",
        ]
    },
    "Waste Reduction": {
        "unit": "%",
        "kais": [
            "Reduce Paper / Board Waste %",
            "Reduce Trim Waste",
            "Reduce Ink & Chemical Waste",
            "Reduce Energy Consumption",
            "Reduce Sheet Waste",
        ]
    },
    "Cost Reduction": {
        "unit": "K€",
        "kais": [
            "Reduce Maintenance Spend",
            "Reduce Material Costs",
            "Reduce Labour Overtime",
            "Reduce Energy Costs",
            "Reduce Rework & Scrap Costs",
        ]
    },
    "Safety Improvement": {
        "unit": "Count",
        "kais": [
            "Reduce Near Miss Incidents",
            "Reduce Lost Time Accidents",
            "Improve Safety Audit Score",
            "Increase Near Miss Reporting Rate",
            "Eliminate Unsafe Conditions (Tags)",
        ]
    },
    "Delivery Performance": {
        "unit": "%",
        "kais": [
            "Reduce Order Lead Time",
            "Improve Schedule Adherence",
            "Improve OTIF Rate",
            "Reduce Order Backlog",
        ]
    },
    "5S Score Improvement": {
        "unit": "Score",
        "kais": [
            "Improve Sort (Seiri) Score",
            "Improve Set in Order (Seiton) Score",
            "Improve Shine (Seiso) Score",
            "Improve Standardise (Seiketsu) Score",
            "Improve Sustain (Shitsuke) Score",
        ]
    },
    "Throughput / Productivity": {
        "unit": "MT",
        "kais": [
            "Increase Net Run Time",
            "Increase Average Speed",
            "Reduce Idle Time",
            "Increase Good Boards Production",
            "Improve Capacity Utilisation",
        ]
    },
}

UNITS = ["% (Percent)","MT","SAR","LM","SQM","GSM","Hits","Hits/Hour",
         "LM/Min","BD Time (hrs)","Hours","Mins","Secs","K€","Count","Score"]

DEFAULT_COMPANY_KPIS = ["Reduce Costs","Improve Customer Satisfaction","Increase OEE","Reduce Waste","Improve Safety","Improve Delivery Performance","Increase Productivity"]

# ════════════════════════════════════════
# HELPERS
# ════════════════════════════════════════
def _get_current_week(launch_date):
    if not launch_date: return 1
    if isinstance(launch_date, str): launch_date = date.fromisoformat(launch_date)
    delta = (date.today() - launch_date).days
    return max(1, min(12, math.ceil(delta / 7) if delta > 0 else 1))

def _to_b64(f):
    if f is None: return None
    return base64.b64encode(f.getvalue()).decode()

def _fig_to_png(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    buf.seek(0)
    return buf

def _gantt_chart(steps, weekly_updates, current_week):
    if not steps: return None
    n = len(steps)
    fig, ax = plt.subplots(figsize=(14, max(3, n*0.8+1.5)), dpi=200)
    fig.patch.set_facecolor("#FAFAFA")
    ax.set_facecolor("#FAFAFA")

    bar_h = 0.5
    for i, step in enumerate(steps):
        ps = step.get("planned_start_week",1)
        pe = step.get("planned_end_week",2)
        # Background planned bar
        ax.barh(i, pe-ps, left=ps-1, height=bar_h,
                color="#C8DFF0", alpha=1.0, zorder=2)
        ax.barh(i, pe-ps, left=ps-1, height=bar_h,
                color="none", edgecolor="#006394", linewidth=1.2, zorder=3)
        # Actual progress
        pct = 0
        for wu in weekly_updates:
            sp_raw = wu.get("step_progress") or []
            if isinstance(sp_raw, str):
                try: sp_raw = json.loads(sp_raw)
                except: sp_raw = []
            for sp in sp_raw:
                if not isinstance(sp, dict): continue
                if sp.get("step_id") == str(step.get("id","")):
                    pct = max(pct, sp.get("pct_complete",0))
        actual_w = (pe - ps) * pct / 100
        if actual_w > 0:
            ax.barh(i, actual_w, left=ps-1, height=bar_h,
                    color="#006394", alpha=0.85, zorder=4)
        # % label inside bar
        if pct > 0:
            ax.text(ps-1 + actual_w/2, i, f"{int(pct)}%",
                    ha="center", va="center", fontsize=8,
                    color="white", fontweight="bold", zorder=5)
        # Step owner label
        owner = step.get("owner","")
        if owner:
            ax.text(pe - 0.05, i + bar_h/2 + 0.05, owner,
                    ha="right", va="bottom", fontsize=7, color="#555555", zorder=5)

    # Current week line
    ax.axvline(current_week-1, color="#DE201B", linewidth=2,
               linestyle="--", zorder=6, label=f"Now (W{current_week})")
    ax.text(current_week-1+0.05, n-0.3, f"W{current_week}",
            color="#DE201B", fontsize=8, fontweight="bold", zorder=7)

    # Week grid lines
    for w in range(12):
        ax.axvline(w, color="#DDDDDD", linewidth=0.5, zorder=1)

    ax.set_yticks(range(n))
    ax.set_yticklabels([s.get("step_name","") for s in steps], fontsize=9, fontweight="bold")
    ax.set_xticks(range(12))
    ax.set_xticklabels([f"W{i+1}" for i in range(12)], fontsize=9)
    ax.set_xlim(0, 12)
    ax.set_ylim(-0.6, n-0.4)
    ax.invert_yaxis()
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.tick_params(left=False)

    # Legend
    from matplotlib.patches import Patch
    legend_elements = [
        Patch(facecolor="#C8DFF0", edgecolor="#006394", label="Planned"),
        Patch(facecolor="#006394", label="Actual Progress"),
        plt.Line2D([0],[0], color="#DE201B", linewidth=2, linestyle="--", label="Current Week"),
    ]
    ax.legend(handles=legend_elements, loc="lower right", fontsize=8,
              frameon=True, framealpha=0.9, edgecolor="#CCCCCC")
    plt.tight_layout(pad=1.5)
    return fig

def _kpi_trend_chart(kpi_data, baseline, target, weeks_data):
    fig, ax = plt.subplots(figsize=(10,4), dpi=120)
    weeks = [w["week_number"] for w in weeks_data if w.get("kpi_value") is not None]
    values = [w["kpi_value"] for w in weeks_data if w.get("kpi_value") is not None]
    if baseline is not None:
        ax.axhline(baseline, color="#888888", linewidth=1.5, linestyle=":", label=f"Baseline ({baseline})")
    if target is not None:
        ax.axhline(target, color="#27AE60", linewidth=1.5, linestyle="--", label=f"Target ({target})")
    if weeks and values:
        ax.plot(weeks, values, "o-", color="#006394", linewidth=2, markersize=6, label="Actual")
        for w,v in zip(weeks,values):
            ax.annotate(f"{v}", (w,v), textcoords="offset points", xytext=(0,8), fontsize=8, ha="center")
    ax.set_xlabel("Week"); ax.set_ylabel(kpi_data.get("unit","Value") if kpi_data else "Value")
    ax.set_xticks(range(1,13)); ax.set_xticklabels([f"W{i}" for i in range(1,13)], fontsize=8)
    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
    ax.legend(fontsize=9, frameon=False)
    plt.tight_layout()
    return fig

def _score_project(project, team, kpi, steps, weekly_updates, actions, stab, audit_records):
    scores = {q: 0 for q in range(1, 37)}  # default all zero
    stab = stab or {}
    team = team or []
    steps = steps or []
    actions = actions or []
    weekly_updates = weekly_updates or []
    audit_records = audit_records or []
    wu_by_week = {w["week_number"]: w for w in weekly_updates}

    def _safe_date(d):
        try: return date.fromisoformat(str(d)[:10])
        except: return None

    # Q1 — team members listed
    scores[1] = 1 if len(team) >= 1 else 0
    # Q2 — all have roles
    scores[2] = 1 if team and all(m.get("role") for m in team) else 0
    # Q3 — problem + company kpi link
    scores[3] = 4 if (project.get("problem_statement") and project.get("company_kpi_link")) else 0
    # Q4 — cost/benefit (week 5+)
    scores[4] = 0  # Q4 scored when cost/benefit section filled (checked separately)
    # Q5 — KPI baseline
    scores[5] = 1 if (kpi and kpi.get("kpi_name") and kpi.get("baseline_value") is not None) else 0
    # Q6 — KPI target
    scores[6] = 1 if (kpi and kpi.get("target_value") is not None and kpi.get("target_date")) else 0
    # Q7 — KPI sub-components
    scores[7] = 1 if (kpi and (kpi.get("sub_components") or project.get("kpi_not_applicable"))) else 0
    # Q8 — master plan steps
    scores[8] = 2 if len(steps) >= 2 else 0
    # Q9 — steps with targets
    scores[9] = 1 if all(s.get("planned_start_week") and s.get("planned_end_week") for s in steps) else 0
    # Q10 — sub-activities exist
    def _parse_sp(wu):
        sp = wu.get("step_progress") or []
        if isinstance(sp, str):
            try: return json.loads(sp)
            except: return []
        return sp
    scores[10] = 1 if any(_parse_sp(wu) for wu in weekly_updates) else 0
    # Q11 — RCA documented
    scores[11] = 1 if any(wu.get("rca_performed") and wu.get("rca_findings") for wu in weekly_updates) else 0
    # Q12 — causes verified
    scores[12] = 1 if any(wu.get("causes_verified") for wu in weekly_updates) else 0
    # Q13 — shifts covered
    scores[13] = 1 if any(wu.get("shifts_covered") and "All" in (wu.get("shifts_covered") or "") for wu in weekly_updates) else 0
    # Q14 — structured RCA tool used
    structured = {"5-Why","Fishbone","Pareto","FMEA"}
    scores[14] = 2 if any(wu.get("rca_method") in structured for wu in weekly_updates) else 0
    # Q15 — reoccurrence analysis
    scores[15] = 1 if any(wu.get("reoccurrence_description") for wu in weekly_updates) else 0
    # Q16 — single problem analysis
    scores[16] = 1 if any(wu.get("single_problem_analysis") and wu.get("single_problem_notes") for wu in weekly_updates) else 0
    # Q17 — countermeasures linked to root causes
    scores[17] = 2 if actions and all(a.get("root_cause_addressed") for a in actions) else 0
    # Q18 — basic conditions
    scores[18] = 3 if any(wu.get("basic_conditions_restored") and wu.get("basic_conditions_after_b64") for wu in weekly_updates) else 0
    # Q19 — actions with dates
    scores[19] = 5 if actions and any(a.get("target_date") for a in actions) else 0
    # Q20 — every action has owner
    scores[20] = 2 if actions and all(a.get("owner") for a in actions) else 0
    # Q21 — action plan updated this week
    current_week = _get_current_week(project.get("launch_date"))
    recent_actions = [a for a in actions if a.get("created_week") == current_week]
    scores[21] = 2 if recent_actions else 0
    # Q22 — majority completed on time
    past_due = [a for a in actions if a.get("target_date") and _safe_date(a["target_date"]) and _safe_date(a["target_date"]) < date.today()]
    on_time = [a for a in past_due if a.get("status") == "Completed"]
    scores[22] = 3 if (past_due and len(on_time)/len(past_due) >= 0.5) else 0
    # Q23 — evidence uploaded
    scores[23] = 4 if any(a.get("evidence_b64") for a in actions if a.get("status")=="Completed") else 0
    # Q24 — positive trend
    kpi_vals = [wu["kpi_value"] for wu in sorted(weekly_updates, key=lambda x:x["week_number"]) if wu.get("kpi_value") is not None]
    if kpi and kpi.get("target_value") and len(kpi_vals) >= 3:
        baseline = kpi.get("baseline_value", kpi_vals[0])
        target_val = kpi.get("target_value")
        improving = (target_val > baseline and kpi_vals[-1] > kpi_vals[-3]) or (target_val < baseline and kpi_vals[-1] < kpi_vals[-3])
        scores[24] = 15 if improving else 0
    else:
        scores[24] = 0
    # Q25 — goal achieved or 80% progress
    if kpi and kpi.get("target_value") and kpi.get("baseline_value") is not None and kpi_vals:
        baseline = kpi.get("baseline_value")
        target_val = kpi.get("target_value")
        current_val = kpi_vals[-1]
        total_gap = abs(target_val - baseline)
        progress = abs(current_val - baseline) / total_gap * 100 if total_gap > 0 else 0
        scores[25] = 15 if progress >= 80 else 0
    else:
        scores[25] = 0
    # Q26 — procedures
    scores[26] = 2 if stab and stab.get("procedures_created") and stab.get("procedures") else 0
    # Q27 — monitoring in place
    scores[27] = 2 if stab and stab.get("monitoring_in_place") else 0
    # Q28 — monitoring active and updated
    scores[28] = 2 if stab and stab.get("monitoring_active") and stab.get("monitoring_last_update") else 0
    # Q29 — OPLs created
    scores[29] = 2 if stab and stab.get("opls") else 0
    # Q30 — training matrix
    scores[30] = 2 if stab and stab.get("training_matrix") else 0
    # Q31 — CIL score >= 90
    scores[31] = 3 if stab and stab.get("cil_standards_defined") and stab.get("cil_audit_score") and float(stab.get("cil_audit_score",0)) >= 90 else 0
    # Q32 — 5S rating >= 3
    scores[32] = 1 if stab and stab.get("five_s_rating") and int(stab.get("five_s_rating",0)) >= 3 else 0
    # Q33 — improvement photos
    scores[33] = 1 if stab and stab.get("improvements_visible") and stab.get("improvements_photos_b64") else 0
    # Q34, Q35 — from audit records
    for ar in audit_records:
        qs = ar.get("question_scores") or {}
        if qs.get("34"): scores[34] = 5
        if qs.get("35"): scores[35] = 3
    scores.setdefault(34, 0); scores.setdefault(35, 0)
    # Q36 — meetings
    scores[36] = 2 if any(wu.get("meeting_held") for wu in weekly_updates) else 0

    total = sum(scores.values())
    return scores, total

# ════════════════════════════════════════
# PDF GENERATOR
# ════════════════════════════════════════
def _generate_project_pdf(project, team, kpi, steps, weekly_updates, actions, stab, audit_records, scores, total_score):
    W, H = A4
    buf = io.BytesIO()
    c = _rl_canvas.Canvas(buf, pagesize=A4)
    M = 40

    def new_page():
        c.showPage()

    def header_bar(title, y=H-50):
        c.setFillColor(HexColor("#006394"))
        c.rect(M, y, W-2*M, 22, fill=1, stroke=0)
        c.setFillColor(white)
        c.setFont("Helvetica-Bold", 12)
        c.drawString(M+8, y+6, title)
        return y - 10

    def body_text(text, x, y, w, font="Helvetica", size=9, color="#000000"):
        c.setFillColor(HexColor(color))
        c.setFont(font, size)
        words = str(text).split()
        line = ""; lines = []
        for word in words:
            test = (line+" "+word).strip()
            if c.stringWidth(test, font, size) < w:
                line = test
            else:
                if line: lines.append(line)
                line = word
        if line: lines.append(line)
        for ln in lines:
            c.drawString(x, y, ln)
            y -= size + 3
        return y

    # ── PAGE 1: COVER ──
    c.setFillColor(HexColor("#006394"))
    c.rect(0, H-120, W, 120, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 22)
    c.drawCentredString(W/2, H-55, "FUNCTIONAL IMPROVEMENT PROJECT")
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(W/2, H-80, project.get("project_name",""))
    c.setFont("Helvetica", 10)
    c.drawCentredString(W/2, H-100, f"Final Report — Generated {date.today().strftime('%d %B %Y')}")

    # Score badge
    score_color = "#27AE60" if total_score >= 70 else "#E67E22" if total_score >= 45 else "#DE201B"
    c.setFillColor(HexColor(score_color))
    c.circle(W/2, H-200, 45, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 28)
    c.drawCentredString(W/2, H-207, str(int(total_score)))
    c.setFont("Helvetica", 10)
    c.drawCentredString(W/2, H-222, "/ 100")

    # Project info
    info_y = H-290
    for label, val in [
        ("Problem Statement", project.get("problem_statement","")),
        ("Target Area", project.get("target_area","")),
        ("Launch Date", str(project.get("launch_date",""))),
        ("Expected Completion", str(project.get("expected_completion_date",""))),
        ("Team Size", f"{len(team)} members"),
    ]:
        c.setFillColor(HexColor("#444444"))
        c.setFont("Helvetica-Bold", 9)
        c.drawString(M, info_y, f"{label}:")
        c.setFont("Helvetica", 9)
        c.drawString(M+130, info_y, str(val)[:70])
        info_y -= 18

    new_page()

    # ── PAGE 2: TEAM & KPI ──
    y = H - M
    y = header_bar("PROJECT OVERVIEW", y)
    y -= 20

    # Team
    c.setFont("Helvetica-Bold", 11); c.setFillColor(black)
    c.drawString(M, y, "Team Members"); y -= 16
    for m in team:
        c.setFont("Helvetica", 9)
        c.drawString(M+10, y, f"• {m.get('member_name','')} — {m.get('role','')} ({m.get('department','')})")
        y -= 13
    y -= 10

    # KPI
    if kpi:
        c.setFont("Helvetica-Bold", 11); c.setFillColor(black)
        c.drawString(M, y, "KPI Summary"); y -= 16
        for label, val in [
            ("KPI Name", kpi.get("kpi_name","")),
            ("Unit", kpi.get("unit","")),
            ("Baseline", kpi.get("baseline_value","")),
            ("Target", kpi.get("target_value","")),
            ("Target Date", str(kpi.get("target_date",""))),
        ]:
            c.setFont("Helvetica", 9)
            c.drawString(M+10, y, f"{label}: {val}")
            y -= 13

    new_page()

    # ── PAGE 3: GANTT ──
    y = H - M
    y = header_bar("MASTER PLAN — GANTT CHART", y)
    y -= 20
    if steps:
        wu_list = weekly_updates
        current_week = _get_current_week(project.get("launch_date"))
        gantt_fig = _gantt_chart(steps, wu_list, current_week)
        if gantt_fig:
            gantt_buf = _fig_to_png(gantt_fig)
            gantt_img = ImageReader(gantt_buf)
            c.drawImage(gantt_img, M, y-220, width=W-2*M, height=220, preserveAspectRatio=True)
            plt.close(gantt_fig)
            y -= 230

    new_page()

    # ── PAGE 4: KPI TREND ──
    y = H - M
    y = header_bar("KPI TREND CHART", y)
    y -= 20
    if kpi:
        trend_fig = _kpi_trend_chart(kpi, kpi.get("baseline_value"), kpi.get("target_value"), weekly_updates)
        trend_buf = _fig_to_png(trend_fig)
        trend_img = ImageReader(trend_buf)
        c.drawImage(trend_img, M, y-220, width=W-2*M, height=220, preserveAspectRatio=True)
        plt.close(trend_fig)

    new_page()

    # ── PAGE 5: ACTIONS ──
    y = H - M
    y = header_bar("ACTION PLAN SUMMARY", y)
    y -= 20
    if actions:
        cols = ["Description","Owner","Target Date","Status"]
        col_ws = [210, 80, 80, 60]
        cx = M
        c.setFillColor(HexColor("#006394"))
        c.rect(M, y-14, W-2*M, 14, fill=1, stroke=0)
        c.setFillColor(white); c.setFont("Helvetica-Bold", 8)
        for col, cw in zip(cols, col_ws):
            c.drawString(cx+2, y-10, col); cx += cw
        y -= 18
        for i, action in enumerate(actions):
            if y < 60: new_page(); y = H-M
            bg = HexColor("#F5F5F5") if i%2==0 else white
            c.setFillColor(bg)
            c.rect(M, y-12, W-2*M, 13, fill=1, stroke=0)
            c.setFillColor(black); c.setFont("Helvetica", 8)
            cx = M
            status_color = "#27AE60" if action.get("status")=="Completed" else "#DE201B" if action.get("status")=="Overdue" else "#000000"
            vals = [action.get("description","")[:40], action.get("owner",""), str(action.get("target_date","")), action.get("status","")]
            for j,(val,cw) in enumerate(zip(vals,col_ws)):
                if j==3: c.setFillColor(HexColor(status_color))
                c.drawString(cx+2, y-9, str(val))
                c.setFillColor(black)
                cx += cw
            y -= 14

    new_page()

    # ── PAGE 6: AUDIT SCORE ──
    y = H - M
    y = header_bar("AUDIT SCORE DETAIL", y)
    y -= 16
    c.setFont("Helvetica-Bold", 10)
    c.drawString(M, y, f"Total Score: {int(total_score)} / 100")
    y -= 20
    # Per question table
    hdrs = ["Q#","Question","Weight","Status"]
    hw = [25, 280, 40, 60]
    c.setFillColor(HexColor("#EEEEEE")); c.rect(M, y-13, W-2*M, 13, fill=1, stroke=0)
    c.setFillColor(black); c.setFont("Helvetica-Bold", 8)
    cx = M
    for h, w2 in zip(hdrs, hw):
        c.drawString(cx+2, y-10, h); cx += w2
    y -= 15
    for qn, qdata in QUESTIONS.items():
        if y < 50: new_page(); y = H-M
        score = scores.get(qn, 0)
        satisfied = score >= qdata["score"]
        bg = HexColor("#E8F5E9") if satisfied else HexColor("#FFEBEE")
        c.setFillColor(bg); c.rect(M, y-11, W-2*M, 11, fill=1, stroke=0)
        c.setFillColor(black); c.setFont("Helvetica", 7.5)
        cx = M
        vals = [str(qn), qdata["text"][:55], str(qdata["score"]), "✓ Met" if satisfied else "✗ Not Met"]
        for val, w2 in zip(vals, hw):
            c.drawString(cx+2, y-8, val); cx += w2
        y -= 13

    c.save()
    buf.seek(0)
    return buf


# ════════════════════════════════════════
# MAIN RENDER FUNCTION
# ════════════════════════════════════════
def render_fi_projects_tab(supabase, role, pillar, name):
    can_create = (role == "plant_manager") or (role == "pillar_leader" and pillar == "FI")
    is_auditor = (role in ["plant_manager", "pillar_leader"])

    st.markdown("### 📋 FI Projects")

    # ── Project selector ──
    try:
        proj_resp = supabase.table("fi_projects").select("*").order("created_at", desc=True).execute()
        all_projects = proj_resp.data or []
    except Exception as e:
        st.error(f"Could not load projects: {e}"); return

    col_sel, col_new = st.columns([4,1])
    if not all_projects:
        st.info("No FI projects yet. Create one to get started.")
        selected_project = None
    else:
        proj_names = {p["id"]: f"{p['project_name']} (Started: {str(p.get('launch_date',''))[:10]})" for p in all_projects}
        sel_id = col_sel.selectbox("Select Project", list(proj_names.keys()),
                                    format_func=lambda x: proj_names[x], key="fi_proj_select")
        selected_project = next((p for p in all_projects if p["id"] == sel_id), None)

    if can_create:
        if col_new.button("➕ New Project", key="fi_new_proj"):
            st.session_state["fi_creating_project"] = True

    # ── New project modal ──
    if st.session_state.get("fi_creating_project") and can_create:
        with st.expander("🆕 Create New Project", expanded=True):
            # Section/machine selector outside form
            _np_cols = st.columns(2)
            _new_section = _np_cols[0].selectbox("Section / Area *", list(PLANT_SECTIONS.keys()), key="fi_new_section")
            _new_machines = PLANT_SECTIONS.get(_new_section,[])
            if _new_machines:
                _new_machine = _np_cols[1].selectbox("Machine", ["— All / General —"]+_new_machines, key="fi_new_machine")
                _new_area = f"{_new_section} — {_new_machine}" if _new_machine != "— All / General —" else _new_section
            else:
                _np_cols[1].caption("No sub-machines for this section")
                _new_area = _new_section

            with st.form("fi_new_proj_form"):
                np1, np2 = st.columns(2)
                new_name    = np1.text_input("Project Name *")
                # Section + machine for new project
                new_section = np2.selectbox("Section / Area *", list(PLANT_SECTIONS.keys()), key="fi_new_section")
                _new_machines = PLANT_SECTIONS.get(new_section,[])
                if _new_machines:
                    new_machine = st.selectbox("Machine", ["— All / General —"]+_new_machines, key="fi_new_machine")
                    new_area = f"{new_section} — {new_machine}" if new_machine != "— All / General —" else new_section
                else:
                    new_area = new_section
                new_problem  = st.text_area("Problem Statement *", height=80)
                nd1, nd2 = st.columns(2)
                new_launch   = nd1.date_input("Launch Date", value=date.today())
                new_expected = nd2.date_input("Expected Completion (auto: 12 weeks)",
                    value=new_launch+timedelta(weeks=12) if 'new_launch' in dir() else date.today()+timedelta(weeks=12))
                new_kpi_link = st.text_input("Which company / plant KPI does this project support?")
                new_kpi_imp  = st.text_area("How does this project impact that KPI?", height=60)
                sub = st.form_submit_button("🚀 Create Project", type="primary")
                if sub:
                    if not new_name or not new_problem:
                        st.error("Project Name and Problem Statement are required.")
                    else:
                        resp = supabase.table("fi_projects").insert({
                            "project_name": new_name,
                            "problem_statement": new_problem,
                            "target_area": _new_area,
                            "launch_date": str(new_launch),
                            "expected_completion_date": str(new_expected),
                            "company_kpi_link": new_kpi_link,
                            "kpi_impact": new_kpi_imp,
                            "created_by": name,
                        }).execute()
                        st.success("✅ Project created!")
                        st.session_state["fi_creating_project"] = False
                        st.rerun()
            if st.button("Cancel", key="fi_cancel_new"):
                st.session_state["fi_creating_project"] = False
                st.rerun()

    # ── Plant Manager: Company KPI Settings ──
    if role == "plant_manager":
        with st.expander("⚙️ Manage Company KPIs (Plant Manager Only)", expanded=False):
            st.caption("Add custom company KPIs that appear in the project setup dropdown.")
            try:
                _existing_ckpis = supabase.table("fi_company_kpis").select("*").execute().data or []
            except Exception:
                _existing_ckpis = []
            if _existing_ckpis:
                st.dataframe(pd.DataFrame(_existing_ckpis)[["kpi_name"]], use_container_width=True, hide_index=True)
            with st.form("fi_add_company_kpi"):
                new_ckpi = st.text_input("New Company KPI")
                if st.form_submit_button("➕ Add"):
                    if new_ckpi:
                        try:
                            supabase.table("fi_company_kpis").insert({"kpi_name":new_ckpi}).execute()
                            st.success("✅ Added"); st.rerun()
                        except Exception as e:
                            st.error(f"Error: {e}")

    if not selected_project:
        return

    # ── Load project data ──
    pid = selected_project["id"]
    current_week = _get_current_week(selected_project.get("launch_date"))

    try:
        team       = (supabase.table("fi_project_team").select("*").eq("project_id",pid).execute().data or [])
        kpi_rows   = (supabase.table("fi_project_kpi").select("*").eq("project_id",pid).execute().data or [])
        kpi        = kpi_rows[0] if kpi_rows else None
        steps      = (supabase.table("fi_project_steps").select("*").eq("project_id",pid).order("sort_order").execute().data or [])
        cost_rows  = (supabase.table("fi_project_cost").select("*").eq("project_id",pid).execute().data or [])
        cost       = cost_rows[0] if cost_rows else None
        wu_rows    = (supabase.table("fi_weekly_updates").select("*").eq("project_id",pid).order("week_number").execute().data or [])
        actions    = (supabase.table("fi_actions").select("*").eq("project_id",pid).order("created_at").execute().data or [])
        stab_rows  = (supabase.table("fi_stabilisation").select("*").eq("project_id",pid).execute().data or [])
        stab       = stab_rows[0] if stab_rows else None
        audit_records = (supabase.table("fi_audit_records").select("*").eq("project_id",pid).order("week_number").execute().data or [])
    except Exception as e:
        st.error(f"Error loading project data: {e}"); return

    wu_by_week = {w["week_number"]: w for w in wu_rows}
    current_wu = wu_by_week.get(current_week, {})

    # Week badge
    st.markdown(f"**Project:** {selected_project['project_name']} &nbsp;|&nbsp; "
                f"**Area:** {selected_project.get('target_area','')} &nbsp;|&nbsp; "
                f"**Current Week:** <span style='background:#006394;color:white;padding:2px 10px;border-radius:10px;'>W{current_week}</span>",
                unsafe_allow_html=True)

    # ── Sub-tabs ──
    audit_tabs = ["📋 Project Setup","📅 Weekly Update","📈 Results","🔒 Stabilisation"]
    if is_auditor:
        audit_tabs.append("🔍 Audit View")
    subtabs = st.tabs(audit_tabs)

    # ════════════════════════════════════════
    # SUB-TAB 1 — PROJECT SETUP
    # ════════════════════════════════════════
    with subtabs[0]:
        st.markdown("#### 📋 Project Setup")
        st.caption("Complete this once when starting the project. All Week 1 requirements live here.")

        # ── Section A: Project Identity ──
        with st.expander("🏷️ Project Identity", expanded=True):
            # Section selector OUTSIDE form so machine dropdown reacts dynamically
            _cur_area = selected_project.get("target_area","") or ""
            _cur_section = _cur_area.split(" — ")[0] if " — " in _cur_area else list(PLANT_SECTIONS.keys())[0]
            _cur_machine = _cur_area.split(" — ")[1] if " — " in _cur_area else ""
            _area_cols = st.columns(2)
            setup_section = _area_cols[0].selectbox("Section / Area *",
                list(PLANT_SECTIONS.keys()),
                index=list(PLANT_SECTIONS.keys()).index(_cur_section) if _cur_section in PLANT_SECTIONS else 0,
                key="fi_setup_section")
            _machines = PLANT_SECTIONS.get(setup_section, [])
            if _machines:
                _mach_opts = ["— All / General —"] + _machines
                _mach_idx  = (_machines.index(_cur_machine)+1) if _cur_machine in _machines else 0
                setup_machine = _area_cols[1].selectbox("Machine", _mach_opts,
                    index=_mach_idx, key="fi_setup_machine")
                setup_area = f"{setup_section} — {setup_machine}" if setup_machine != "— All / General —" else setup_section
            else:
                _area_cols[1].caption("No sub-machines for this section")
                setup_area = setup_section

            with st.form("fi_setup_identity"):
                setup_name = st.text_input("Project Name *", value=selected_project.get("project_name",""))
                st.info(f"📍 Target Area: **{setup_area}**")

                setup_problem = st.text_area("Problem Statement *", value=selected_project.get("problem_statement",""), height=80)

                d1, d2 = st.columns(2)
                setup_launch = d1.date_input("Launch Date",
                    value=date.fromisoformat(str(selected_project.get("launch_date", date.today()))[:10]))
                # Expected completion auto-set to 12 weeks from launch
                _auto_end = date.fromisoformat(str(selected_project.get("launch_date", date.today()))[:10]) + timedelta(weeks=12)
                setup_end = d2.date_input("Expected Completion (auto: 12 weeks)",
                    value=date.fromisoformat(str(selected_project.get("expected_completion_date", _auto_end))[:10]))

                # Company KPI dropdown
                _company_kpis = DEFAULT_COMPANY_KPIS.copy()
                try:
                    _custom_kpis = supabase.table("fi_company_kpis").select("kpi_name").execute().data or []
                    _company_kpis += [k["kpi_name"] for k in _custom_kpis if k.get("kpi_name") not in _company_kpis]
                except Exception:
                    pass
                _cur_kpi_link = selected_project.get("company_kpi_link","")
                _kpi_idx = _company_kpis.index(_cur_kpi_link) if _cur_kpi_link in _company_kpis else 0
                setup_kpi_link = st.selectbox("Which company/plant KPI does this project support?",
                    _company_kpis, index=_kpi_idx)
                setup_kpi_imp = st.text_area("How does this project impact that KPI?",
                    value=selected_project.get("kpi_impact",""), height=60)

                if st.form_submit_button("💾 Save Project Identity"):
                    supabase.table("fi_projects").update({
                        "project_name": setup_name, "target_area": setup_area,
                        "problem_statement": setup_problem,
                        "launch_date": str(setup_launch),
                        "expected_completion_date": str(setup_end),
                        "company_kpi_link": setup_kpi_link, "kpi_impact": setup_kpi_imp,
                    }).eq("id", pid).execute()
                    st.success("✅ Saved"); st.rerun()

        # ── Section B: Team Members ──
        with st.expander("👥 Team Members", expanded=True):
            st.caption("Add everyone on the team with their role.")
            if team:
                # Show team with their KPI-assigned targets
                _subs = []
                if kpi:
                    _raw = kpi.get("sub_components") or []
                    if isinstance(_raw, str):
                        try: _raw = json.loads(_raw)
                        except: _raw = []
                    _subs = [s for s in _raw if isinstance(s,dict)]
                _owner_map = {}
                for s in _subs:
                    owner = s.get("owner","")
                    if owner:
                        _owner_map.setdefault(owner,[]).append(
                            f"{s.get('name','')} ({s.get('baseline','')}→{s.get('target','')} {s.get('unit','')})"
                        )
                df_team = pd.DataFrame([{
                    "Name": m["member_name"],
                    "Role": m.get("role",""),
                    "Department": m.get("department",""),
                    "KPI Targets": " | ".join(_owner_map.get(m["member_name"],["—"]))
                } for m in team])
                st.dataframe(df_team, use_container_width=True, hide_index=True)
                if can_create:
                    del_name = st.selectbox("Remove member", ["—"]+[m["member_name"] for m in team], key="fi_del_member")
                    if st.button("🗑️ Remove", key="fi_remove_member") and del_name != "—":
                        mid = next((m["id"] for m in team if m["member_name"]==del_name), None)
                        if mid:
                            supabase.table("fi_project_team").delete().eq("id",mid).execute()
                            st.rerun()
            with st.form("fi_add_member"):
                mc1,mc2,mc3 = st.columns(3)
                m_name = mc1.text_input("Name *")
                m_role = mc2.selectbox("Role", TEAM_ROLES)
                m_dept = mc3.text_input("Department")
                st.caption("💡 Contribution targets are assigned per KPI focus area in the KPI section below.")
                if st.form_submit_button("➕ Add Team Member"):
                    if m_name:
                        supabase.table("fi_project_team").insert({
                            "project_id":pid,"member_name":m_name,"role":m_role,
                            "department":m_dept,"contribution_target":""
                        }).execute()
                        st.rerun()

        # ── Section C: KPI & KAI ──
        with st.expander("📊 KPI & KAI", expanded=True):
            kpi_vals = kpi or {}
            _launch  = selected_project.get("launch_date", date.today())
            _auto_dt = date.fromisoformat(str(_launch)[:10]) + timedelta(weeks=12)

            # Existing KAIs
            _existing_subs = kpi_vals.get("sub_components") or []
            if isinstance(_existing_subs, str):
                try: _existing_subs = json.loads(_existing_subs)
                except: _existing_subs = []

            # ── KPI selector OUTSIDE form so KAIs react ──
            _kpi_cats = list(KPI_TREE.keys())
            _cur_cat  = kpi_vals.get("kpi_category","OEE Improvement")
            _cat_idx  = _kpi_cats.index(_cur_cat) if _cur_cat in _kpi_cats else 0

            st.markdown("**Step 1 — Select the KPI this project targets:**")
            kpi_category = st.selectbox("KPI", _kpi_cats, index=_cat_idx, key="fi_kpi_cat")
            _def_unit    = KPI_TREE[kpi_category]["unit"]
            _kai_opts    = KPI_TREE[kpi_category]["kais"]

            # KAI selector OUTSIDE form so it reacts to KPI change
            st.markdown("**Step 2 — Which KAIs will this project address?**")
            st.caption(f"KAIs are the sub-activities that drive the **{kpi_category}** KPI.")
            _existing_kai_names = [s.get("name","") for s in _existing_subs if isinstance(s,dict)]
            _def_kais = [k for k in _existing_kai_names if k in _kai_opts]
            selected_kais = st.multiselect("KAIs (select all that apply)",
                _kai_opts, default=_def_kais, key="fi_kai_select")

            # ── Unified form: KPI + KAIs together ──
            # KPI overall baseline/target — compact, outside form style
            st.markdown(f"##### 🎯 {kpi_category}")
            with st.form("fi_kpi_kai_form"):
                # KPI numbers: just baseline, target, unit, date
                _pk1,_pk2,_pk3,_pk4 = st.columns(4)
                _pk1.caption("Unit")
                _pk2.caption("Baseline")
                _pk3.caption("Target")
                _pk4.caption("Target Date")
                _cur_unit = kpi_vals.get("unit",_def_unit)
                _u_idx    = UNITS.index(_cur_unit) if _cur_unit in UNITS else 0
                kpi_unit     = _pk1.selectbox("Unit", UNITS, index=_u_idx, label_visibility="collapsed")
                kpi_baseline = _pk2.number_input("Baseline", label_visibility="collapsed",
                    value=float(kpi_vals.get("baseline_value",0) or 0))
                kpi_target   = _pk3.number_input("Target", label_visibility="collapsed",
                    value=float(kpi_vals.get("target_value",0) or 0))
                kpi_tdate    = _pk4.date_input("Target Date", label_visibility="collapsed",
                    value=date.fromisoformat(str(kpi_vals.get("target_date",_auto_dt))[:10])
                    if kpi_vals.get("target_date") else _auto_dt)

                # KAI rows
                kai_rows = []
                if selected_kais:
                    st.divider()
                    _rh1,_rh2,_rh3,_rh4,_rh5 = st.columns([3,1.5,1.2,1.2,2])
                    _rh1.caption("KAI"); _rh2.caption("Unit"); _rh3.caption("Baseline"); _rh4.caption("Target"); _rh5.caption("Responsible")
                    _sub_by_name = {s.get("name",""): s for s in _existing_subs if isinstance(s,dict)}
                    _team_names  = [""]+[m["member_name"] for m in team]
                    for si, kai in enumerate(selected_kais):
                        ex   = _sub_by_name.get(kai,{})
                        k1,k2,k3,k4,k5 = st.columns([3,1.5,1.2,1.2,2])
                        k1.markdown(f"↳ {kai}")
                        _ku  = ex.get("unit",_def_unit)
                        _kui = UNITS.index(_ku) if _ku in UNITS else 0
                        _ko  = ex.get("owner","")
                        _koi = _team_names.index(_ko) if _ko in _team_names else 0
                        kai_rows.append({
                            "name":     kai,
                            "unit":     k2.selectbox("Unit", UNITS, index=_kui, key=f"kai_u_{si}", label_visibility="collapsed"),
                            "baseline": k3.number_input("Baseline", value=float(ex.get("baseline",0)), key=f"kai_b_{si}", label_visibility="collapsed"),
                            "target":   k4.number_input("Target",   value=float(ex.get("target",0)),   key=f"kai_t_{si}", label_visibility="collapsed"),
                            "owner":    k5.selectbox("Responsible", _team_names, index=_koi, key=f"kai_o_{si}", label_visibility="collapsed"),
                        })

                if st.form_submit_button("💾 Save KPI & KAIs", type="primary"):
                    _save = {
                        "project_id": pid, "kpi_name": kpi_category, "unit": kpi_unit,
                        "kpi_category": kpi_category, "sub_kpi_focus": ",".join(selected_kais),
                        "baseline_value": kpi_baseline, "target_value": kpi_target,
                        "target_date": str(kpi_tdate), "sub_components": json.dumps(kai_rows),
                    }
                    if kpi: supabase.table("fi_project_kpi").update(_save).eq("id",kpi["id"]).execute()
                    else:   supabase.table("fi_project_kpi").insert(_save).execute()
                    st.success("✅ Saved"); st.rerun()

            # Summary card
            if kpi and _existing_subs:
                st.divider()
                st.markdown(f"**Current KPI:** {kpi_vals.get('kpi_name','')} — "
                            f"Baseline: **{kpi_vals.get('baseline_value','')}** → "
                            f"Target: **{kpi_vals.get('target_value','')}** {kpi_vals.get('unit','')}")
                _kd = pd.DataFrame([{
                    "KAI": s.get("name",""),
                    "Unit": s.get("unit",""),
                    "Baseline": s.get("baseline",""),
                    "Target": s.get("target",""),
                    "Responsible": s.get("owner","—"),
                } for s in _existing_subs if isinstance(s,dict)])
                st.dataframe(_kd, use_container_width=True, hide_index=True)

        # ── Section D: Cost/Benefit (unlocks week 5) ──
        if current_week >= 5:
            with st.expander("💰 Cost / Benefit", expanded=False):
                cost_vals = cost or {}
                with st.form("fi_cost_form"):
                    cb1,cb2 = st.columns(2)
                    cost_est  = cb1.number_input("Estimated Project Cost (K€)", value=float(cost_vals.get("estimated_cost",0) or 0))
                    cost_save = cb2.number_input("Estimated Annual Savings (K€/yr)", value=float(cost_vals.get("estimated_savings",0) or 0))
                    payback = round(cost_est / cost_save * 12, 1) if cost_save > 0 else 0
                    st.info(f"📐 Payback Period: **{payback} months** (auto-calculated)")
                    if st.form_submit_button("💾 Save Cost/Benefit"):
                        cb_data = {"project_id":pid,"estimated_cost":cost_est,"estimated_savings":cost_save,"payback_months":payback}
                        if cost:
                            supabase.table("fi_project_cost").update(cb_data).eq("id",cost["id"]).execute()
                        else:
                            supabase.table("fi_project_cost").insert(cb_data).execute()
                        st.success("✅ Saved"); st.rerun()
                if cost:
                    fig_cb, ax_cb = plt.subplots(figsize=(5,3), dpi=100)
                    ax_cb.bar(["Project Cost","Annual Savings"],
                              [cost.get("estimated_cost",0), cost.get("estimated_savings",0)],
                              color=["#DE201B","#27AE60"], width=0.5)
                    ax_cb.set_ylabel("K€")
                    ax_cb.spines["top"].set_visible(False); ax_cb.spines["right"].set_visible(False)
                    plt.tight_layout()
                    st.pyplot(fig_cb); plt.close(fig_cb)

        # ── Section E: Master Plan / Gantt ──
        with st.expander("🗺️ Master Plan & Gantt Chart", expanded=True):
            with st.form("fi_add_step"):
                st.caption("Define the project phases/steps to auto-generate a Gantt chart.")
                ms1,ms2 = st.columns(2)
                step_name = ms1.text_input("Step / Phase Name *")
                step_desc = ms2.text_input("Description")
                mw1,mw2,mw3 = st.columns(3)
                step_start = mw1.number_input("Planned Start Week", min_value=1, max_value=12, value=1)
                step_end   = mw2.number_input("Planned End Week",   min_value=1, max_value=12, value=2)
                step_owner = mw3.selectbox("Owner", [""]+[m["member_name"] for m in team])
                if st.form_submit_button("➕ Add Step"):
                    if step_name:
                        supabase.table("fi_project_steps").insert({
                            "project_id":pid,"step_name":step_name,"description":step_desc,
                            "planned_start_week":int(step_start),"planned_end_week":int(step_end),
                            "owner":step_owner,"sort_order":len(steps)
                        }).execute()
                        st.rerun()

            if steps:
                df_steps = pd.DataFrame(steps)[["step_name","description","planned_start_week","planned_end_week","owner"]]
                df_steps.columns = ["Step","Description","Start Week","End Week","Owner"]
                st.dataframe(df_steps, use_container_width=True, hide_index=True)
                gantt_fig = _gantt_chart(steps, wu_rows, current_week)
                if gantt_fig:
                    st.pyplot(gantt_fig); plt.close(gantt_fig)
                # Delete step
                if can_create:
                    del_step = st.selectbox("Remove step", ["—"]+[s["step_name"] for s in steps], key="fi_del_step")
                    if st.button("🗑️ Remove Step", key="fi_remove_step") and del_step != "—":
                        sid = next((s["id"] for s in steps if s["step_name"]==del_step), None)
                        if sid:
                            supabase.table("fi_project_steps").delete().eq("id",sid).execute()
                            st.rerun()

    # ════════════════════════════════════════
    # SUB-TAB 2 — WEEKLY UPDATE
    # ════════════════════════════════════════
    with subtabs[1]:
        st.markdown(f"#### 📅 Weekly Update — Week {current_week}")
        st.caption("This is your weekly check-in. Fill in what happened this week.")

        # Week selector
        sel_week = st.selectbox("View / edit week:", list(range(1,13)),
                                 index=current_week-1, key="fi_week_sel")
        wu = wu_by_week.get(sel_week, {})
        wu_id = wu.get("id")

        def _save_wu(updates):
            if wu_id:
                supabase.table("fi_weekly_updates").update(updates).eq("id",wu_id).execute()
            else:
                updates["project_id"] = pid
                updates["week_number"] = sel_week
                supabase.table("fi_weekly_updates").insert(updates).execute()
            st.rerun()

        # ── A: Step Progress ──
        with st.expander("📊 Step Progress", expanded=True):
            if not steps:
                st.info("No steps defined yet. Add them in Project Setup → Master Plan.")
            else:
                step_progress = wu.get("step_progress") or []
                if isinstance(step_progress, str):
                    try: step_progress = json.loads(step_progress)
                    except: step_progress = []
                sp_by_id = {sp["step_id"]: sp for sp in step_progress if isinstance(sp,dict)}
                PCT_OPTS = ["0%","25%","50%","75%","100%"]
                STAT_OPTS = ["Not Started","In Progress","Completed"]

                # Only show steps that have started by selected week
                _due_steps = [s for s in steps if s.get("planned_start_week",1) <= sel_week]
                _future_steps = [s for s in steps if s.get("planned_start_week",1) > sel_week]
                if _future_steps:
                    st.caption(f"⏳ {len(_future_steps)} step(s) not yet due — will appear in later weeks.")
                if not _due_steps:
                    st.info(f"No steps are due yet in Week {sel_week}.")
                else:
                    with st.form(f"fi_step_progress_{sel_week}"):
                        new_sp = []
                        _chunks = [_due_steps[i:i+5] for i in range(0,len(_due_steps),5)]
                        for chunk in _chunks:
                            cols = st.columns(len(chunk))
                            for ci, step in enumerate(chunk):
                                sid = str(step["id"])
                                ex  = sp_by_id.get(sid,{})
                                with cols[ci]:
                                    st.markdown(f"**{step['step_name']}**")
                                    st.caption(f"W{step.get('planned_start_week','')}→W{step.get('planned_end_week','')} | {step.get('owner','') or '—'}")
                                    _pct_str = f"{int(ex.get('pct_complete',0))}%"
                                    _pct_str = _pct_str if _pct_str in PCT_OPTS else "0%"
                                    pct_sel = st.selectbox("Progress", PCT_OPTS,
                                        index=PCT_OPTS.index(_pct_str), key=f"sp_p_{sid}",
                                        label_visibility="collapsed")
                                    notes = st.text_input("Notes", value=ex.get("notes",""),
                                        key=f"sp_n_{sid}", placeholder="Notes…",
                                        label_visibility="collapsed")
                                    _pct_int = int(pct_sel.replace("%",""))
                                    _auto_status = "Completed" if _pct_int==100 else "In Progress" if _pct_int>0 else "Not Started"
                                    new_sp.append({
                                        "step_id": sid, "status": _auto_status,
                                        "pct_complete": _pct_int, "notes": notes
                                    })
                        if st.form_submit_button("💾 Save Step Progress", type="primary"):
                            _save_wu({"step_progress":json.dumps(new_sp), "updated_by":name})
                            st.rerun()

                # Gantt
                _wu_fresh = supabase.table("fi_weekly_updates").select("*").eq("project_id",pid).order("week_number").execute().data or []
                gf = _gantt_chart(steps, _wu_fresh, current_week)
                if gf: st.pyplot(gf); plt.close(gf)

        # ── B: KPI & KAI Update ──
        with st.expander("📈 KPI & KAI Update", expanded=True):
            if not kpi:
                st.info("No KPI defined yet. Add it in Project Setup → KPI & KAI.")
            else:
                # Parse KAIs from project setup
                _kai_subs = kpi.get("sub_components") or []
                if isinstance(_kai_subs, str):
                    try: _kai_subs = json.loads(_kai_subs)
                    except: _kai_subs = []
                _kai_subs = [s for s in _kai_subs if isinstance(s,dict)]

                # Load existing weekly KAI values
                _wu_kai = wu.get("kpi_notes") or ""
                try: _wu_kai_data = json.loads(_wu_kai) if _wu_kai.startswith("{") else {}
                except: _wu_kai_data = {}

                with st.form(f"fi_kpi_entry_{sel_week}"):
                    # KPI header
                    st.markdown(f"**🎯 {kpi.get('kpi_name','')}** — Baseline: {kpi.get('baseline_value','')} | Target: {kpi.get('target_value','')} {kpi.get('unit','')}")
                    ke1,ke2 = st.columns(2)
                    kpi_val = ke1.number_input(
                        f"This week's value ({kpi.get('unit','')})",
                        value=float(wu.get("kpi_value",0) or 0))
                    collected_by = ke2.multiselect("Collected by",
                        [m["member_name"] for m in team],
                        default=[x for x in (wu.get("kpi_collected_by","") or "").split(",")
                                 if x.strip() in [m["member_name"] for m in team]])
                    shifts = []  # removed

                    # KAI values — columnar, max 4 per row
                    kai_weekly = {}
                    if _kai_subs:
                        st.divider()
                        st.caption("↳ KAI updates this week:")
                        _kai_chunks = [_kai_subs[i:i+4] for i in range(0,len(_kai_subs),4)]
                        for chunk in _kai_chunks:
                            _kcols = st.columns(len(chunk))
                            for ci, kai in enumerate(chunk):
                                kn = kai.get("name","")
                                _ex_val = float(_wu_kai_data.get(kn, {}).get("value", 0))
                                with _kcols[ci]:
                                    st.markdown(f"**{kn}**")
                                    st.caption(f"Baseline: {kai.get('baseline','')} → Target: {kai.get('target','')} {kai.get('unit','')}")
                                    st.caption(f"Owner: {kai.get('owner','—')}")
                                    kai_weekly[kn] = {
                                        "value": st.number_input(
                                            f"{kn} value", value=_ex_val,
                                            key=f"kai_w_{sel_week}_{ci}",
                                            label_visibility="collapsed")
                                    }

                    if st.form_submit_button("💾 Save KPI & KAI Update", type="primary"):
                        _save_wu({
                            "kpi_value": kpi_val,
                            "kpi_collected_by": ",".join(collected_by),
                            "shifts_covered": ",".join(shifts),
                            "kpi_notes": json.dumps(kai_weekly),
                            "updated_by": name
                        })
                        st.rerun()

                # Trend chart
                trend_wu = supabase.table("fi_weekly_updates").select("*").eq("project_id",pid).order("week_number").execute().data or []
                tf = _kpi_trend_chart(kpi, kpi.get("baseline_value"), kpi.get("target_value"), trend_wu)
                st.pyplot(tf); plt.close(tf)

        # ── C: Root Cause Analysis (unlocks week 3) ──
        if sel_week >= 3:
            with st.expander(f"🔍 Root Cause Analysis {'(Available from Week 3)' if sel_week < 3 else ''}", expanded=False):
                with st.form(f"fi_rca_{sel_week}"):
                    rca_done = st.checkbox("Has root cause analysis been performed this week?",
                                           value=bool(wu.get("rca_performed")))
                    rca_method = rca_findings = rca_file = None
                    causes_verified = causes_method = None
                    if rca_done:
                        r1,r2 = st.columns(2)
                        rca_method  = r1.selectbox("Method used", RCA_METHODS,
                                        index=RCA_METHODS.index(wu.get("rca_method","5-Why")) if wu.get("rca_method") in RCA_METHODS else 0)
                        rca_file    = r2.file_uploader("Upload findings (optional)", type=["pdf","png","jpg"], key=f"rca_f_{sel_week}")
                        rca_findings= st.text_area("Describe findings *", value=wu.get("rca_findings",""), height=100)
                        causes_verified = st.checkbox("Have causes been verified with data?", value=bool(wu.get("causes_verified")))
                        if causes_verified:
                            causes_method = st.text_input("Describe verification method", value=wu.get("causes_verification_method",""))
                    if sel_week >= 7:
                        st.divider()
                        st.markdown("**Reoccurrence Analysis** (Week 7+)")
                        reoc_before = st.checkbox("Has a similar problem occurred before?", value=bool(wu.get("reoccurrence_before")))
                        reoc_desc   = st.text_area("Describe previous occurrence + what was done", value=wu.get("reoccurrence_description",""), height=60)
                        reoc_prev   = st.checkbox("Is reoccurrence prevention in place?", value=bool(wu.get("reoccurrence_prevention")))
                        single_pa   = st.checkbox("Single problem analysis applied?", value=bool(wu.get("single_problem_analysis")))
                        single_notes= st.text_input("Follow-up notes", value=wu.get("single_problem_notes",""))
                    else:
                        reoc_before=reoc_desc=reoc_prev=single_pa=single_notes = None
                    if st.form_submit_button("💾 Save RCA"):
                        updates = {"rca_performed":rca_done,"updated_by":name}
                        if rca_done:
                            updates.update({"rca_method":rca_method,"rca_findings":rca_findings,
                                            "causes_verified":causes_verified or False,
                                            "causes_verification_method":causes_method or ""})
                            if rca_file: updates["rca_file_b64"] = _to_b64(rca_file)
                        if sel_week >= 7 and reoc_desc is not None:
                            updates.update({"reoccurrence_before":reoc_before,"reoccurrence_description":reoc_desc,
                                            "reoccurrence_prevention":reoc_prev,"single_problem_analysis":single_pa,
                                            "single_problem_notes":single_notes or ""})
                        _save_wu(updates)

        # ── D: Action Plan (unlocks week 4) ──
        if sel_week >= 4:
            with st.expander("✅ Action Plan", expanded=True):
                if actions:
                    df_act = pd.DataFrame(actions)[["description","root_cause_addressed","owner","target_date","status"]]
                    df_act.columns = ["Action","Root Cause Addressed","Owner","Target Date","Status"]
                    st.dataframe(df_act.style.apply(
                        lambda row: ["background-color:#ffcccc" if row["Status"]=="Overdue"
                                     else "background-color:#ccffcc" if row["Status"]=="Completed"
                                     else "" for _ in row], axis=1),
                        use_container_width=True, hide_index=True)
                    # Update status
                    if can_create:
                        act_sel = st.selectbox("Update action status",
                                               ["—"]+[a["description"][:50] for a in actions], key="fi_act_update")
                        if act_sel != "—":
                            act_obj = next((a for a in actions if a["description"][:50]==act_sel), None)
                            if act_obj:
                                nc1,nc2,nc3 = st.columns(3)
                                new_status = nc1.selectbox("New Status", ACTION_STATUSES,
                                    index=ACTION_STATUSES.index(act_obj.get("status","Open")), key="fi_new_status")
                                new_ev = nc2.file_uploader("Upload Evidence", type=["pdf","png","jpg","xlsx"], key="fi_ev_upload")
                                if nc3.button("💾 Update Action", key="fi_update_act"):
                                    upd = {"status":new_status}
                                    if new_status == "Completed": upd["completed_date"] = str(date.today())
                                    if new_ev: upd["evidence_b64"] = _to_b64(new_ev); upd["evidence_filename"] = new_ev.name
                                    supabase.table("fi_actions").update(upd).eq("id",act_obj["id"]).execute()
                                    st.rerun()

                with st.form(f"fi_add_action_{sel_week}"):
                    st.caption("➕ Add a new action")
                    ac1,ac2 = st.columns(2)
                    act_desc  = ac1.text_area("Action Description *", height=70)
                    act_rc    = ac2.text_area("Root Cause Addressed", height=70)
                    ac3,ac4,ac5 = st.columns(3)
                    act_owner = ac3.selectbox("Owner", [""]+[m["member_name"] for m in team], key=f"act_owner_{sel_week}")
                    act_date  = ac4.date_input("Target Date", value=date.today()+timedelta(weeks=2))
                    act_ev    = ac5.file_uploader("Evidence (optional)", type=["pdf","png","jpg"], key=f"act_ev_{sel_week}")
                    if st.form_submit_button("➕ Add Action"):
                        if act_desc:
                            supabase.table("fi_actions").insert({
                                "project_id":pid,"description":act_desc,"root_cause_addressed":act_rc,
                                "owner":act_owner,"target_date":str(act_date),"status":"Open",
                                "created_week":sel_week,
                                "evidence_b64": _to_b64(act_ev) if act_ev else None,
                                "evidence_filename": act_ev.name if act_ev else None,
                            }).execute()
                            st.rerun()

        # ── E: Basic Conditions (week 4+) ──
        if sel_week >= 4:
            with st.expander("🔧 Basic Conditions", expanded=False):
                with st.form(f"fi_basic_{sel_week}"):
                    bc_done = st.checkbox("Have critical areas been restored to basic conditions?",
                                           value=bool(wu.get("basic_conditions_restored")))
                    if bc_done:
                        bc_desc  = st.text_area("Describe what was restored", value=wu.get("basic_conditions_description",""), height=60)
                        bc_b1,bc_b2 = st.columns(2)
                        bc_before = bc_b1.file_uploader("Before photo", type=["jpg","jpeg","png"], key=f"bc_before_{sel_week}")
                        bc_after  = bc_b2.file_uploader("After photo",  type=["jpg","jpeg","png"], key=f"bc_after_{sel_week}")
                        bc_date   = st.date_input("Date completed", value=date.today())
                    else:
                        bc_desc=bc_before=bc_after=bc_date = None
                    if st.form_submit_button("💾 Save Basic Conditions"):
                        upd = {"basic_conditions_restored":bc_done,"updated_by":name}
                        if bc_done:
                            upd["basic_conditions_description"] = bc_desc or ""
                            upd["basic_conditions_date"] = str(bc_date) if bc_date else None
                            if bc_before: upd["basic_conditions_before_b64"] = _to_b64(bc_before)
                            if bc_after:  upd["basic_conditions_after_b64"]  = _to_b64(bc_after)
                        _save_wu(upd)

        # ── F: Meeting Log ──
        with st.expander("📋 Team Meeting Log", expanded=False):
            with st.form(f"fi_meeting_{sel_week}"):
                mt1,mt2 = st.columns(2)
                mtg_held  = mt1.checkbox("Meeting held this week?", value=bool(wu.get("meeting_held")))
                attendees = mt2.multiselect("Attendees", [m["member_name"] for m in team],
                    default=[x.strip() for x in (wu.get("meeting_attendees","") or "").split(",") if x.strip() in [m["member_name"] for m in team]])
                mtg_notes = st.text_area("Meeting notes / summary", value=wu.get("meeting_notes",""), height=60)
                if team:
                    att_rate = len(attendees)/len(team)*100
                    st.info(f"Attendance rate: **{att_rate:.0f}%** ({len(attendees)}/{len(team)} members)")
                if st.form_submit_button("💾 Save Meeting Log"):
                    _save_wu({"meeting_held":mtg_held,"meeting_attendees":",".join(attendees),"meeting_notes":mtg_notes,"updated_by":name})

    # ════════════════════════════════════════
    # SUB-TAB 3 — RESULTS TRACKER
    # ════════════════════════════════════════
    with subtabs[2]:
        st.markdown("#### 📈 Results Tracker")
        if not kpi:
            st.info("No KPI defined yet.")
        else:
            kpi_vals_list = [wu["kpi_value"] for wu in sorted(wu_rows, key=lambda x:x["week_number"]) if wu.get("kpi_value") is not None]
            baseline = kpi.get("baseline_value") or 0
            target_v = kpi.get("target_value") or 0

            # Progress indicator
            if kpi_vals_list:
                current_val = kpi_vals_list[-1]
                total_gap   = abs(target_v - baseline)
                progress    = abs(current_val - baseline) / total_gap * 100 if total_gap > 0 else 0
                progress    = min(100, progress)
                rc1,rc2,rc3,rc4 = st.columns(4)
                rc1.metric("Baseline",     f"{baseline} {kpi.get('unit','')}")
                rc2.metric("Current",      f"{current_val} {kpi.get('unit','')}")
                rc3.metric("Target",       f"{target_v} {kpi.get('unit','')}")
                rc4.metric("Progress",     f"{progress:.0f}%",
                           delta=f"{progress:.0f}% toward target",
                           delta_color="normal" if progress > 0 else "off")
                # Projected completion
                if len(kpi_vals_list) >= 2:
                    rate = (kpi_vals_list[-1] - kpi_vals_list[0]) / len(kpi_vals_list)
                    if rate != 0:
                        weeks_left = (target_v - kpi_vals_list[-1]) / rate
                        proj_week  = current_week + weeks_left
                        st.info(f"📐 Projected completion: **Week {proj_week:.0f}** (linear extrapolation)")

            # KPI trend chart
            tf2 = _kpi_trend_chart(kpi, baseline, target_v, wu_rows)
            st.pyplot(tf2); plt.close(tf2)

            # Cost/benefit tracker
            if cost:
                st.divider()
                st.markdown("#### 💰 Cost / Benefit Tracker")
                cb_c1,cb_c2,cb_c3 = st.columns(3)
                cb_c1.metric("Estimated Cost",         f"{cost.get('estimated_cost',0):.1f} K€")
                cb_c2.metric("Estimated Annual Savings",f"{cost.get('estimated_savings',0):.1f} K€/yr")
                cb_c3.metric("Payback Period",          f"{cost.get('payback_months',0):.1f} months")

    # ════════════════════════════════════════
    # SUB-TAB 4 — STABILISATION
    # ════════════════════════════════════════
    with subtabs[3]:
        st.markdown("#### 🔒 Stabilisation")
        st.caption("Sections unlock progressively. Ensures gains are held and documented.")
        stab_vals = stab or {}

        # ── Procedures (week 10+) ──
        unlock_note = lambda w: f"*(Unlocks Week {w})*" if current_week < w else ""
        with st.expander(f"📄 Standards & Procedures {unlock_note(10)}", expanded=current_week>=10):
            with st.form("fi_stab_procedures"):
                proc_done = st.checkbox("Have procedures been created to hold the gains?",
                                         value=bool(stab_vals.get("procedures_created")),
                                         disabled=current_week<10)
                procedures = []
                if proc_done:
                    n_procs = st.number_input("Number of procedures", min_value=1, max_value=10, value=max(1,len(stab_vals.get("procedures") or [])))
                    existing_procs = stab_vals.get("procedures") or []
                    if isinstance(existing_procs, str):
                        try: existing_procs = json.loads(existing_procs)
                        except: existing_procs = []
                    for pi in range(int(n_procs)):
                        ep = existing_procs[pi] if pi < len(existing_procs) else {}
                        pc1,pc2 = st.columns(2)
                        procedures.append({
                            "name": pc1.text_input(f"Procedure {pi+1} Name", value=ep.get("name",""), key=f"proc_n_{pi}"),
                            "description": pc2.text_input(f"Description", value=ep.get("description",""), key=f"proc_d_{pi}"),
                        })
                if st.form_submit_button("💾 Save Procedures"):
                    stab_data = {"procedures_created":proc_done,"procedures":json.dumps(procedures)}
                    if stab: supabase.table("fi_stabilisation").update(stab_data).eq("id",stab["id"]).execute()
                    else: supabase.table("fi_stabilisation").insert({"project_id":pid,**stab_data}).execute()
                    st.rerun()

        # ── CIL Standards (week 5+) ──
        with st.expander(f"🔍 CIL Standards {unlock_note(5)}", expanded=current_week>=5):
            with st.form("fi_stab_cil"):
                cil_def  = st.checkbox("Are CIL standards defined for critical areas?",
                                        value=bool(stab_vals.get("cil_standards_defined")), disabled=current_week<5)
                cil_score= st.number_input("Latest CIL Audit Score (%)", min_value=0.0, max_value=100.0,
                                            value=float(stab_vals.get("cil_audit_score",0) or 0))
                cil_file = st.file_uploader("Upload CIL audit record", type=["pdf","xlsx","png"], key="fi_cil_file")
                if cil_score >= 90:
                    st.success("✅ CIL score meets ≥90% target")
                elif cil_score > 0:
                    st.warning(f"⚠️ CIL score {cil_score:.0f}% — target is ≥90%")
                if st.form_submit_button("💾 Save CIL"):
                    cil_data = {"cil_standards_defined":cil_def,"cil_audit_score":cil_score}
                    if cil_file: cil_data["cil_file_b64"] = _to_b64(cil_file)
                    if stab: supabase.table("fi_stabilisation").update(cil_data).eq("id",stab["id"]).execute()
                    else: supabase.table("fi_stabilisation").insert({"project_id":pid,**cil_data}).execute()
                    st.rerun()

        # ── Monitoring (week 7+) ──
        with st.expander(f"📊 Monitoring Systems {unlock_note(7)}", expanded=current_week>=7):
            with st.form("fi_stab_monitoring"):
                mon_in    = st.checkbox("Are monitoring systems in place?",
                                         value=bool(stab_vals.get("monitoring_in_place")), disabled=current_week<7)
                mon_types = st.multiselect("System types", MONITORING_TYPES,
                                            default=[x.strip() for x in (stab_vals.get("monitoring_types","") or "").split(",") if x.strip() in MONITORING_TYPES])
                mon_active= st.checkbox("Are they being actively used?", value=bool(stab_vals.get("monitoring_active")))
                mon_date  = st.date_input("Last update date",
                                           value=date.fromisoformat(str(stab_vals["monitoring_last_update"])) if stab_vals.get("monitoring_last_update") else date.today())
                mon_ev    = st.file_uploader("Upload evidence", type=["pdf","png","jpg"], key="fi_mon_ev")
                if st.form_submit_button("💾 Save Monitoring"):
                    mon_data = {"monitoring_in_place":mon_in,"monitoring_types":",".join(mon_types),
                                "monitoring_active":mon_active,"monitoring_last_update":str(mon_date)}
                    if mon_ev: mon_data["monitoring_evidence_b64"] = _to_b64(mon_ev)
                    if stab: supabase.table("fi_stabilisation").update(mon_data).eq("id",stab["id"]).execute()
                    else: supabase.table("fi_stabilisation").insert({"project_id":pid,**mon_data}).execute()
                    st.rerun()

        # ── OPLs & Training (week 10+) ──
        with st.expander(f"📝 OPLs & Training {unlock_note(10)}", expanded=current_week>=10):
            with st.form("fi_stab_opls"):
                existing_opls = stab_vals.get("opls") or []
                if isinstance(existing_opls, str):
                    try: existing_opls = json.loads(existing_opls)
                    except: existing_opls = []
                n_opls = st.number_input("Number of OPLs/SOPs created", min_value=0, max_value=20, value=max(0,len(existing_opls)))
                new_opls = []
                for oi in range(int(n_opls)):
                    eo = existing_opls[oi] if oi < len(existing_opls) else {}
                    oc1,oc2,oc3 = st.columns(3)
                    new_opls.append({
                        "title":    oc1.text_input(f"OPL/SOP {oi+1} Title", value=eo.get("title",""), key=f"opl_t_{oi}"),
                        "covers":   oc2.text_input(f"Improvement it covers", value=eo.get("covers",""), key=f"opl_c_{oi}"),
                        "created_by": oc3.selectbox("Created by", [""]+[m["member_name"] for m in team], key=f"opl_cb_{oi}"),
                    })
                st.divider()
                st.markdown("**Training Matrix**")
                existing_tm = stab_vals.get("training_matrix") or []
                if isinstance(existing_tm, str):
                    try: existing_tm = json.loads(existing_tm)
                    except: existing_tm = []
                new_tm = []
                for mi, member in enumerate(team):
                    mname = member["member_name"]
                    em = next((x for x in existing_tm if x.get("member")==mname), {})
                    tc1,tc2,tc3 = st.columns(3)
                    new_tm.append({
                        "member": mname,
                        "opls_trained": tc1.multiselect(f"{mname} — OPLs trained on",
                                                         [o["title"] for o in new_opls if o["title"]],
                                                         default=[x for x in (em.get("opls_trained") or []) if x in [o["title"] for o in new_opls]],
                                                         key=f"tm_opl_{mi}"),
                        "training_date": str(tc2.date_input("Training date", value=date.today(), key=f"tm_date_{mi}")),
                        "plan_confirmed": tc3.checkbox("Training planned?", value=bool(em.get("plan_confirmed")), key=f"tm_plan_{mi}"),
                    })
                if st.form_submit_button("💾 Save OPLs & Training"):
                    opl_data = {"opls":json.dumps(new_opls),"training_matrix":json.dumps(new_tm)}
                    if stab: supabase.table("fi_stabilisation").update(opl_data).eq("id",stab["id"]).execute()
                    else: supabase.table("fi_stabilisation").insert({"project_id":pid,**opl_data}).execute()
                    st.rerun()

        # ── Workplace / 5S (week 5+) ──
        with st.expander(f"🏭 Workplace & Visual Evidence {unlock_note(5)}", expanded=current_week>=5):
            with st.form("fi_stab_5s"):
                five_s = st.slider("5S Rating", 1, 5, int(stab_vals.get("five_s_rating",1) or 1))
                five_s_ph = st.file_uploader("Upload 5S photos", type=["jpg","jpeg","png"], key="fi_5s_photo")
                five_s_notes = st.text_area("5S notes", value=stab_vals.get("five_s_notes",""), height=60)
                imp_vis = st.checkbox("Are improvements on the machine/area evident?",
                                       value=bool(stab_vals.get("improvements_visible")))
                imp_ph = st.file_uploader("Upload improvement photos", type=["jpg","jpeg","png"], key="fi_imp_photo")
                if st.form_submit_button("💾 Save Workplace Evidence"):
                    ws_data = {"five_s_rating":five_s,"five_s_notes":five_s_notes,"improvements_visible":imp_vis}
                    if five_s_ph: ws_data["five_s_photos_b64"] = _to_b64(five_s_ph)
                    if imp_ph: ws_data["improvements_photos_b64"] = _to_b64(imp_ph)
                    if stab: supabase.table("fi_stabilisation").update(ws_data).eq("id",stab["id"]).execute()
                    else: supabase.table("fi_stabilisation").insert({"project_id":pid,**ws_data}).execute()
                    st.rerun()

    # ════════════════════════════════════════
    # SUB-TAB 5 — AUDIT VIEW (auditor only)
    # ════════════════════════════════════════
    if is_auditor and len(subtabs) > 4:
        with subtabs[4]:
            st.markdown("#### 🔍 Audit View")
            st.caption("Visible to Plant Manager and Pillar Leader only. The team never sees this.")

            # Refresh data for scoring
            _stab_fresh_rows = supabase.table("fi_stabilisation").select("*").eq("project_id",pid).execute().data or []
            stab_fresh = _stab_fresh_rows[0] if _stab_fresh_rows else {}
            try:
                scores, total = _score_project(selected_project, team, kpi, steps, wu_rows, actions, stab_fresh, audit_records)
            except Exception as _score_err:
                st.error(f"Scoring error: {_score_err}")
                scores = {q: 0 for q in QUESTIONS}
                total = 0

            # Overall score
            score_color = "green" if total >= 70 else "orange" if total >= 45 else "red"
            st.markdown(f"## Total Score: :{score_color}[{int(total)} / 100]")

            # Target ramp
            target_this_week = TARGET_RAMP.get(current_week, 100)
            gap = total - target_this_week
            if gap >= 0:
                st.success(f"✅ On track — {int(total)}/100 vs target {target_this_week}/100 (+{gap:.0f} pts ahead)")
            else:
                st.warning(f"⚠️ Behind target — {int(total)}/100 vs target {target_this_week}/100 ({gap:.0f} pts behind)")

            # Score trajectory chart
            traj_weeks = sorted(TARGET_RAMP.keys())
            traj_targets = [TARGET_RAMP[w] for w in traj_weeks]
            _audit_recs_safe = audit_records or []
            audit_scores_by_week = {ar["week_number"]: ar.get("total_score",0) for ar in _audit_recs_safe}
            actual_weeks = sorted(audit_scores_by_week.keys())
            actual_scores = [audit_scores_by_week[w] for w in actual_weeks]
            fig_traj, ax_traj = plt.subplots(figsize=(10,4), dpi=100)
            ax_traj.plot(traj_weeks, traj_targets, "o--", color="#888888", label="Target ramp", linewidth=1.5)
            if actual_weeks:
                ax_traj.plot(actual_weeks, actual_scores, "o-", color="#006394", label="Actual score", linewidth=2)
            ax_traj.axvline(current_week, color="#DE201B", linewidth=1.5, linestyle=":", label=f"Week {current_week}")
            ax_traj.set_xticks(range(1,13)); ax_traj.set_xticklabels([f"W{i}" for i in range(1,13)])
            ax_traj.set_ylabel("Score / 100"); ax_traj.set_ylim(0,105)
            ax_traj.spines["top"].set_visible(False); ax_traj.spines["right"].set_visible(False)
            ax_traj.legend(fontsize=9, frameon=False)
            plt.tight_layout()
            st.pyplot(fig_traj); plt.close(fig_traj)

            # Traffic lights
            st.markdown("#### 🚦 Dimension Traffic Lights")
            dimensions = {
                "Involvement":   [1,2,34,35,36],
                "Method":        [8,9,10,11,12,13,14,15,16],
                "Action Plan":   [17,18,19,20,21,22,23],
                "Results":       [3,4,5,6,7,24,25],
                "Stabilisation": [26,27,28,29,30,31,32,33],
            }
            dim_cols = st.columns(len(dimensions))
            for ci, (dim, qs) in enumerate(dimensions.items()):
                achieved = sum(scores.get(q,0) for q in qs)
                possible = sum(QUESTIONS[q]["score"] for q in qs)
                pct = achieved/possible*100 if possible > 0 else 0
                color = "🟢" if pct >= 90 else "🟡" if pct >= 70 else "🔴"
                dim_cols[ci].metric(f"{color} {dim}", f"{achieved}/{possible}", f"{pct:.0f}%")

            # Per-question table
            st.divider()
            st.markdown("#### 📋 Per-Question Status")
            q_data = []
            for qn, qdata in QUESTIONS.items():
                achieved = scores.get(qn,0)
                due = current_week >= qdata["week"]
                status = ("✅ Met" if achieved >= qdata["score"] else
                          "⚠️ Not Met" if due else "⏳ Not Due Yet")
                q_data.append({
                    "Q#": qn, "Question": qdata["text"][:60],
                    "Weight": qdata["score"], "Due Week": qdata["week"],
                    "Status": status, "Score": f"{achieved}/{qdata['score']}"
                })
            df_q = pd.DataFrame(q_data)
            st.dataframe(df_q.style.apply(
                lambda row: ["background-color:#d4edda" if "✅" in str(row["Status"])
                             else "background-color:#fff3cd" if "⏳" in str(row["Status"])
                             else "background-color:#f8d7da" for _ in row], axis=1),
                use_container_width=True, hide_index=True)

            # Gap actions
            gaps = [(qn, QUESTIONS[qn]) for qn, qdata in QUESTIONS.items()
                    if scores.get(qn,0) < qdata["score"] and current_week >= qdata["week"]]
            if gaps:
                st.divider()
                st.markdown("#### 🚨 Gap Register — Actions Required")
                for qn, qdata in gaps:
                    st.markdown(f"**Q{qn}** ({qdata['score']} pts) — {qdata['text']}")

            # Team understanding check
            st.divider()
            st.markdown("#### 👤 Team Understanding Check")
            with st.form("fi_audit_understanding"):
                au1,au2 = st.columns(2)
                tested_member = au1.text_input("Team member tested (random pick)")
                understanding_pass = au2.checkbox("Did they explain correctly?")
                audit_notes = st.text_area("Auditor notes", height=80)
                if st.form_submit_button("💾 Save Audit Record"):
                    q_scores = {str(qn): scores.get(qn,0) for qn in QUESTIONS}
                    if understanding_pass:
                        q_scores["34"] = True; q_scores["35"] = True
                    ar_data = {
                        "project_id":pid,"week_number":current_week,
                        "question_scores":json.dumps(q_scores),
                        "total_score":total,
                        "team_understanding_tested":bool(tested_member),
                        "member_tested":tested_member,
                        "understanding_pass":understanding_pass,
                        "auditor_notes":audit_notes,
                        "audited_by":name,
                    }
                    existing_ar = next((ar for ar in audit_records if ar["week_number"]==current_week), None)
                    if existing_ar:
                        supabase.table("fi_audit_records").update(ar_data).eq("id",existing_ar["id"]).execute()
                    else:
                        supabase.table("fi_audit_records").insert(ar_data).execute()
                    st.success("✅ Audit record saved"); st.rerun()

            # PDF Export
            st.divider()
            st.markdown("#### 📄 Generate PDF Report")
            if st.button("📥 Generate Full Project Report PDF", type="primary", key="fi_pdf_gen"):
                with st.spinner("Generating PDF..."):
                    stab_for_pdf = supabase.table("fi_stabilisation").select("*").eq("project_id",pid).execute().data
                    stab_for_pdf = stab_for_pdf[0] if stab_for_pdf else {}
                    pdf_buf = _generate_project_pdf(
                        selected_project, team, kpi, steps,
                        wu_rows, actions, stab_for_pdf,
                        audit_records, scores, total
                    )
                st.download_button("📥 Download PDF",
                    data=pdf_buf,
                    file_name=f"FI_Project_{selected_project['project_name'].replace(' ','_')}.pdf",
                    mime="application/pdf",
                    key="fi_pdf_download")
