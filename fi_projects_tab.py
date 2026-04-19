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
    1:12, 2:15, 3:24, 4:33, 5:56, 6:58,
    7:77, 8:79, 9:79, 10:85, 11:100, 12:100
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
         "LM/Min","BD Time (hrs)","Hours","Mins","Secs","K SAR","Count","Score"]

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
# BEAUTIFUL PDF REPORT GENERATOR
# ════════════════════════════════════════
def _generate_project_pdf(project, team, kpi, steps, weekly_updates, actions, stab, audit_records, scores, total_score):
    """QM-style 960x540 landscape slides report for FI projects."""
    import base64 as _b64, tempfile as _tf
    from reportlab.pdfgen import canvas as _rl_canvas
    from reportlab.lib.utils import ImageReader
    from reportlab.lib.colors import HexColor
    import matplotlib.patches as mpatches

    SW, SH = 960, 540   # slide dimensions (points)

    # ── Napco logo ──
    _logo_b64 = "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAIBAQEBAQIBAQECAgICAgQDAgICAgUEBAMEBgUGBgYFBgYGBwkIBgcJBwYGCAsICQoKCgoKBggLDAsKDAkKCgr/2wBDAQICAgICAgUDAwUKBwYHCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgr/wAARCAF/BVYDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9+KKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigBB7jFLnsazvEXiPSvCuhX3ijXr+O1sdNtJLm+upThYoY1LO59goJ/CvxZ/bG/4OE/2nPGnxFu9I/ZI1Oz8G+FrG4aOy1GfSbe7vtQUHHmuLlJI41PUIq7gOrZ4HsZPkWYZ5VlDDJWju27JX26M+fz/iXLOHaMZ4pu8top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iRo3yPStf7Wy/CK+Bw6Uv5pvma9FZJW6PUy/sTNMc/8AhRxLcf5aacE/V3bd+qTPkP8AZZ/4IrfsW/s1SQeItU8GHxx4ijbe2seLUW4VZP7yQY8tTnnJDEHnNfXVtClvCsUcQRVGFVegHoKkGDz/AFpa8bF43F46pz15uT83+Xb5Hu4LLsFl1Pkw8FFeS/Pq35sKKKK5jtCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigVkgooooGFFFFABgelFFFArIKKKKBhRRRQAYHpRRRQCSQUUUUBZBXwv/wcOf8AKPK4/wCxz0v+ctfdFfC//Bw5/wAo8rj/ALHPS/5y17HD3/I8w/8Ajj+Z8/xR/wAk5iv+vcvyZ+Dh+6K/Zj/g2X/5N/8AiN/2N0H/AKTivxnP3RX7Mf8ABsv/AMm//Eb/ALG6D/0nFfr/AB3/AMk5P1j/AOlI/CPDL/kqYf4ZfkfpnRRRX4Qf0vZBRRRQAUUUUAFFFFABRRRQAYHpRRRQKyDA9KKKKB2QUUUUAFFFFAWQUUUUAFFFFABRRRQAUUUUAFFFFABgelFFFArIKKKKBhRRRQAYA6CjA9KKKBWQUUUUDCiiigAooooCyCiiigLIKKKKACiiigApCiHqg/KlooE0nuFBAPUUUUDshNq+gpaKKBWSCiiigYUUUUAGAOgooooCyQUUUUAFFFFABgelFFFAWSCiiigAooooAKKKKAsgooooAKKKKADA9KKKKAskFGB1xRRQKyCiiigYUUUUAIVU9QD9aAqrwoA+lLRQKyuFFFFAwooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAQ98da+F/8Ag4ebH/BPG4GP+Zz0z/2rX3RjJ5HTpXkv7Zv7IHw8/be+DMnwQ+Jus6nY6ZJqdvfNPpMiLL5kW7aMurDHzHPFehlOJp4PM6Ver8MZJv0TPKzvCVcwyivh6XxTi0r6K7VkfzHjBUtiv2X/AODZdgv7PvxGz/0N0H/pNXRj/g20/Y4xg/Ezx1j/AK/rf/41X01+wr+wB8Kv2BPB2teCvhR4h1rULbW9SW9upNYljZlkVNgC7FXjFfoHFHFmU5tlMsNh3LmbT1Vtnc/L+DOCc6yPO44rEqPIk1o7vVaaWPd6KKK/Lz9lCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA//9k="
    _napco_logo_reader = None
    try:
        _logo_bytes = _b64.b64decode(_logo_b64)
        _tmp = _tf.NamedTemporaryFile(delete=False, suffix=".jpeg")
        _tmp.write(_logo_bytes); _tmp.flush()
        _napco_logo_reader = ImageReader(_tmp.name)
    except Exception:
        _napco_logo_reader = None

    # ── Colour palette (matches QM) ──
    DBLUE  = "#0C5595"
    RED    = "#DE201B"
    LGREY  = "#F4F6F8"
    MGREY  = "#BDC3C7"
    DGREY  = "#566573"
    GREEN  = "#1E8449"
    AMBER  = "#D68910"
    WHITE  = "#FFFFFF"
    BLACK  = "#1A1A2E"

    # ── Parse data ──
    kpi  = kpi or {}; stab = stab or {}; team = team or []
    steps= steps or []; actions = actions or []; weekly_updates = weekly_updates or []
    audit_records = audit_records or []
    launch_str = str(project.get("launch_date",""))[:10]
    try: launch_d = date.fromisoformat(launch_str)
    except: launch_d = date.today()
    cw = _get_current_week(project.get("launch_date"))
    kpi_baseline = float(kpi.get("baseline_value",0) or 0)
    kpi_target   = float(kpi.get("target_value",0)   or 0)
    kpi_vals     = sorted([w for w in weekly_updates if w.get("kpi_value") is not None], key=lambda x:x["week_number"])
    kpi_current  = float(kpi_vals[-1]["kpi_value"]) if kpi_vals else kpi_baseline
    kpi_progress = abs(kpi_current-kpi_baseline)/abs(kpi_target-kpi_baseline)*100 if abs(kpi_target-kpi_baseline)>0 else 0
    sub_comps = kpi.get("sub_components") or []
    if isinstance(sub_comps, str):
        try: sub_comps = json.loads(sub_comps)
        except: sub_comps = []
    sub_comps = [s for s in sub_comps if isinstance(s,dict)]
    completed_cnt = sum(1 for a in actions if a.get("status")=="Completed")
    in_prog_cnt   = sum(1 for a in actions if a.get("status")=="In Progress")
    open_cnt      = len(actions)-completed_cnt-in_prog_cnt
    ot_rate       = int(completed_cnt/len(actions)*100) if actions else 0
    TARGET_RAMP_L = {1:12,2:15,3:20,4:27,5:45,6:47,7:62,8:64,9:64,10:70,11:85,12:100}
    score_col = GREEN if total_score>=70 else AMBER if total_score>=45 else RED

    def _reader(fig):
        buf2=io.BytesIO(); fig.savefig(buf2,format="png",dpi=180,bbox_inches="tight",facecolor=fig.get_facecolor()); buf2.seek(0); plt.close(fig); return ImageReader(buf2)

    def draw_logo(c):
        if _napco_logo_reader:
            try: c.drawImage(_napco_logo_reader, 18, SH-115, width=260, height=100, preserveAspectRatio=True, mask="auto")
            except: pass

    def draw_sep(c, y):
        c.setFillColor(HexColor(RED));   c.rect(40, y, 110, 4, fill=1, stroke=0)
        c.setFillColor(HexColor(DBLUE)); c.rect(155, y, SW-195, 4, fill=1, stroke=0)

    def title_bar(c, text):
        c.setFillColor(HexColor(DBLUE)); c.rect(0, SH-50, SW, 50, fill=1, stroke=0)
        c.setFillColor(HexColor(WHITE)); c.setFont("Helvetica-Bold", 18)
        c.drawString(20, SH-35, text)

    def cover_slide(c):
        c.setFillColor(HexColor(WHITE)); c.rect(0,0,SW,SH,fill=1,stroke=0)
        draw_logo(c)
        draw_sep(c, SH*0.58)
        # Title
        c.setFillColor(HexColor(DBLUE)); c.setFont("Helvetica-Bold", 28)
        c.drawString(40, SH-160, project.get("project_name","")[:55])
        # Subtitle line
        c.setFillColor(HexColor(DGREY)); c.setFont("Helvetica", 13)
        c.drawString(40, SH-190, f"Focused Improvement  ·  {project.get('target_area','')}  ·  Week {cw} of 12")
        # Score badge
        col = score_col
        c.setFillColor(HexColor(col)); c.circle(SW-80, SH-100, 55, fill=1, stroke=0)
        c.setFillColor(HexColor(WHITE)); c.setFont("Helvetica-Bold", 32); c.drawCentredString(SW-80, SH-112, str(int(total_score)))
        c.setFont("Helvetica-Bold", 11); c.drawCentredString(SW-80, SH-130, "/ 100")
        c.setFont("Helvetica", 9); c.drawCentredString(SW-80, SH-145, "AUDIT SCORE")
        # KPI badges row
        bw = 130; bh = 65; bx = 40; by = 20
        for lbl, val, sub, col, light in [
            ("Baseline",  f"{kpi_baseline}{kpi.get('unit','')}", "Start", DBLUE, True),
            ("Current",   f"{kpi_current:.1f}{kpi.get('unit','')}", f"W{cw}", DBLUE, False),
            ("Target",    f"{kpi_target}{kpi.get('unit','')}", "Goal",  GREEN,  False),
            ("Progress",  f"{kpi_progress:.0f}%", "toward target", AMBER, False),
        ]:
            bg = HexColor(LGREY) if light else HexColor(col)
            c.setFillColor(bg); c.roundRect(bx, by, bw, bh, 5, fill=1, stroke=0)
            tc = HexColor(DBLUE) if light else HexColor(WHITE)
            c.setFillColor(tc); c.setFont("Helvetica", 7); c.drawCentredString(bx+bw/2, by+bh-12, lbl.upper())
            c.setFont("Helvetica-Bold", 17 if len(str(val))<=6 else 13); c.drawCentredString(bx+bw/2, by+18, str(val))
            if sub:
                c.setFont("Helvetica", 7); c.setFillColor(HexColor(MGREY) if not light else HexColor(DGREY))
                c.drawCentredString(bx+bw/2, by+6, sub)
            bx += bw + 8
        # Problem statement
        c.setFillColor(HexColor(DGREY)); c.setFont("Helvetica", 10)
        pb = project.get("problem_statement","")
        words = pb.split(); line=""; py = SH*0.52
        for w in words:
            test=(line+" "+w).strip()
            if c.stringWidth(test,"Helvetica",10)<SW-520: line=test
            else:
                c.drawString(40, py, line); py-=14; line=w
        if line: c.drawString(40, py, line)
        # Generated date
        c.setFillColor(HexColor(MGREY)); c.setFont("Helvetica-Oblique", 9)
        c.drawRightString(SW-20, 8, f"Generated {date.today().strftime('%d %B %Y')}")

    def content_slide_header(c, title):
        c.setFillColor(HexColor(WHITE)); c.rect(0,0,SW,SH,fill=1,stroke=0)
        title_bar(c, title)
        # Logo top-right on content slides (doesn't overlap title)
        if _napco_logo_reader:
            try: c.drawImage(_napco_logo_reader, SW-220, SH-110, width=200, height=90, preserveAspectRatio=True, mask="auto")
            except: pass
        draw_sep(c, SH-125)

    def gantt_fig():
        n=len(steps)
        if n==0: return None
        fig,ax=plt.subplots(figsize=(11,max(2,n*0.55+0.8)),dpi=180)
        fig.patch.set_facecolor("#FAFAFA"); ax.set_facecolor("#FAFAFA")
        sp_map={}
        for wu in weekly_updates:
            sp_raw=wu.get("step_progress") or []
            if isinstance(sp_raw,str):
                try: sp_raw=json.loads(sp_raw)
                except: sp_raw=[]
            for sp in sp_raw:
                if isinstance(sp,dict):
                    sid=sp.get("step_id",""); sp_map[sid]=max(sp_map.get(sid,0),sp.get("pct_complete",0))
        for i,step in enumerate(steps):
            ps=step.get("planned_start_week",1); pe=step.get("planned_end_week",2)
            ax.barh(i,pe-ps,left=ps-1,height=0.5,color="#C8DFF0",zorder=2)
            ax.barh(i,pe-ps,left=ps-1,height=0.5,color="none",edgecolor=DBLUE,lw=0.8,zorder=3)
            pct=sp_map.get(str(step.get("id","")),0)
            if pct>0:
                col=GREEN if pct==100 else DBLUE
                ax.barh(i,(pe-ps)*pct/100,left=ps-1,height=0.5,color=col,alpha=0.85,zorder=4)
                if pct>=20: ax.text(ps-1+((pe-ps)*pct/100)/2,i,f"{pct}%",ha="center",va="center",fontsize=7,color="#FFFFFF",fontweight="bold",zorder=5)
        for w in range(13): ax.axvline(w,color="#EEEEEE",lw=0.5,zorder=1)
        ax.axvline(cw-1,color=RED,lw=2,ls="--",zorder=6,alpha=0.8)
        ax.text(cw-0.8,n-0.1,f"W{cw}",color=RED,fontsize=7,fontweight="bold")
        ax.set_yticks(range(n)); ax.set_yticklabels([s.get("step_name","") for s in steps],fontsize=8)
        ax.set_xticks(range(13)); ax.set_xticklabels([""]+[f"W{i}" for i in range(1,13)],fontsize=7.5)
        ax.set_xlim(0,12); ax.set_ylim(-0.6,n-0.4); ax.invert_yaxis()
        for sp in ["top","right","left"]: ax.spines[sp].set_visible(False)
        ax.tick_params(left=False,colors=DGREY)
        patches=[mpatches.Patch(facecolor="#C8DFF0",edgecolor=DBLUE,label="Planned"),
                 mpatches.Patch(facecolor=DBLUE,label="In Progress"),
                 mpatches.Patch(facecolor=GREEN,label="Completed"),
                 plt.Line2D([0],[0],color=RED,lw=2,ls="--",label=f"W{cw}")]
        ax.legend(handles=patches,loc="lower right",fontsize=7,frameon=True,framealpha=0.9,edgecolor=MGREY,ncol=4)
        plt.tight_layout(pad=0.4); return fig

    def kpi_trend_fig():
        weeks=[w["week_number"] for w in kpi_vals]
        vals =[float(w["kpi_value"]) for w in kpi_vals]
        if not weeks: return None
        fig,ax=plt.subplots(figsize=(11,3.5),dpi=180)
        fig.patch.set_facecolor("#FAFAFA"); ax.set_facecolor("#FAFAFA")
        if kpi_baseline: ax.fill_between(weeks,kpi_baseline,vals,alpha=0.1,color=DBLUE)
        ax.axhline(kpi_baseline,color=MGREY,lw=1.5,ls=":",label=f"Baseline {kpi_baseline}")
        ax.axhline(kpi_target,color=GREEN,lw=1.5,ls="--",label=f"Target {kpi_target}")
        ax.plot(weeks,vals,"o-",color=DBLUE,lw=2.5,ms=7,zorder=5,label="Actual")
        for w,v in zip(weeks,vals):
            ax.annotate(f"{v:.1f}",(w,v),textcoords="offset points",xytext=(0,8),fontsize=8,ha="center",color=DBLUE,fontweight="bold")
        ax.axvline(cw,color=RED,lw=1.5,ls=":",alpha=0.6)
        ax.set_xlim(0.3,12.7); mn=min(vals+[kpi_baseline])-3 if vals else 0; mx=max(vals+[kpi_target])+5 if vals else 100
        ax.set_ylim(mn,mx)
        ax.set_xticks(range(1,13)); ax.set_xticklabels([f"W{i}" for i in range(1,13)],fontsize=8)
        ax.set_ylabel(f"{kpi.get('kpi_name','')} {kpi.get('unit','')}",fontsize=9,color=DGREY)
        ax.tick_params(colors=DGREY,labelsize=8)
        ax.legend(fontsize=8,frameon=False,loc="upper left")
        for sp in ["top","right"]: ax.spines[sp].set_visible(False)
        ax.grid(axis="y",color="#EEEEEE",lw=0.5); plt.tight_layout(pad=0.4); return fig

    def kai_fig():
        if not sub_comps: return None
        kai_actuals={}
        for wu in sorted(weekly_updates,key=lambda x:x.get("week_number",0)):
            notes=wu.get("kpi_notes") or ""
            try:
                nd=json.loads(notes) if notes.startswith("{") else {}
                for k,v in nd.items():
                    if isinstance(v,dict): kai_actuals[k]=float(v.get("value",0))
            except: pass
        fig,axes=plt.subplots(1,min(4,len(sub_comps)),figsize=(11,2.8),dpi=180)
        if len(sub_comps)==1: axes=[axes]
        fig.patch.set_facecolor("#FAFAFA")
        for i,kai in enumerate(sub_comps[:4]):
            ax=axes[i]; b=float(kai.get("baseline",0)); t=float(kai.get("target",0)); cur=float(kai_actuals.get(kai.get("name",""),b))
            going_down=t<b; raw=(b-cur)/(b-t)*100 if going_down else (cur-b)/(t-b)*100 if (t-b)!=0 else 0
            pct=max(0,min(100,raw)); col=GREEN if pct>=80 else DBLUE if pct>=50 else AMBER
            ax.pie([pct,100-pct],colors=[col,"#EEEEEE"],startangle=90,wedgeprops={"width":0.38,"edgecolor":"white","linewidth":1.5})
            ax.text(0,0.08,f"{int(pct)}%",ha="center",va="center",fontsize=13,fontweight="bold",color=col)
            ax.text(0,-0.22,"progress",ha="center",va="center",fontsize=6.5,color=DGREY)
            ax.set_title(kai.get("name","")[:22],fontsize=8,fontweight="bold",color=BLACK,pad=3)
            fig.text((i+0.5)/min(4,len(sub_comps)),0.04,f"{b}→{cur}→{t} {kai.get('unit','')}",ha="center",fontsize=7,color=DGREY)
        plt.tight_layout(pad=0.3,rect=[0,0.12,1,1]); return fig

    def score_traj_fig():
        fig,ax=plt.subplots(figsize=(9,3.2),dpi=180)
        fig.patch.set_facecolor("#FAFAFA"); ax.set_facecolor("#FAFAFA")
        tw=sorted(TARGET_RAMP_L.keys()); tt=[TARGET_RAMP_L[w] for w in tw]
        ax.fill_between(tw,0,tt,alpha=0.06,color=MGREY)
        ax.plot(tw,tt,"o--",color=MGREY,lw=1.5,ms=4,label="Target ramp")
        aw=list(range(1,cw+1)); av=[min(total_score,TARGET_RAMP_L.get(w,total_score)) for w in aw]
        ax.fill_between(aw,0,av,alpha=0.1,color=DBLUE)
        ax.plot(aw,av,"o-",color=DBLUE,lw=2.5,ms=7,label="Actual",zorder=5)
        ax.scatter([cw],[total_score],s=80,color=DBLUE,zorder=6)
        tgt_w=TARGET_RAMP_L.get(cw,100)
        ax.scatter([cw],[tgt_w],s=60,color=RED,marker="D",zorder=6,label=f"Target W{cw}: {tgt_w}")
        ax.axvline(cw,color=RED,lw=1.5,ls=":",alpha=0.6)
        ax.set_xlim(0.5,12.5); ax.set_ylim(0,108)
        ax.set_xticks(range(1,13)); ax.set_xticklabels([f"W{i}" for i in range(1,13)],fontsize=8)
        ax.set_ylabel("Score / 100",fontsize=9,color=DGREY); ax.tick_params(colors=DGREY,labelsize=8)
        ax.legend(fontsize=8,frameon=False,loc="upper left")
        for sp in ["top","right"]: ax.spines[sp].set_visible(False)
        ax.grid(axis="y",color="#EEEEEE",lw=0.5); plt.tight_layout(pad=0.4); return fig

    # ── Build PDF ──
    buf = io.BytesIO()
    c = _rl_canvas.Canvas(buf, pagesize=(SW,SH))

    # ── SLIDE 1: COVER ──
    cover_slide(c); c.showPage()

    # ── SLIDE 2: TEAM ──
    content_slide_header(c, "Project Team & Overview")
    y = SH-145
    c.setFont("Helvetica-Bold",10); c.setFillColor(HexColor(DBLUE))
    for tx,th in [(20,"NAME"),(200,"ROLE"),(340,"DEPARTMENT"),(500,"KPI TARGET")]:
        c.drawString(tx,y,th)
    y-=6; c.setStrokeColor(HexColor(MGREY)); c.setLineWidth(0.5); c.line(15,y,SW-15,y); y-=4
    for ti,m in enumerate(team[:8]):
        bg="#F8FBFF" if ti%2==0 else WHITE
        c.setFillColor(HexColor(bg)); c.rect(15,y-4,SW-30,15,fill=1,stroke=0)
        c.setFillColor(HexColor(DBLUE)); c.setFont("Helvetica-Bold",8.5); c.drawString(20,y+3,m.get("member_name",""))
        c.setFillColor(HexColor(DGREY)); c.setFont("Helvetica",8.5)
        c.drawString(200,y+3,m.get("role","")); c.drawString(340,y+3,m.get("department",""))
        # KAI assignment
        kai_assign = next((s.get("name","")[:28] for s in sub_comps if s.get("owner","")==m.get("member_name","")), "—")
        c.drawString(500,y+3,kai_assign); y-=16
    y-=8
    # Problem statement
    c.setFillColor(HexColor(DBLUE)); c.setFont("Helvetica-Bold",10); c.drawString(20,y,"PROBLEM STATEMENT"); y-=14
    c.setFillColor(HexColor(BLACK)); c.setFont("Helvetica",9)
    words=project.get("problem_statement","").split(); line=""; py=y
    for w in words:
        test=(line+" "+w).strip()
        if c.stringWidth(test,"Helvetica",9)<SW-40: line=test
        else: c.drawString(20,py,line); py-=13; line=w
    if line: c.drawString(20,py,line)
    # Footer bar
    c.setFillColor(HexColor(DBLUE)); c.rect(0,0,SW,22,fill=1,stroke=0)
    c.setFillColor(HexColor(WHITE)); c.setFont("Helvetica",8)
    c.drawString(20,6,f"KPI: {kpi.get('kpi_name','')}  ·  Baseline: {kpi_baseline}  →  Target: {kpi_target} {kpi.get('unit','')}  ·  Progress: {kpi_progress:.0f}%")
    c.showPage()

    # ── SLIDE 3: GANTT + KPI TREND ──
    content_slide_header(c, "Master Plan & KPI Trend")
    y = SH-140
    gf=gantt_fig()
    if gf:
        gh=max(130,len(steps)*26+40); gr=_reader(gf)
        c.drawImage(gr,15,y-gh,width=SW-30,height=gh,preserveAspectRatio=True); y-=gh+10
    tf=kpi_trend_fig()
    if tf:
        th2=max(100,SH-y-35); tr=_reader(tf)
        c.drawImage(tr,15,30,width=SW-30,height=th2,preserveAspectRatio=True)
    # Progress bar
    prog_w = (SW-40) * kpi_progress/100
    c.setFillColor(HexColor(LGREY)); c.roundRect(20,10,SW-40,8,4,fill=1,stroke=0)
    if prog_w>0: c.setFillColor(HexColor(GREEN)); c.roundRect(20,10,prog_w,8,4,fill=1,stroke=0)
    c.showPage()

    # ── SLIDE 4: KAIs + ACTIONS ──
    content_slide_header(c, "KAI Progress & Action Plan")
    y = SH-140
    kf=kai_fig()
    if kf and sub_comps:
        kh=115; kr=_reader(kf); c.drawImage(kr,15,y-kh,width=SW-30,height=kh,preserveAspectRatio=True); y-=kh+8
    # Action table
    c.setFillColor(HexColor(DBLUE)); c.rect(15,y-16,SW-30,16,fill=1,stroke=0)
    c.setFillColor(HexColor(WHITE)); c.setFont("Helvetica-Bold",8)
    col_xs=[20,230,340,410,490]; hdrs=["Action","Owner","Date","Status","KAI"]
    for cx,hdr in zip(col_xs,hdrs): c.drawString(cx,y-12,hdr)
    y-=18
    STATUS_COLS={"Completed":GREEN,"In Progress":DBLUE,"Open":MGREY,"Overdue":RED}
    for ai,action in enumerate(actions[:8]):
        bg="#F8FBFF" if ai%2==0 else WHITE
        c.setFillColor(HexColor(bg)); c.rect(15,y-13,SW-30,14,fill=1,stroke=0)
        c.setFillColor(HexColor(BLACK)); c.setFont("Helvetica",7.5)
        c.drawString(20,y-10,str(action.get("description",""))[:34])
        c.drawString(230,y-10,str(action.get("owner","")).split()[0] if action.get("owner") else "")
        c.drawString(340,y-10,str(action.get("target_date",""))[:7])
        sc=STATUS_COLS.get(action.get("status","Open"),MGREY)
        c.setFillColor(HexColor(sc)); c.roundRect(409,y-12,68,12,3,fill=1,stroke=0)
        c.setFillColor(HexColor(WHITE) if action.get("status")!="Open" else HexColor(BLACK))
        c.setFont("Helvetica-Bold",6.5); c.drawCentredString(443,y-7,action.get("status",""))
        c.setFillColor(HexColor(DGREY)); c.setFont("Helvetica",7)
        c.drawString(490,y-10,str(action.get("root_cause_addressed","—"))[:18]); y-=14
    # Summary
    c.setFillColor(HexColor(DBLUE)); c.rect(0,0,SW,22,fill=1,stroke=0)
    c.setFillColor(HexColor(WHITE)); c.setFont("Helvetica-Bold",8.5)
    c.drawString(20,6,f"✔ {completed_cnt} Completed  ●  {in_prog_cnt} In Progress  ○  {open_cnt} Open  ·  On-time: {ot_rate}%")
    c.showPage()

    # ── SLIDE 5: AUDIT SCORE ──
    content_slide_header(c, "Audit Score & Gap Analysis")
    y=SH-145
    # Score trajectory (left side)
    sf=score_traj_fig(); sr=_reader(sf)
    c.drawImage(sr,15,SH-380,width=440,height=200,preserveAspectRatio=True)
    # Traffic lights (right)
    DIMENSIONS=[("Involvement",[1,2,34,35,36],12),("Method",[8,9,10,11,12,13,14,15,16],11),
                ("Action Plan",[17,18,19,20,21,22,23],21),("Results",[3,4,5,6,7,24,25],41),
                ("Stabilisation",[26,27,28,29,30,31,32,33],15)]
    dx=470; dw=(SW-dx-15)/len(DIMENSIONS)
    for i,(dim,qs,possible) in enumerate(DIMENSIONS):
        achieved=sum(scores.get(q,0) for q in qs); pct=achieved/possible*100 if possible>0 else 0
        col=GREEN if pct>=90 else AMBER if pct>=60 else RED
        ddx=dx+i*(dw+2)
        c.setFillColor(HexColor(LGREY)); c.roundRect(ddx,SH-360,dw-2,100,4,fill=1,stroke=0)
        c.setFillColor(HexColor(col)); c.circle(ddx+dw/2,SH-290,22,fill=1,stroke=0)
        c.setFillColor(HexColor(WHITE)); c.setFont("Helvetica-Bold",10); c.drawCentredString(ddx+dw/2,SH-296,f"{int(pct)}%")
        c.setFillColor(HexColor(BLACK)); c.setFont("Helvetica-Bold",7); c.drawCentredString(ddx+dw/2,SH-330,dim[:12])
        c.setFillColor(HexColor(DGREY)); c.setFont("Helvetica",7); c.drawCentredString(ddx+dw/2,SH-342,f"{achieved}/{possible}")
    # Gap register
    y=SH-400
    QUESTIONS_L={4:"Cost/Benefit",15:"Reoccurrence",16:"Single problem",24:"KPI trend +",25:"Goal achieved",26:"Procedures",28:"Monitoring updated",29:"OPLs created",30:"Training matrix",31:"CIL ≥90%"}
    WDUE={4:5,15:7,16:7,24:7,25:11,26:10,28:8,29:10,30:10,31:5}
    gaps=[(qn,txt) for qn,txt in QUESTIONS_L.items() if scores.get(qn,0)==0 and WDUE.get(qn,99)<=cw]
    c.setFillColor(HexColor(DBLUE)); c.setFont("Helvetica-Bold",9); c.drawString(20,y,"GAP REGISTER:"); y-=14
    for qn,txt in gaps[:5]:
        c.setFillColor(HexColor(RED)); c.circle(26,y+4,4,fill=1,stroke=0)
        c.setFillColor(HexColor(BLACK)); c.setFont("Helvetica",8.5)
        c.drawString(34,y,f"Q{qn}: {txt}  (due W{WDUE.get(qn,'?')})"); y-=13
    # Auditor notes
    ar=audit_records[-1] if audit_records else {}
    notes=ar.get("auditor_notes","")
    if notes:
        c.setFillColor(HexColor(LGREY)); c.rect(15,10,SW-30,55,fill=1,stroke=0)
        c.setFillColor(HexColor(DBLUE)); c.setFont("Helvetica-Bold",8); c.drawString(20,56,"AUDITOR NOTES")
        c.setFillColor(HexColor(BLACK)); c.setFont("Helvetica",8); c.drawString(20,43,notes[:90])
        c.setFillColor(HexColor(DGREY)); c.setFont("Helvetica-Oblique",7.5); c.drawRightString(SW-20,13,f"Audited: {ar.get('audited_by','')}  ·  W{ar.get('week_number','')}")
    c.showPage()

    c.save(); buf.seek(0)
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
