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
# BEAUTIFUL PDF REPORT GENERATOR
# ════════════════════════════════════════
def _generate_project_pdf(project, team, kpi, steps, weekly_updates,
                          actions, stab, audit_records, scores, total_score):
    """
    Redesigned PDF report generator.
    Slides: Cover | Team | Gantt | KPI Trend | KAI(s) | Actions | Audit
    """
    import io, json, base64 as _b64, tempfile as _tf, math
    from datetime import date
    import matplotlib.pyplot as plt
    import matplotlib.patches as mpatches
    import matplotlib as mpl
    from reportlab.pdfgen import canvas as _rl_canvas
    from reportlab.lib.utils import ImageReader
    from reportlab.lib.colors import HexColor
    mpl.rcParams['font.family'] = 'DejaVu Sans'

    # ── Embedded helpers ──
    import math as _math
    def _get_current_week(launch_date):
        if not launch_date: return 1
        if isinstance(launch_date, str):
            try: launch_date = date.fromisoformat(launch_date)
            except: return 1
        delta = (date.today() - launch_date).days
        return max(1, min(12, _math.ceil(delta / 7) if delta > 0 else 1))

    
    
    
    
    # ── Palette ──────────────────────────────────────────────
    DBLUE  = "#0C5595"
    LBLUE  = "#1A73C8"
    RED    = "#DE201B"
    LGREY  = "#F4F6F8"
    MGREY  = "#BDC3C7"
    DGREY  = "#566573"
    GREEN  = "#1E8449"
    LGREEN = "#D5F5E3"
    AMBER  = "#D68910"
    LAMBER = "#FEF9E7"
    WHITE  = "#FFFFFF"
    BLACK  = "#1A1A2E"
    DKBG   = "#0A1628"   # dark navy for score ring slide
    
    
    # ── Helpers ──────────────────────────────────────────────
    def _reader(fig):
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=180, bbox_inches="tight",
                    facecolor=fig.get_facecolor())
        buf.seek(0)
        plt.close(fig)
        return ImageReader(buf)
    
    
    def _b64_to_reader(b64str):
        try:
            data = _b64.b64decode(b64str)
            tmp = _tf.NamedTemporaryFile(delete=False, suffix=".png")
            tmp.write(data); tmp.flush()
            return ImageReader(tmp.name)
        except Exception:
            return None
    
    
    def _napco_logo_reader_from_b64(logo_b64):
        try:
            data = _b64.b64decode(logo_b64)
            tmp = _tf.NamedTemporaryFile(delete=False, suffix=".jpeg")
            tmp.write(data); tmp.flush()
            return ImageReader(tmp.name)
        except Exception:
            return None
    
    
    def _wrap_text(c, text, x, y, max_width, font, size, color=None, line_h=None):
        """Draw wrapped text; return final y."""
        if color:
            c.setFillColor(HexColor(color))
        c.setFont(font, size)
        lh = line_h or size * 1.4
        words = (text or "").split()
        line = ""
        for w in words:
            test = (line + " " + w).strip()
            if c.stringWidth(test, font, size) <= max_width:
                line = test
            else:
                c.drawString(x, y, line)
                y -= lh
                line = w
        if line:
            c.drawString(x, y, line)
            y -= lh
        return y
    
    
    # ── Shared slide chrome ───────────────────────────────────
    def _slide_bg(c):
        c.setFillColor(HexColor(WHITE))
        c.rect(0, 0, SW, SH, fill=1, stroke=0)
    
    
    def _top_bar(c, title, subtitle=""):
        """Clean dark-navy top bar, no logo."""
        c.setFillColor(HexColor(DKBG))
        c.rect(0, SH - 60, SW, 60, fill=1, stroke=0)
        # left accent stripe
        c.setFillColor(HexColor(RED))
        c.rect(0, SH - 60, 6, 60, fill=1, stroke=0)
        c.setFillColor(HexColor(WHITE))
        c.setFont("Helvetica-Bold", 20)
        c.drawString(20, SH - 38, title)
        if subtitle:
            c.setFont("Helvetica", 10)
            c.setFillColor(HexColor(MGREY))
            c.drawString(20, SH - 54, subtitle)
    
    
    def _bottom_bar(c, left_text="", right_text=""):
        c.setFillColor(HexColor(DKBG))
        c.rect(0, 0, SW, 28, fill=1, stroke=0)
        c.setFillColor(HexColor(MGREY))
        c.setFont("Helvetica", 8)
        if left_text:
            c.drawString(16, 9, left_text)
        if right_text:
            c.drawRightString(SW - 16, 9, right_text)
    
    
    def _section_label(c, text, x, y, w=200, h=18, bg=DBLUE, fg=WHITE):
        c.setFillColor(HexColor(bg))
        c.roundRect(x, y, w, h, 4, fill=1, stroke=0)
        c.setFillColor(HexColor(fg))
        c.setFont("Helvetica-Bold", 8)
        c.drawString(x + 8, y + 5, text.upper())
    
    
    def _kv_badge(c, x, y, w, h, label, value, sub="", bg=DBLUE, fg=WHITE, vfont_size=22):
        """Rounded rectangle badge with label/value/sub."""
        c.setFillColor(HexColor(bg))
        c.roundRect(x, y, w, h, 7, fill=1, stroke=0)
        c.setFillColor(HexColor(fg))
        c.setFont("Helvetica", 7.5)
        c.drawCentredString(x + w / 2, y + h - 12, label.upper())
        c.setFont("Helvetica-Bold", vfont_size if len(str(value)) <= 7 else vfont_size - 4)
        c.drawCentredString(x + w / 2, y + h / 2 - 4, str(value))
        if sub:
            c.setFont("Helvetica", 7)
            c.setFillColor(HexColor(MGREY) if bg == DBLUE else HexColor(DGREY))
            c.drawCentredString(x + w / 2, y + 7, str(sub))
    
    
    # ════════════════════════════════════════════════════════
    # SLIDE 1 – COVER
    # ════════════════════════════════════════════════════════
    def _slide_cover(c, project, kpi, cw, total_score, score_col, kpi_progress,
                     logo_reader):
        _slide_bg(c)
    
        # ── full-width dark header band ──
        c.setFillColor(HexColor(DKBG))
        c.rect(0, SH - 150, SW, 150, fill=1, stroke=0)
        c.setFillColor(HexColor(RED))
        c.rect(0, SH - 150, 8, 150, fill=1, stroke=0)
    
        # Logo in header
        if logo_reader:
            try:
                c.drawImage(logo_reader, 20, SH - 135, width=230, height=90,
                            preserveAspectRatio=True, mask="auto")
            except Exception:
                pass
    
        # Project title
        c.setFillColor(HexColor(WHITE))
        c.setFont("Helvetica-Bold", 26)
        pname = project.get("project_name", "")
        c.drawString(270, SH - 55, pname[:50])
        c.setFont("Helvetica", 12)
        c.setFillColor(HexColor(MGREY))
        c.drawString(270, SH - 78, f"{project.get('target_area','')}   ·   Week {cw} of 12")
        c.setFont("Helvetica", 10)
        c.drawString(270, SH - 98, f"Launch: {str(project.get('launch_date',''))[:10]}   ·   "
                                   f"Expected: {str(project.get('expected_completion_date',''))[:10]}")
    
        # ── Score ring (right side of header) ──
        cx, cy, r = SW - 75, SH - 80, 52
        # outer ring bg
        c.setFillColor(HexColor("#1E2D45"))
        c.circle(cx, cy, r + 6, fill=1, stroke=0)
        # coloured ring
        c.setFillColor(HexColor(score_col))
        c.circle(cx, cy, r, fill=1, stroke=0)
        # inner dark circle
        c.setFillColor(HexColor(DKBG))
        c.circle(cx, cy, r - 10, fill=1, stroke=0)
        c.setFillColor(HexColor(WHITE))
        c.setFont("Helvetica-Bold", 28)
        c.drawCentredString(cx, cy - 9, str(int(total_score)))
        c.setFont("Helvetica", 8)
        c.setFillColor(HexColor(MGREY))
        c.drawCentredString(cx, cy + 20, "SCORE")
        c.drawCentredString(cx, cy - 20, "/ 100")
    
        # ── KPI Metric row ──────────────────────────────────
        # 4 badges across full width with nice spacing
        pad = 30
        bw = (SW - pad * 2 - 30) / 4
        bh = 80
        by = SH - 260
        kpi_unit = kpi.get("unit", "") if kpi else ""
        kpi_baseline = float(kpi.get("baseline_value", 0) or 0) if kpi else 0
        kpi_target = float(kpi.get("target_value", 0) or 0) if kpi else 0
        kpi_vals_list = []  # passed from caller via kpi_current
    
        badges = [
            ("Baseline",  f"{kpi_baseline}{kpi_unit}", "Start value",   "#17375E", WHITE),
            ("Current",   f"{project.get('_kpi_current', kpi_baseline)}{kpi_unit}", f"Week {cw}", DBLUE, WHITE),
            ("Target",    f"{kpi_target}{kpi_unit}",   "Goal",          "#145A32", WHITE),
            ("Progress",  f"{kpi_progress:.0f}%",      "Toward target",
             GREEN if kpi_progress >= 80 else AMBER if kpi_progress >= 40 else RED, WHITE),
        ]
        for i, (lbl, val, sub, bg, fg) in enumerate(badges):
            bx = pad + i * (bw + 10)
            _kv_badge(c, bx, by, bw, bh, lbl, val, sub, bg, fg, vfont_size=24)
    
        # ── Progress bar ──
        bar_y = by - 22
        bar_w = SW - pad * 2
        c.setFillColor(HexColor(LGREY))
        c.roundRect(pad, bar_y, bar_w, 10, 5, fill=1, stroke=0)
        filled = bar_w * min(kpi_progress / 100, 1)
        if filled > 0:
            bar_col = GREEN if kpi_progress >= 80 else AMBER if kpi_progress >= 40 else RED
            c.setFillColor(HexColor(bar_col))
            c.roundRect(pad, bar_y, filled, 10, 5, fill=1, stroke=0)
        c.setFillColor(HexColor(DGREY))
        c.setFont("Helvetica", 7.5)
        c.drawRightString(pad + bar_w, bar_y - 8, f"{kpi_progress:.0f}% complete")
    
        # ── Problem statement ──
        prob_y = bar_y - 26
        _section_label(c, "Problem Statement", pad, prob_y, w=160)
        prob_y -= 18
        c.setFont("Helvetica", 9.5)
        c.setFillColor(HexColor(BLACK))
        prob_text = project.get("problem_statement", "")
        words = prob_text.split()
        line = ""
        for w in words:
            test = (line + " " + w).strip()
            if c.stringWidth(test, "Helvetica", 9.5) < SW - pad * 2:
                line = test
            else:
                c.drawString(pad, prob_y, line)
                prob_y -= 14
                line = w
        if line:
            c.drawString(pad, prob_y, line)
    
        # Company KPI link
        if project.get("company_kpi_link"):
            _bottom_bar(c,
                        f"Company KPI: {project['company_kpi_link']}",
                        f"Generated {date.today().strftime('%d %B %Y')}")
        else:
            _bottom_bar(c, right_text=f"Generated {date.today().strftime('%d %B %Y')}")
    
    
    # ════════════════════════════════════════════════════════
    # SLIDE 2 – TEAM & OVERVIEW
    # ════════════════════════════════════════════════════════
    def _slide_team(c, project, team, kpi, cw):
        _slide_bg(c)
        _top_bar(c, "Project Team", f"{project.get('target_area','')}  ·  Week {cw}")
    
        sub_comps = []
        if kpi:
            raw = kpi.get("sub_components") or []
            if isinstance(raw, str):
                try: raw = json.loads(raw)
                except: raw = []
            sub_comps = [s for s in raw if isinstance(s, dict)]
        owner_map = {}
        for s in sub_comps:
            ow = s.get("owner", "")
            if ow:
                owner_map.setdefault(ow, []).append(
                    f"{s.get('name','')} ({s.get('baseline','')}→{s.get('target','')} {s.get('unit','')})")
    
        # ── Table ──
        PAD = 22
        col_x  = [PAD, 210, 330, 460, 680]
        col_w  = [185, 115, 125, 215, 200]
        hdrs   = ["TEAM MEMBER", "ROLE", "DEPARTMENT", "KAI TARGET", "CONTRIBUTION"]
        row_h  = 28
        hdr_y  = SH - 85
    
        # header row
        c.setFillColor(HexColor(DBLUE))
        c.rect(PAD, hdr_y - 4, SW - PAD * 2, row_h, fill=1, stroke=0)
        c.setFillColor(HexColor(WHITE))
        c.setFont("Helvetica-Bold", 9)
        for x, hdr in zip(col_x, hdrs):
            c.drawString(x + 6, hdr_y + 8, hdr)
    
        row_y = hdr_y - row_h
        for ti, m in enumerate(team[:12]):
            bg = LGREY if ti % 2 == 0 else WHITE
            c.setFillColor(HexColor(bg))
            c.rect(PAD, row_y - 4, SW - PAD * 2, row_h, fill=1, stroke=0)
    
            mname = m.get("member_name", "")
            vals = [
                mname,
                m.get("role", ""),
                m.get("department", ""),
                "; ".join(owner_map.get(mname, ["—"]))[:45],
                m.get("contribution_target", "—") or "—",
            ]
            c.setFillColor(HexColor(BLACK))
            c.setFont("Helvetica-Bold" if ti == 0 else "Helvetica", 9)
            for xi, (x, val) in enumerate(zip(col_x, vals)):
                txt = str(val)
                # truncate to fit column
                while c.stringWidth(txt, "Helvetica", 9) > col_w[xi] - 10 and len(txt) > 4:
                    txt = txt[:-2] + "…"
                c.drawString(x + 6, row_y + 7, txt)
            row_y -= row_h
    
        # thin separator lines
        c.setStrokeColor(HexColor(MGREY))
        c.setLineWidth(0.3)
        for cx in col_x[1:]:
            c.line(cx, SH - 85, cx, row_y + row_h)
    
        # ── Right panel: KPI summary card ──
        # Only if space
        if kpi:
            kp_x = SW - 295
            kp_y = 38
            kp_w = 275
            kp_h = (hdr_y - 4) - kp_y - 10
            if kp_h > 60:
                c.setFillColor(HexColor(LGREY))
                c.roundRect(kp_x, kp_y, kp_w, kp_h, 6, fill=1, stroke=0)
                _section_label(c, "KPI Summary", kp_x + 10, kp_y + kp_h - 22, w=130)
                yy = kp_y + kp_h - 42
                c.setFont("Helvetica-Bold", 11)
                c.setFillColor(HexColor(DBLUE))
                c.drawString(kp_x + 10, yy, kpi.get("kpi_name", ""))
                yy -= 18
                for lbl, val in [
                    ("Baseline", f"{kpi.get('baseline_value','')} {kpi.get('unit','')}"),
                    ("Target",   f"{kpi.get('target_value','')} {kpi.get('unit','')}"),
                    ("Due",      str(kpi.get("target_date",""))[:10]),
                ]:
                    c.setFont("Helvetica-Bold", 8)
                    c.setFillColor(HexColor(DGREY))
                    c.drawString(kp_x + 10, yy, lbl + ":")
                    c.setFont("Helvetica", 9)
                    c.setFillColor(HexColor(BLACK))
                    c.drawString(kp_x + 70, yy, str(val))
                    yy -= 16
    
        _bottom_bar(c, f"{len(team)} team members",
                    f"Project: {project.get('project_name','')[:50]}")
    
    
    # ════════════════════════════════════════════════════════
    # SLIDE 3 – MASTER PLAN (Gantt only, full slide)
    # ════════════════════════════════════════════════════════
    def _gantt_fig_full(steps, weekly_updates, current_week):
        if not steps:
            return None
        n = len(steps)
        fig_h = max(4.5, n * 0.55 + 1.5)
        fig, ax = plt.subplots(figsize=(13.5, fig_h), dpi=180)
        fig.patch.set_facecolor("#FAFAFA")
        ax.set_facecolor("#FAFAFA")
    
        sp_map = {}
        for wu in weekly_updates:
            sp_raw = wu.get("step_progress") or []
            if isinstance(sp_raw, str):
                try: sp_raw = json.loads(sp_raw)
                except: sp_raw = []
            for sp in sp_raw:
                if isinstance(sp, dict):
                    sid = sp.get("step_id", "")
                    sp_map[sid] = max(sp_map.get(sid, 0), sp.get("pct_complete", 0))
    
        bar_h = 0.55
        for i, step in enumerate(steps):
            ps = step.get("planned_start_week", 1)
            pe = step.get("planned_end_week", 2)
            # background
            ax.barh(i, pe - ps, left=ps - 1, height=bar_h,
                    color="#C8DFF0", zorder=2)
            ax.barh(i, pe - ps, left=ps - 1, height=bar_h,
                    color="none", edgecolor=DBLUE, linewidth=1.0, zorder=3)
            pct = sp_map.get(str(step.get("id", "")), 0)
            bar_col = GREEN if pct == 100 else DBLUE
            if pct > 0:
                ax.barh(i, (pe - ps) * pct / 100, left=ps - 1, height=bar_h,
                        color=bar_col, alpha=0.88, zorder=4)
                if pct >= 20:
                    ax.text(ps - 1 + (pe - ps) * pct / 200, i, f"{pct}%",
                            ha="center", va="center", fontsize=7.5,
                            color="white", fontweight="bold", zorder=5)
            # owner tag
            if step.get("owner"):
                ax.text(pe - 0.06, i - bar_h / 2 - 0.06,
                        step["owner"][:18], ha="right", va="top",
                        fontsize=6.5, color=DGREY, zorder=5)
            # due indicator if not started and overdue
            if pct == 0 and ps <= current_week:
                ax.barh(i, pe - ps, left=ps - 1, height=bar_h,
                        color="none", edgecolor=RED, linewidth=1.5,
                        linestyle="--", zorder=5)
    
        # week grid
        for w in range(13):
            ax.axvline(w, color="#EBEBEB", linewidth=0.5, zorder=1)
    
        # current week line
        ax.axvline(current_week - 1, color=RED, linewidth=2,
                   linestyle="--", zorder=6, alpha=0.85)
        ax.text(current_week - 0.9, n - 0.2,
                f"W{current_week}", color=RED, fontsize=8,
                fontweight="bold", zorder=7)
    
        ax.set_yticks(range(n))
        ax.set_yticklabels([s.get("step_name", "") for s in steps],
                           fontsize=9, fontweight="bold", color=BLACK)
        ax.set_xticks(range(13))
        ax.set_xticklabels([""] + [f"W{i}" for i in range(1, 13)],
                           fontsize=9, color=DGREY)
        ax.set_xlim(0, 12)
        ax.set_ylim(-0.7, n - 0.3)
        ax.invert_yaxis()
        for sp in ["top", "right", "left"]:
            ax.spines[sp].set_visible(False)
        ax.spines["bottom"].set_color(MGREY)
        ax.tick_params(left=False)
    
        patches = [
            mpatches.Patch(facecolor="#C8DFF0", edgecolor=DBLUE, label="Planned"),
            mpatches.Patch(facecolor=DBLUE, label="In Progress"),
            mpatches.Patch(facecolor=GREEN, label="Completed"),
            plt.Line2D([0], [0], color=RED, lw=2, ls="--", label=f"Current (W{current_week})"),
        ]
        ax.legend(handles=patches, loc="lower right", fontsize=8,
                  frameon=True, framealpha=0.92, edgecolor=MGREY, ncol=4)
        plt.tight_layout(pad=0.6)
        return fig
    
    
    def _slide_gantt(c, project, steps, wu_rows, cw):
        _slide_bg(c)
        _top_bar(c, "Master Plan", f"Week {cw} of 12  ·  {len(steps)} phases")
    
        fig = _gantt_fig_full(steps, wu_rows, cw)
        if fig:
            ir = _reader(fig)
            # Use almost full slide below header
            img_y = 32
            img_h = SH - 68 - 32
            c.drawImage(ir, 14, img_y, width=SW - 28, height=img_h,
                        preserveAspectRatio=True)
        else:
            c.setFillColor(HexColor(MGREY))
            c.setFont("Helvetica", 14)
            c.drawCentredString(SW / 2, SH / 2, "No steps defined yet.")
    
        _bottom_bar(c, project.get("project_name", "")[:60],
                    f"Generated {date.today().strftime('%d %b %Y')}")
    
    
    # ════════════════════════════════════════════════════════
    # SLIDE 4 – KPI TREND (full slide)
    # ════════════════════════════════════════════════════════
    def _kpi_trend_fig_full(kpi, wu_rows, cw):
        kpi_vals = sorted(
            [w for w in wu_rows if w.get("kpi_value") is not None],
            key=lambda x: x["week_number"]
        )
        baseline = float(kpi.get("baseline_value", 0) or 0)
        target   = float(kpi.get("target_value", 0) or 0)
        weeks    = [w["week_number"] for w in kpi_vals]
        vals     = [float(w["kpi_value"]) for w in kpi_vals]
    
        fig, ax = plt.subplots(figsize=(13.5, 5.5), dpi=180)
        fig.patch.set_facecolor("#FAFAFA")
        ax.set_facecolor("#FAFAFA")
    
        ax.axhline(baseline, color=MGREY, lw=1.5, ls=":", label=f"Baseline {baseline}")
        ax.axhline(target,   color=GREEN, lw=2.0, ls="--", label=f"Target {target}")
    
        if weeks:
            ax.fill_between(weeks, baseline, vals, alpha=0.12, color=DBLUE)
            ax.plot(weeks, vals, "o-", color=DBLUE, lw=2.5, ms=8,
                    zorder=5, label="Actual")
            for w, v in zip(weeks, vals):
                ax.annotate(f"{v:.1f}", (w, v),
                            textcoords="offset points", xytext=(0, 10),
                            fontsize=9, ha="center",
                            color=DBLUE, fontweight="bold")
    
        ax.axvline(cw, color=RED, lw=1.5, ls=":", alpha=0.65,
                   label=f"Now (W{cw})")
        ax.set_xlim(0.5, 12.5)
    
        mn = min(vals + [baseline]) - abs(target - baseline) * 0.1 if vals else 0
        mx = max(vals + [target])   + abs(target - baseline) * 0.15 if vals else 100
        ax.set_ylim(mn, mx)
    
        ax.set_xticks(range(1, 13))
        ax.set_xticklabels([f"W{i}" for i in range(1, 13)], fontsize=10)
        ax.set_ylabel(
            f"{kpi.get('kpi_name', '')} ({kpi.get('unit', '')})",
            fontsize=10, color=DGREY
        )
        ax.tick_params(colors=DGREY, labelsize=9)
        ax.legend(fontsize=9, frameon=False, loc="upper left")
        ax.grid(axis="y", color="#EBEBEB", lw=0.6)
        for sp in ["top", "right"]:
            ax.spines[sp].set_visible(False)
        ax.spines["bottom"].set_color(MGREY)
        ax.spines["left"].set_color(MGREY)
    
        plt.tight_layout(pad=0.6)
        return fig
    
    
    def _slide_kpi(c, project, kpi, wu_rows, cw, kpi_progress):
        _slide_bg(c)
        _top_bar(c, f"KPI Trend — {kpi.get('kpi_name','')}", f"Week {cw}")
    
        fig = _kpi_trend_fig_full(kpi, wu_rows, cw)
        ir = _reader(fig)
        c.drawImage(ir, 14, 32, width=SW - 28, height=SH - 68 - 32,
                    preserveAspectRatio=True)
    
        _bottom_bar(c,
                    f"Baseline {kpi.get('baseline_value','')} → Target {kpi.get('target_value','')} {kpi.get('unit','')}  ·  Progress: {kpi_progress:.0f}%",
                    project.get("project_name", "")[:50])
    
    
    # ════════════════════════════════════════════════════════
    # SLIDE 5a – KAI PROGRESS (combined if ≤4, else one each)
    # ════════════════════════════════════════════════════════
    def _kai_detail_fig(kai, wu_rows):
        """Trend chart for a single KAI."""
        kn = kai.get("name", "")
        baseline = float(kai.get("baseline", 0) or 0)
        target   = float(kai.get("target", 0) or 0)
    
        actuals = {}
        for wu in sorted(wu_rows, key=lambda x: x.get("week_number", 0)):
            notes = wu.get("kpi_notes") or ""
            try:
                nd = json.loads(notes) if notes.startswith("{") else {}
                if kn in nd:
                    actuals[wu["week_number"]] = float(nd[kn].get("value", 0))
            except Exception:
                pass
    
        weeks = sorted(actuals.keys())
        vals  = [actuals[w] for w in weeks]
    
        fig, ax = plt.subplots(figsize=(6, 3.2), dpi=180)
        fig.patch.set_facecolor("#FAFAFA")
        ax.set_facecolor("#FAFAFA")
    
        ax.axhline(baseline, color=MGREY, lw=1.5, ls=":", label=f"Baseline {baseline}")
        ax.axhline(target,   color=GREEN, lw=2.0, ls="--", label=f"Target {target}")
    
        if weeks:
            ax.fill_between(weeks, baseline, vals, alpha=0.15, color=DBLUE)
            ax.plot(weeks, vals, "o-", color=DBLUE, lw=2.5, ms=7,
                    zorder=5, label="Actual")
            for w, v in zip(weeks, vals):
                ax.annotate(f"{v:.1f}", (w, v),
                            textcoords="offset points", xytext=(0, 8),
                            fontsize=8, ha="center", color=DBLUE,
                            fontweight="bold")
    
        ax.set_xlim(0.5, 12.5)
        ax.set_xticks(range(1, 13))
        ax.set_xticklabels([f"W{i}" for i in range(1, 13)], fontsize=8)
        ax.set_ylabel(f"{kn} ({kai.get('unit','')})", fontsize=8, color=DGREY)
        ax.tick_params(colors=DGREY, labelsize=8)
        ax.legend(fontsize=8, frameon=False)
        ax.grid(axis="y", color="#EBEBEB", lw=0.5)
        for sp in ["top", "right"]:
            ax.spines[sp].set_visible(False)
        plt.tight_layout(pad=0.5)
        return fig
    
    
    def _kai_donut_value(kai, wu_rows):
        """Return current value and progress % for a KAI."""
        kn = kai.get("name", "")
        baseline = float(kai.get("baseline", 0) or 0)
        target   = float(kai.get("target", 0) or 0)
        cur = baseline
        for wu in sorted(wu_rows, key=lambda x: x.get("week_number", 0)):
            notes = wu.get("kpi_notes") or ""
            try:
                nd = json.loads(notes) if notes.startswith("{") else {}
                if kn in nd:
                    cur = float(nd[kn].get("value", 0))
            except Exception:
                pass
        total_gap = abs(target - baseline)
        pct = abs(cur - baseline) / total_gap * 100 if total_gap > 0 else 0
        return cur, min(100, max(0, pct))
    
    
    def _slide_kai_combined(c, project, sub_comps, wu_rows, kpi, cw):
        """One slide with all KAIs as donut + mini-trend side by side."""
        _slide_bg(c)
        _top_bar(c, "KAI Progress", f"Week {cw}  ·  {len(sub_comps)} Key Activity Indicators")
    
        n = len(sub_comps)
        cols = min(n, 2)
        rows = math.ceil(n / cols)
        cell_w = (SW - 40) / cols
        cell_h = (SH - 100) / rows
        pad = 20
    
        for idx, kai in enumerate(sub_comps):
            col_i = idx % cols
            row_i = idx // cols
            bx = 20 + col_i * cell_w
            by = SH - 80 - (row_i + 1) * cell_h + pad
    
            cur, pct = _kai_donut_value(kai, wu_rows)
            baseline = float(kai.get("baseline", 0) or 0)
            target   = float(kai.get("target", 0) or 0)
            bar_col  = GREEN if pct >= 80 else AMBER if pct >= 40 else RED
    
            # card bg
            c.setFillColor(HexColor(LGREY))
            c.roundRect(bx + 4, by, cell_w - 8, cell_h - pad, 8, fill=1, stroke=0)
    
            # KAI name
            c.setFont("Helvetica-Bold", 10)
            c.setFillColor(HexColor(DBLUE))
            c.drawString(bx + 16, by + cell_h - pad - 16, kai.get("name", "")[:40])
            c.setFont("Helvetica", 8)
            c.setFillColor(HexColor(DGREY))
            c.drawString(bx + 16, by + cell_h - pad - 28,
                         f"Owner: {kai.get('owner','—')}   |   "
                         f"Baseline: {baseline}  →  Target: {target} {kai.get('unit','')}")
    
            # progress bar
            bar_x = bx + 16
            bar_bw = cell_w - 40
            bar_y2 = by + cell_h - pad - 44
            c.setFillColor(HexColor(MGREY))
            c.roundRect(bar_x, bar_y2, bar_bw, 8, 4, fill=1, stroke=0)
            if pct > 0:
                c.setFillColor(HexColor(bar_col))
                c.roundRect(bar_x, bar_y2, bar_bw * pct / 100, 8, 4, fill=1, stroke=0)
            c.setFont("Helvetica-Bold", 9)
            c.setFillColor(HexColor(bar_col))
            c.drawString(bar_x + bar_bw + 6, bar_y2, f"{pct:.0f}%")
    
            # current value badge
            c.setFillColor(HexColor(bar_col))
            c.roundRect(bx + 16, by + 8, 80, 28, 5, fill=1, stroke=0)
            c.setFillColor(HexColor(WHITE))
            c.setFont("Helvetica-Bold", 12)
            c.drawCentredString(bx + 56, by + 18, f"{cur:.1f}")
            c.setFont("Helvetica", 7)
            c.drawCentredString(bx + 56, by + 9, "Current")
    
            # mini trend chart
            fig = _kai_detail_fig(kai, wu_rows)
            ir = _reader(fig)
            chart_h = cell_h - pad - 58
            chart_w = cell_w - 120
            if chart_h > 40 and chart_w > 80:
                c.drawImage(ir, bx + 110, by + 8,
                            width=chart_w, height=chart_h,
                            preserveAspectRatio=True)
    
        _bottom_bar(c, project.get("project_name", "")[:60],
                    f"KPI: {kpi.get('kpi_name','') if kpi else ''}")
    
    
    def _slide_kai_single(c, project, kai, wu_rows, kpi, cw, idx, total):
        """One full slide for a single KAI."""
        _slide_bg(c)
        _top_bar(c, f"KAI: {kai.get('name','')[:55]}",
                 f"Week {cw}  ·  KAI {idx+1} of {total}")
    
        cur, pct = _kai_donut_value(kai, wu_rows)
        baseline = float(kai.get("baseline", 0) or 0)
        target   = float(kai.get("target", 0) or 0)
        bar_col  = GREEN if pct >= 80 else AMBER if pct >= 40 else RED
    
        # Left info panel
        px, py, pw = 20, 40, 230
        c.setFillColor(HexColor(LGREY))
        c.roundRect(px, py, pw, SH - 100, 8, fill=1, stroke=0)
    
        c.setFont("Helvetica-Bold", 22)
        c.setFillColor(HexColor(bar_col))
        c.drawCentredString(px + pw / 2, SH - 108, f"{pct:.0f}%")
        c.setFont("Helvetica", 9)
        c.setFillColor(HexColor(DGREY))
        c.drawCentredString(px + pw / 2, SH - 124, "progress toward target")
    
        # big current value
        c.setFillColor(HexColor(bar_col))
        c.roundRect(px + 20, SH - 200, pw - 40, 55, 8, fill=1, stroke=0)
        c.setFillColor(HexColor(WHITE))
        c.setFont("Helvetica-Bold", 30)
        c.drawCentredString(px + pw / 2, SH - 180, f"{cur:.1f}")
        c.setFont("Helvetica", 9)
        c.drawCentredString(px + pw / 2, SH - 163, f"{kai.get('unit','')}  Current")
    
        # stats
        y_s = SH - 222
        for lbl, val in [
            ("Baseline", f"{baseline} {kai.get('unit','')}"),
            ("Target",   f"{target} {kai.get('unit','')}"),
            ("Gap left", f"{abs(target - cur):.1f} {kai.get('unit','')}"),
            ("Owner",    kai.get("owner", "—")),
        ]:
            c.setFont("Helvetica-Bold", 8)
            c.setFillColor(HexColor(DGREY))
            c.drawString(px + 12, y_s, lbl + ":")
            c.setFont("Helvetica", 9)
            c.setFillColor(HexColor(BLACK))
            c.drawString(px + 80, y_s, str(val))
            y_s -= 18
    
        # horizontal progress bar
        c.setFillColor(HexColor(MGREY))
        c.roundRect(px + 12, y_s - 6, pw - 24, 10, 5, fill=1, stroke=0)
        if pct > 0:
            c.setFillColor(HexColor(bar_col))
            c.roundRect(px + 12, y_s - 6, (pw - 24) * pct / 100, 10, 5, fill=1, stroke=0)
        y_s -= 24
    
        # Right: full-width trend chart
        fig = _kai_detail_fig(kai, wu_rows)
        ir = _reader(fig)
        c.drawImage(ir, 265, 36, width=SW - 280, height=SH - 100 - 36,
                    preserveAspectRatio=True)
    
        _bottom_bar(c, f"KPI: {kpi.get('kpi_name','') if kpi else ''}",
                    project.get("project_name", "")[:50])
    
    
    # ════════════════════════════════════════════════════════
    # SLIDE 6 – ACTION PLAN
    # ════════════════════════════════════════════════════════
    def _slide_actions(c, project, actions, cw):
        _slide_bg(c)
        completed = sum(1 for a in actions if a.get("status") == "Completed")
        in_prog   = sum(1 for a in actions if a.get("status") == "In Progress")
        overdue   = sum(1 for a in actions
                        if a.get("status") == "Overdue" or
                        (a.get("target_date") and
                         date.fromisoformat(str(a["target_date"])[:10]) < date.today()
                         and a.get("status") != "Completed"))
        _top_bar(c, "Action Plan",
                 f"{len(actions)} actions  ·  {completed} completed  ·  {in_prog} in progress  ·  {overdue} overdue")
    
        STATUS_COL = {
            "Completed":  GREEN,
            "In Progress": DBLUE,
            "Open":       MGREY,
            "Overdue":    RED,
        }
        PAD = 16
        col_x = [PAD, 340, 470, 570, 680, 800]
        col_w = [320, 125, 95,  105, 115, 120]
        hdrs  = ["ACTION", "OWNER", "DUE DATE", "STATUS", "ROOT CAUSE", "EVIDENCE"]
        row_h = 26
        hdr_y = SH - 82
    
        # header
        c.setFillColor(HexColor(DBLUE))
        c.rect(PAD, hdr_y - 2, SW - PAD * 2, row_h, fill=1, stroke=0)
        c.setFillColor(HexColor(WHITE))
        c.setFont("Helvetica-Bold", 8.5)
        for x, hdr in zip(col_x, hdrs):
            c.drawString(x + 5, hdr_y + 8, hdr)
    
        row_y = hdr_y - row_h
        max_rows = min(len(actions), 13)
        for ai, action in enumerate(actions[:max_rows]):
            bg = LGREY if ai % 2 == 0 else WHITE
            c.setFillColor(HexColor(bg))
            c.rect(PAD, row_y - 2, SW - PAD * 2, row_h, fill=1, stroke=0)
    
            # status colour pill
            status = action.get("status", "Open")
            sc = STATUS_COL.get(status, MGREY)
            c.setFillColor(HexColor(sc))
            c.roundRect(col_x[3] + 3, row_y, 95, row_h - 6, 4, fill=1, stroke=0)
    
            cells = [
                str(action.get("description", ""))[:52],
                str(action.get("owner", "")).split()[0][:18] if action.get("owner") else "—",
                str(action.get("target_date", ""))[:10],
                "",  # status pill drawn above
                str(action.get("root_cause_addressed", "—"))[:22],
                "✔" if action.get("evidence_b64") else "—",
            ]
            c.setFillColor(HexColor(BLACK))
            c.setFont("Helvetica", 8.5)
            for ci, (x, txt) in enumerate(zip(col_x, cells)):
                if ci == 3:
                    c.setFillColor(HexColor(WHITE) if status != "Open" else HexColor(BLACK))
                    c.setFont("Helvetica-Bold", 7.5)
                    c.drawCentredString(col_x[3] + 48, row_y + 6, status)
                    c.setFillColor(HexColor(BLACK))
                    c.setFont("Helvetica", 8.5)
                else:
                    c.drawString(x + 5, row_y + 6, txt)
            row_y -= row_h
    
        # summary bar
        ot_rate = int(completed / len(actions) * 100) if actions else 0
        _bottom_bar(c,
                    f"✔ {completed} Done  ●  {in_prog} In Progress  ○ {len(actions)-completed-in_prog} Open  ·  On-time: {ot_rate}%",
                    project.get("project_name", "")[:50])
    
    
    # ════════════════════════════════════════════════════════
    # SLIDE 7 – AUDIT SCORE & GAP
    # ════════════════════════════════════════════════════════
    TARGET_RAMP_L = {
        1:12, 2:15, 3:20, 4:27, 5:45, 6:47,
        7:62, 8:64, 9:64, 10:70, 11:85, 12:100
    }
    QUESTIONS_DICT = {
        1:{"text":"Are team members listed?","score":1,"week":1},
        2:{"text":"All members assigned clear roles?","score":1,"week":1},
        3:{"text":"Problem linked to company KPI?","score":4,"week":1},
        4:{"text":"Cost/Benefit chart introduced?","score":4,"week":5},
        5:{"text":"Historical data clearly shown?","score":1,"week":1},
        6:{"text":"Performance indicator target clear?","score":1,"week":2},
        7:{"text":"KPI subdivided into components?","score":1,"week":1},
        8:{"text":"Route & Master Plan visible?","score":2,"week":1},
        9:{"text":"Step targets clear?","score":1,"week":2},
        10:{"text":"Steps subdivided into activities?","score":1,"week":3},
        11:{"text":"Root-cause analysis documented?","score":1,"week":4},
        12:{"text":"Causes verified with data?","score":1,"week":5},
        13:{"text":"Data collection consistent?","score":1,"week":2},
        14:{"text":"Route methods/tools used?","score":2,"week":5},
        15:{"text":"Reoccurrence analysis updated?","score":1,"week":7},
        16:{"text":"Single problem analysis applied?","score":1,"week":7},
        17:{"text":"Countermeasures logical?","score":2,"week":6},
        18:{"text":"Basic conditions restored?","score":3,"week":4},
        19:{"text":"Actions with target dates?","score":5,"week":4},
        20:{"text":"Owner for each action?","score":2,"week":5},
        21:{"text":"Action plan up to date?","score":2,"week":5},
        22:{"text":"Majority of actions on time?","score":3,"week":5},
        23:{"text":"Evidence of implemented actions?","score":4,"week":5},
        24:{"text":"KPI trend positive?","score":15,"week":7},
        25:{"text":"Goal achieved / substantial progress?","score":15,"week":11},
        26:{"text":"Procedures to hold gains?","score":2,"week":10},
        27:{"text":"Monitoring systems in place?","score":2,"week":7},
        28:{"text":"Monitoring devices up-to-date?","score":2,"week":8},
        29:{"text":"OPLs/SOPs for improvements?","score":2,"week":10},
        30:{"text":"Training matrix for OPLs?","score":2,"week":10},
        31:{"text":"CIL audits ≥ 90%?","score":3,"week":5},
        32:{"text":"Workplace well organised (5S)?","score":1,"week":5},
        33:{"text":"Improvements evident?","score":1,"week":5},
        34:{"text":"Methodology understood by team?","score":5,"week":3},
        35:{"text":"Random member can explain board?","score":3,"week":3},
        36:{"text":"Meetings organised & attended?","score":2,"week":1},
    }
    
    
    def _score_traj_fig(total_score, audit_records, cw):
        fig, ax = plt.subplots(figsize=(7, 4), dpi=180)
        fig.patch.set_facecolor("#FAFAFA")
        ax.set_facecolor("#FAFAFA")
    
        tw = sorted(TARGET_RAMP_L.keys())
        tt = [TARGET_RAMP_L[w] for w in tw]
        ax.fill_between(tw, 0, tt, alpha=0.05, color=MGREY)
        ax.plot(tw, tt, "o--", color=MGREY, lw=1.5, ms=4, label="Target ramp")
    
        scores_by_week = {ar["week_number"]: ar.get("total_score", 0)
                          for ar in (audit_records or [])}
        aw = sorted(scores_by_week.keys())
        av = [scores_by_week[w] for w in aw]
        if aw:
            ax.fill_between(aw, 0, av, alpha=0.12, color=DBLUE)
            ax.plot(aw, av, "o-", color=DBLUE, lw=2.5, ms=7, label="Actual", zorder=5)
    
        ax.scatter([cw], [total_score], s=90, color=DBLUE, zorder=7)
        tgt_w = TARGET_RAMP_L.get(cw, 100)
        ax.scatter([cw], [tgt_w], s=70, color=RED, marker="D", zorder=7,
                   label=f"Target W{cw}: {tgt_w}")
        ax.axvline(cw, color=RED, lw=1.5, ls=":", alpha=0.6)
    
        ax.set_xlim(0.5, 12.5)
        ax.set_ylim(0, 108)
        ax.set_xticks(range(1, 13))
        ax.set_xticklabels([f"W{i}" for i in range(1, 13)], fontsize=9)
        ax.set_ylabel("Score / 100", fontsize=9, color=DGREY)
        ax.tick_params(colors=DGREY, labelsize=9)
        ax.legend(fontsize=9, frameon=False, loc="upper left")
        ax.grid(axis="y", color="#EBEBEB", lw=0.5)
        for sp in ["top", "right"]:
            ax.spines[sp].set_visible(False)
        plt.tight_layout(pad=0.5)
        return fig
    
    
    def _slide_audit(c, project, scores, total_score, audit_records, cw):
        _slide_bg(c)
        score_col = GREEN if total_score >= 70 else AMBER if total_score >= 45 else RED
        tgt = TARGET_RAMP_L.get(cw, 100)
        gap_txt = f"+{int(total_score-tgt)} ahead" if total_score >= tgt else f"{int(total_score-tgt)} behind"
        _top_bar(c, "Audit Score",
                 f"Week {cw}  ·  Score: {int(total_score)}/100  ·  Target: {tgt}  ·  {gap_txt}")
    
        DIMS = [
            ("Involvement",   [1,2,34,35,36], 12),
            ("Method",        [8,9,10,11,12,13,14,15,16], 11),
            ("Action Plan",   [17,18,19,20,21,22,23], 21),
            ("Results",       [3,4,5,6,7,24,25], 41),
            ("Stabilisation", [26,27,28,29,30,31,32,33], 15),
        ]
    
        # ── Left: Score ring + trajectory ──
        # Big score circle
        cx, cy, r = 95, SH - 170, 62
        c.setFillColor(HexColor("#E8ECF1"))
        c.circle(cx, cy, r + 4, fill=1, stroke=0)
        c.setFillColor(HexColor(score_col))
        c.circle(cx, cy, r, fill=1, stroke=0)
        c.setFillColor(HexColor(WHITE))
        c.circle(cx, cy, r - 14, fill=1, stroke=0)
        c.setFillColor(HexColor(score_col))
        c.setFont("Helvetica-Bold", 32)
        c.drawCentredString(cx, cy - 10, str(int(total_score)))
        c.setFont("Helvetica", 9)
        c.setFillColor(HexColor(DGREY))
        c.drawCentredString(cx, cy + 20, "SCORE")
        c.drawCentredString(cx, cy - 22, "/ 100")
    
        # Trajectory chart
        fig_t = _score_traj_fig(total_score, audit_records, cw)
        ir_t  = _reader(fig_t)
        c.drawImage(ir_t, 12, 36, width=450, height=SH - 100 - 36,
                    preserveAspectRatio=True)
    
        # ── Right: Dimension cards + gap list ──
        rx = 475
        dim_w = (SW - rx - 16) / len(DIMS)
    
        for i, (dim, qs, possible) in enumerate(DIMS):
            achieved = sum(scores.get(q, 0) for q in qs)
            pct = achieved / possible * 100 if possible > 0 else 0
            dc = GREEN if pct >= 90 else AMBER if pct >= 60 else RED
            dx = rx + i * dim_w
    
            # card
            c.setFillColor(HexColor(LGREY))
            c.roundRect(dx + 2, SH - 200, dim_w - 4, 128, 6, fill=1, stroke=0)
            # colour band top
            c.setFillColor(HexColor(dc))
            c.roundRect(dx + 2, SH - 80, dim_w - 4, 10, 3, fill=1, stroke=0)
    
            # donut-style arc using filled circles
            ddx = dx + dim_w / 2
            ddy = SH - 145
            rr  = 30
            c.setFillColor(HexColor("#DDEAF5"))
            c.circle(ddx, ddy, rr, fill=1, stroke=0)
            c.setFillColor(HexColor(dc))
            c.circle(ddx, ddy, rr, fill=0, stroke=1)
            # inner white
            c.setFillColor(HexColor(WHITE))
            c.circle(ddx, ddy, rr - 8, fill=1, stroke=0)
    
            c.setFillColor(HexColor(dc))
            c.setFont("Helvetica-Bold", 13)
            c.drawCentredString(ddx, ddy - 5, f"{int(pct)}%")
    
            c.setFillColor(HexColor(BLACK))
            c.setFont("Helvetica-Bold", 7.5)
            # wrap dim name
            words = dim.split()
            line = ""
            ty = SH - 192
            for w in words:
                test = (line + " " + w).strip()
                if c.stringWidth(test, "Helvetica-Bold", 7.5) < dim_w - 8:
                    line = test
                else:
                    c.drawCentredString(ddx, ty, line)
                    ty -= 10
                    line = w
            c.drawCentredString(ddx, ty, line)
    
            c.setFillColor(HexColor(DGREY))
            c.setFont("Helvetica", 7)
            c.drawCentredString(ddx, SH - 203, f"{achieved}/{possible} pts")
    
        # Gap register
        gy = SH - 218
        _section_label(c, "Gap Register", rx, gy, w=SW - rx - 16, h=16)
        gy -= 20
    
        WDUE = {4:5,15:7,16:7,24:7,25:11,26:10,28:8,29:10,30:10,31:5}
        gaps = [(qn, QUESTIONS_DICT[qn])
                for qn in QUESTIONS_DICT
                if scores.get(qn, 0) < QUESTIONS_DICT[qn]["score"]
                and cw >= QUESTIONS_DICT[qn]["week"]]
    
        if not gaps:
            c.setFillColor(HexColor(GREEN))
            c.setFont("Helvetica-Bold", 9)
            c.drawString(rx + 6, gy, "✓ No gaps — all due questions met!")
        else:
            for qn, qdata in gaps[:8]:
                # pill
                c.setFillColor(HexColor("#FDECEA"))
                c.roundRect(rx + 2, gy - 3, SW - rx - 18, 14, 3, fill=1, stroke=0)
                c.setFillColor(HexColor(RED))
                c.setFont("Helvetica-Bold", 7.5)
                c.drawString(rx + 6, gy, f"Q{qn}")
                c.setFillColor(HexColor(BLACK))
                c.setFont("Helvetica", 7.5)
                c.drawString(rx + 24, gy, qdata["text"][:62])
                c.setFillColor(HexColor(DGREY))
                c.drawRightString(SW - 18, gy, f"{qdata['score']}pt  due W{qdata['week']}")
                gy -= 16
    
        ar = (audit_records or [None])[-1]
        if ar and ar.get("auditor_notes"):
            _bottom_bar(c, f"Notes: {ar['auditor_notes'][:80]}",
                        f"Audited: {ar.get('audited_by','')}  W{ar.get('week_number','')}")
        else:
            _bottom_bar(c, project.get("project_name","")[:60],
                        f"Generated {date.today().strftime('%d %b %Y')}")
    
    
    # ════════════════════════════════════════════════════════
    # MASTER FUNCTION
    # ════════════════════════════════════════════════════════

    SW, SH = 960, 540
    DBLUE='#0C5595'; RED='#DE201B'; LGREY='#F4F6F8'
    MGREY='#BDC3C7'; DGREY='#566573'; GREEN='#1E8449'
    AMBER='#D68910'; WHITE='#FFFFFF'; BLACK='#1A1A2E'; DKBG='#0A1628'


    # ── Parse logo ──
    _LOGO_B64 = "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAIBAQEBAQIBAQECAgICAgQDAgICAgUEBAMEBgUGBgYFBgYGBwkIBgcJBwYGCAsICQoKCgoKBggLDAsKDAkKCgr/2wBDAQICAgICAgUDAwUKBwYHCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgr/wAARCAF/BVYDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9/KKKKACiiigAooooA"  # truncated placeholder

    logo_reader = _napco_logo_reader_from_b64(_LOGO_B64)

    cw = _get_current_week(project.get("launch_date"))
    kpi = kpi or {}
    stab = stab or {}
    team = team or []
    steps = steps or []
    actions = actions or []
    weekly_updates = weekly_updates or []
    audit_records = audit_records or []

    # KPI current value
    kpi_vals_sorted = sorted(
        [w for w in weekly_updates if w.get("kpi_value") is not None],
        key=lambda x: x["week_number"]
    )
    kpi_baseline = float(kpi.get("baseline_value", 0) or 0)
    kpi_target   = float(kpi.get("target_value", 0)   or 0)
    kpi_current  = float(kpi_vals_sorted[-1]["kpi_value"]) if kpi_vals_sorted else kpi_baseline
    total_gap    = abs(kpi_target - kpi_baseline)
    kpi_progress = abs(kpi_current - kpi_baseline) / total_gap * 100 if total_gap > 0 else 0

    # Stash current value so cover slide can use it
    project = dict(project)
    project["_kpi_current"] = round(kpi_current, 2)

    score_col = GREEN if total_score >= 70 else AMBER if total_score >= 45 else RED

    # Sub-components
    sub_comps = kpi.get("sub_components") or []
    if isinstance(sub_comps, str):
        try: sub_comps = json.loads(sub_comps)
        except: sub_comps = []
    sub_comps = [s for s in sub_comps if isinstance(s, dict)]

    buf = io.BytesIO()
    c = _rl_canvas.Canvas(buf, pagesize=(SW, SH))

    # ── S1: Cover ──
    _slide_cover(c, project, kpi, cw, total_score, score_col, kpi_progress, logo_reader)
    c.showPage()

    # ── S2: Team ──
    _slide_team(c, project, team, kpi, cw)
    c.showPage()

    # ── S3: Gantt ──
    _slide_gantt(c, project, steps, weekly_updates, cw)
    c.showPage()

    # ── S4: KPI Trend ──
    _slide_kpi(c, project, kpi, weekly_updates, cw, kpi_progress)
    c.showPage()

    # ── S5: KAI slides ──
    if sub_comps:
        if len(sub_comps) <= 4:
            _slide_kai_combined(c, project, sub_comps, weekly_updates, kpi, cw)
            c.showPage()
        else:
            for idx, kai in enumerate(sub_comps):
                _slide_kai_single(c, project, kai, weekly_updates, kpi, cw, idx, len(sub_comps))
                c.showPage()

    # ── S6: Actions ──
    _slide_actions(c, project, actions, cw)
    c.showPage()

    # ── S7: Audit ──
    _slide_audit(c, project, scores, total_score, audit_records, cw)
    c.showPage()

    c.save()
    buf.seek(0)
    return buf
