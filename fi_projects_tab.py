# fi_projects_tab.py
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

mpl.rcParams["font.family"] = "DejaVu Sans"

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
    31: {"text": "Are CIL standards clear? Do CIL audits achieve >= 90%?", "score": 3, "week": 5},
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

KPI_TREE = {
    "OEE Improvement": {
        "unit": "%",
        "kais": [
            "Reduce Breakdown Time","Reduce Minor Stoppages","Reduce Setup & Changeover Time",
            "Reduce Planned Maintenance Time","Reduce Speed Losses","Increase Availability Rate",
            "Increase Performance Rate","Increase Quality Rate",
        ]
    },
    "Quality Defect Reduction": {
        "unit": "%",
        "kais": [
            "Reduce Customer Complaints Count","Reduce Internal Defects (PPM)",
            "Reduce Rework Rate","Reduce Scrap Rate","Improve First Pass Yield","Reduce NCR Count",
        ]
    },
    "Waste Reduction": {
        "unit": "%",
        "kais": [
            "Reduce Paper / Board Waste %","Reduce Trim Waste","Reduce Ink & Chemical Waste",
            "Reduce Energy Consumption","Reduce Sheet Waste",
        ]
    },
    "Cost Reduction": {
        "unit": "K EUR",
        "kais": [
            "Reduce Maintenance Spend","Reduce Material Costs","Reduce Labour Overtime",
            "Reduce Energy Costs","Reduce Rework & Scrap Costs",
        ]
    },
    "Safety Improvement": {
        "unit": "Count",
        "kais": [
            "Reduce Near Miss Incidents","Reduce Lost Time Accidents","Improve Safety Audit Score",
            "Increase Near Miss Reporting Rate","Eliminate Unsafe Conditions (Tags)",
        ]
    },
    "Delivery Performance": {
        "unit": "%",
        "kais": [
            "Reduce Order Lead Time","Improve Schedule Adherence",
            "Improve OTIF Rate","Reduce Order Backlog",
        ]
    },
    "5S Score Improvement": {
        "unit": "Score",
        "kais": [
            "Improve Sort (Seiri) Score","Improve Set in Order (Seiton) Score",
            "Improve Shine (Seiso) Score","Improve Standardise (Seiketsu) Score",
            "Improve Sustain (Shitsuke) Score",
        ]
    },
    "Throughput / Productivity": {
        "unit": "MT",
        "kais": [
            "Increase Net Run Time","Increase Average Speed","Reduce Idle Time",
            "Increase Good Boards Production","Improve Capacity Utilisation",
        ]
    },
}

UNITS = ["% (Percent)","MT","SAR","LM","SQM","GSM","Hits","Hits/Hour",
         "LM/Min","BD Time (hrs)","Hours","Mins","Secs","K EUR","Count","Score"]

DEFAULT_COMPANY_KPIS = [
    "Reduce Costs","Improve Customer Satisfaction","Increase OEE",
    "Reduce Waste","Improve Safety","Improve Delivery Performance","Increase Productivity"
]

# ── Helpers ──────────────────────────────────────────────────────────────────

def _get_current_week(launch_date):
    if not launch_date:
        return 1
    if isinstance(launch_date, str):
        try:
            launch_date = date.fromisoformat(launch_date)
        except Exception:
            return 1
    delta = (date.today() - launch_date).days
    return max(1, min(12, math.ceil(delta / 7) if delta > 0 else 1))


def _to_b64(f):
    if f is None:
        return None
    return base64.b64encode(f.getvalue()).decode()


def _fig_to_buf(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    buf.seek(0)
    return buf


def _gantt_chart(steps, weekly_updates, current_week):
    if not steps:
        return None
    n = len(steps)
    fig, ax = plt.subplots(figsize=(14, max(3, n * 0.8 + 1.5)), dpi=150)
    fig.patch.set_facecolor("#FAFAFA")
    ax.set_facecolor("#FAFAFA")
    bar_h = 0.5
    for i, step in enumerate(steps):
        ps = step.get("planned_start_week", 1)
        pe = step.get("planned_end_week", 2)
        ax.barh(i, pe - ps, left=ps - 1, height=bar_h, color="#C8DFF0", alpha=1.0, zorder=2)
        ax.barh(i, pe - ps, left=ps - 1, height=bar_h, color="none", edgecolor="#006394", linewidth=1.2, zorder=3)
        pct = 0
        for wu in weekly_updates:
            sp_raw = wu.get("step_progress") or []
            if isinstance(sp_raw, str):
                try:
                    sp_raw = json.loads(sp_raw)
                except Exception:
                    sp_raw = []
            for sp in sp_raw:
                if not isinstance(sp, dict):
                    continue
                if sp.get("step_id") == str(step.get("id", "")):
                    pct = max(pct, sp.get("pct_complete", 0))
        actual_w = (pe - ps) * pct / 100
        if actual_w > 0:
            ax.barh(i, actual_w, left=ps - 1, height=bar_h, color="#006394", alpha=0.85, zorder=4)
        if pct > 0:
            ax.text(ps - 1 + actual_w / 2, i, f"{int(pct)}%",
                    ha="center", va="center", fontsize=8, color="white", fontweight="bold", zorder=5)
        owner = step.get("owner", "")
        if owner:
            ax.text(pe - 0.05, i + bar_h / 2 + 0.05, owner,
                    ha="right", va="bottom", fontsize=7, color="#555555", zorder=5)
    ax.axvline(current_week - 1, color="#DE201B", linewidth=2, linestyle="--", zorder=6)
    ax.text(current_week - 1 + 0.05, n - 0.3, f"W{current_week}", color="#DE201B", fontsize=8, fontweight="bold", zorder=7)
    for w in range(12):
        ax.axvline(w, color="#DDDDDD", linewidth=0.5, zorder=1)
    ax.set_yticks(range(n))
    ax.set_yticklabels([s.get("step_name", "") for s in steps], fontsize=9, fontweight="bold")
    ax.set_xticks(range(12))
    ax.set_xticklabels([f"W{i+1}" for i in range(12)], fontsize=9)
    ax.set_xlim(0, 12)
    ax.set_ylim(-0.6, n - 0.4)
    ax.invert_yaxis()
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.tick_params(left=False)
    from matplotlib.patches import Patch
    legend_elements = [
        Patch(facecolor="#C8DFF0", edgecolor="#006394", label="Planned"),
        Patch(facecolor="#006394", label="Actual Progress"),
        plt.Line2D([0], [0], color="#DE201B", linewidth=2, linestyle="--", label="Current Week"),
    ]
    ax.legend(handles=legend_elements, loc="lower right", fontsize=8, frameon=True, framealpha=0.9)
    plt.tight_layout(pad=1.5)
    return fig


def _kpi_trend_chart(kpi_data, baseline, target, weeks_data):
    fig, ax = plt.subplots(figsize=(10, 4), dpi=120)
    weeks = [w["week_number"] for w in weeks_data if w.get("kpi_value") is not None]
    values = [w["kpi_value"] for w in weeks_data if w.get("kpi_value") is not None]
    if baseline is not None:
        ax.axhline(baseline, color="#888888", linewidth=1.5, linestyle=":", label=f"Baseline ({baseline})")
    if target is not None:
        ax.axhline(target, color="#27AE60", linewidth=1.5, linestyle="--", label=f"Target ({target})")
    if weeks and values:
        ax.plot(weeks, values, "o-", color="#006394", linewidth=2, markersize=6, label="Actual")
        for w, v in zip(weeks, values):
            ax.annotate(f"{v}", (w, v), textcoords="offset points", xytext=(0, 8), fontsize=8, ha="center")
    ax.set_xlabel("Week")
    ax.set_ylabel(kpi_data.get("unit", "Value") if kpi_data else "Value")
    ax.set_xticks(range(1, 13))
    ax.set_xticklabels([f"W{i}" for i in range(1, 13)], fontsize=8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.legend(fontsize=9, frameon=False)
    plt.tight_layout()
    return fig


def _score_project(project, team, kpi, steps, weekly_updates, actions, stab, audit_records):
    scores = {q: 0 for q in range(1, 37)}
    stab = stab or {}
    team = team or []
    steps = steps or []
    actions = actions or []
    weekly_updates = weekly_updates or []
    audit_records = audit_records or []

    def _safe_date(d):
        try:
            return date.fromisoformat(str(d)[:10])
        except Exception:
            return None

    scores[1] = 1 if len(team) >= 1 else 0
    scores[2] = 1 if team and all(m.get("role") for m in team) else 0
    scores[3] = 4 if (project.get("problem_statement") and project.get("company_kpi_link")) else 0
    scores[4] = 0
    scores[5] = 1 if (kpi and kpi.get("kpi_name") and kpi.get("baseline_value") is not None) else 0
    scores[6] = 1 if (kpi and kpi.get("target_value") is not None and kpi.get("target_date")) else 0
    scores[7] = 1 if (kpi and (kpi.get("sub_components") or project.get("kpi_not_applicable"))) else 0
    scores[8] = 2 if len(steps) >= 2 else 0
    scores[9] = 1 if all(s.get("planned_start_week") and s.get("planned_end_week") for s in steps) else 0

    def _parse_sp(wu):
        sp = wu.get("step_progress") or []
        if isinstance(sp, str):
            try:
                return json.loads(sp)
            except Exception:
                return []
        return sp

    scores[10] = 1 if any(_parse_sp(wu) for wu in weekly_updates) else 0
    scores[11] = 1 if any(wu.get("rca_performed") and wu.get("rca_findings") for wu in weekly_updates) else 0
    scores[12] = 1 if any(wu.get("causes_verified") for wu in weekly_updates) else 0
    scores[13] = 1 if any(wu.get("shifts_covered") and "All" in (wu.get("shifts_covered") or "") for wu in weekly_updates) else 0
    structured = {"5-Why", "Fishbone", "Pareto", "FMEA"}
    scores[14] = 2 if any(wu.get("rca_method") in structured for wu in weekly_updates) else 0
    scores[15] = 1 if any(wu.get("reoccurrence_description") for wu in weekly_updates) else 0
    scores[16] = 1 if any(wu.get("single_problem_analysis") and wu.get("single_problem_notes") for wu in weekly_updates) else 0
    scores[17] = 2 if actions and all(a.get("root_cause_addressed") for a in actions) else 0
    scores[18] = 3 if any(wu.get("basic_conditions_restored") and wu.get("basic_conditions_after_b64") for wu in weekly_updates) else 0
    scores[19] = 5 if actions and any(a.get("target_date") for a in actions) else 0
    scores[20] = 2 if actions and all(a.get("owner") for a in actions) else 0
    current_week = _get_current_week(project.get("launch_date"))
    recent_actions = [a for a in actions if a.get("created_week") == current_week]
    scores[21] = 2 if recent_actions else 0
    past_due = [a for a in actions if a.get("target_date") and _safe_date(a["target_date"]) and _safe_date(a["target_date"]) < date.today()]
    on_time = [a for a in past_due if a.get("status") == "Completed"]
    scores[22] = 3 if (past_due and len(on_time) / len(past_due) >= 0.5) else 0
    scores[23] = 4 if any(a.get("evidence_b64") for a in actions if a.get("status") == "Completed") else 0
    kpi_vals = [wu["kpi_value"] for wu in sorted(weekly_updates, key=lambda x: x["week_number"]) if wu.get("kpi_value") is not None]
    if kpi and kpi.get("target_value") and len(kpi_vals) >= 3:
        baseline = kpi.get("baseline_value", kpi_vals[0])
        target_val = kpi.get("target_value")
        improving = (target_val > baseline and kpi_vals[-1] > kpi_vals[-3]) or (target_val < baseline and kpi_vals[-1] < kpi_vals[-3])
        scores[24] = 15 if improving else 0
    else:
        scores[24] = 0
    if kpi and kpi.get("target_value") and kpi.get("baseline_value") is not None and kpi_vals:
        baseline = kpi.get("baseline_value")
        target_val = kpi.get("target_value")
        current_val = kpi_vals[-1]
        total_gap = abs(target_val - baseline)
        progress = abs(current_val - baseline) / total_gap * 100 if total_gap > 0 else 0
        scores[25] = 15 if progress >= 80 else 0
    else:
        scores[25] = 0
    scores[26] = 2 if stab and stab.get("procedures_created") and stab.get("procedures") else 0
    scores[27] = 2 if stab and stab.get("monitoring_in_place") else 0
    scores[28] = 2 if stab and stab.get("monitoring_active") and stab.get("monitoring_last_update") else 0
    scores[29] = 2 if stab and stab.get("opls") else 0
    scores[30] = 2 if stab and stab.get("training_matrix") else 0
    scores[31] = 3 if stab and stab.get("cil_standards_defined") and stab.get("cil_audit_score") and float(stab.get("cil_audit_score", 0)) >= 90 else 0
    scores[32] = 1 if stab and stab.get("five_s_rating") and int(stab.get("five_s_rating", 0)) >= 3 else 0
    scores[33] = 1 if stab and stab.get("improvements_visible") and stab.get("improvements_photos_b64") else 0
    for ar in audit_records:
        qs = ar.get("question_scores") or {}
        if qs.get("34"):
            scores[34] = 5
        if qs.get("35"):
            scores[35] = 3
    scores.setdefault(34, 0)
    scores.setdefault(35, 0)
    scores[36] = 2 if any(wu.get("meeting_held") for wu in weekly_updates) else 0
    total = sum(scores.values())
    return scores, total


# ── PDF Generator ─────────────────────────────────────────────────────────────

def _generate_project_pdf(project, team, kpi, steps, weekly_updates, actions, stab, audit_records, scores, total_score):
    import tempfile as _tf

    SW, SH = 960, 540

    DBLUE  = "#0C5595"
    RED    = "#DE201B"
    LGREY  = "#F4F6F8"
    MGREY  = "#BDC3C7"
    DGREY  = "#566573"
    GREEN  = "#1E8449"
    AMBER  = "#D68910"
    WHITE  = "#FFFFFF"
    BLACK  = "#1A1A2E"

    kpi   = kpi   or {}
    stab  = stab  or {}
    team  = team  or []
    steps = steps or []
    actions = actions or []
    weekly_updates = weekly_updates or []
    audit_records  = audit_records  or []

    cw = _get_current_week(project.get("launch_date"))
    kpi_baseline = float(kpi.get("baseline_value", 0) or 0)
    kpi_target   = float(kpi.get("target_value",   0) or 0)
    kpi_vals_s   = sorted([w for w in weekly_updates if w.get("kpi_value") is not None], key=lambda x: x["week_number"])
    kpi_current  = float(kpi_vals_s[-1]["kpi_value"]) if kpi_vals_s else kpi_baseline
    total_gap    = abs(kpi_target - kpi_baseline)
    kpi_progress = abs(kpi_current - kpi_baseline) / total_gap * 100 if total_gap > 0 else 0
    score_col    = GREEN if total_score >= 70 else AMBER if total_score >= 45 else RED

    sub_comps = kpi.get("sub_components") or []
    if isinstance(sub_comps, str):
        try:
            sub_comps = json.loads(sub_comps)
        except Exception:
            sub_comps = []
    sub_comps = [s for s in sub_comps if isinstance(s, dict)]

    completed_cnt = sum(1 for a in actions if a.get("status") == "Completed")
    in_prog_cnt   = sum(1 for a in actions if a.get("status") == "In Progress")
    ot_rate       = int(completed_cnt / len(actions) * 100) if actions else 0

    def _fig_reader(fig):
        buf2 = io.BytesIO()
        fig.savefig(buf2, format="png", dpi=150, bbox_inches="tight", facecolor=fig.get_facecolor())
        buf2.seek(0)
        plt.close(fig)
        return ImageReader(buf2)

    def _slide_bg(c):
        c.setFillColor(HexColor(WHITE))
        c.rect(0, 0, SW, SH, fill=1, stroke=0)

    def _top_bar(c, title, subtitle=""):
        c.setFillColor(HexColor(DBLUE))
        c.rect(0, SH - 55, SW, 55, fill=1, stroke=0)
        c.setFillColor(HexColor(RED))
        c.rect(0, SH - 55, 7, 55, fill=1, stroke=0)
        c.setFillColor(HexColor(WHITE))
        c.setFont("Helvetica-Bold", 20)
        c.drawString(22, SH - 34, title)
        if subtitle:
            c.setFont("Helvetica", 9)
            c.setFillColor(HexColor(MGREY))
            c.drawString(22, SH - 50, subtitle)

    def _bottom_bar(c, left="", right=""):
        c.setFillColor(HexColor(DBLUE))
        c.rect(0, 0, SW, 24, fill=1, stroke=0)
        c.setFont("Helvetica", 8)
        c.setFillColor(HexColor(WHITE))
        if left:
            c.drawString(16, 7, left)
        if right:
            c.drawRightString(SW - 16, 7, right)

    # ── SLIDE 1: COVER ────────────────────────────────────────────────────────
    def slide_cover(c):
        _slide_bg(c)

        # dark header band
        c.setFillColor(HexColor(DBLUE))
        c.rect(0, SH - 155, SW, 155, fill=1, stroke=0)
        c.setFillColor(HexColor(RED))
        c.rect(0, SH - 155, 8, 155, fill=1, stroke=0)

        # logo placeholder area (top-left of header)
        # We skip the logo here to avoid base64 issues - project name takes its place
        c.setFillColor(HexColor(WHITE))
        c.setFont("Helvetica-Bold", 26)
        pname = project.get("project_name", "Unnamed Project")
        # wrap if too long
        if len(pname) > 50:
            pname = pname[:50] + "..."
        c.drawString(24, SH - 55, pname)
        c.setFont("Helvetica", 12)
        c.setFillColor(HexColor(MGREY))
        c.drawString(24, SH - 80, f"{project.get('target_area', '')}   |   Week {cw} of 12")
        c.setFont("Helvetica", 10)
        c.drawString(24, SH - 100, f"Launch: {str(project.get('launch_date', ''))[:10]}   |   Target: {str(project.get('expected_completion_date', ''))[:10]}")

        # score ring
        cx, cy, r = SW - 80, SH - 85, 55
        c.setFillColor(HexColor("#1A2A45"))
        c.circle(cx, cy, r + 5, fill=1, stroke=0)
        c.setFillColor(HexColor(score_col))
        c.circle(cx, cy, r, fill=1, stroke=0)
        c.setFillColor(HexColor(DBLUE))
        c.circle(cx, cy, r - 13, fill=1, stroke=0)
        c.setFillColor(HexColor(WHITE))
        c.setFont("Helvetica-Bold", 28)
        c.drawCentredString(cx, cy - 9, str(int(total_score)))
        c.setFont("Helvetica", 8)
        c.setFillColor(HexColor(MGREY))
        c.drawCentredString(cx, cy + 22, "SCORE")
        c.drawCentredString(cx, cy - 22, "/ 100")

        # KPI badges row - 4 equal badges full width
        pad = 24
        bw = (SW - pad * 2 - 30) / 4
        bh = 82
        by = SH - 265
        badge_data = [
            ("BASELINE",  f"{kpi_baseline} {kpi.get('unit', '')}", "Start",          "#17375E"),
            ("CURRENT",   f"{kpi_current:.1f} {kpi.get('unit', '')}", f"Week {cw}", DBLUE),
            ("TARGET",    f"{kpi_target} {kpi.get('unit', '')}", "Goal",             "#145A32"),
            ("PROGRESS",  f"{kpi_progress:.0f}%", "Toward target",
             GREEN if kpi_progress >= 80 else AMBER if kpi_progress >= 40 else RED),
        ]
        for i, (lbl, val, sub, bg) in enumerate(badge_data):
            bx = pad + i * (bw + 10)
            c.setFillColor(HexColor(bg))
            c.roundRect(bx, by, bw, bh, 7, fill=1, stroke=0)
            c.setFillColor(HexColor(WHITE))
            c.setFont("Helvetica", 7)
            c.drawCentredString(bx + bw / 2, by + bh - 12, lbl)
            c.setFont("Helvetica-Bold", 18 if len(str(val)) <= 8 else 13)
            c.drawCentredString(bx + bw / 2, by + bh / 2 - 5, str(val))
            c.setFont("Helvetica", 7)
            c.setFillColor(HexColor(MGREY))
            c.drawCentredString(bx + bw / 2, by + 7, sub)

        # progress bar
        bar_y = by - 20
        bar_w = SW - pad * 2
        c.setFillColor(HexColor(LGREY))
        c.roundRect(pad, bar_y, bar_w, 9, 4, fill=1, stroke=0)
        filled = bar_w * min(kpi_progress / 100, 1)
        if filled > 0:
            bc = GREEN if kpi_progress >= 80 else AMBER if kpi_progress >= 40 else RED
            c.setFillColor(HexColor(bc))
            c.roundRect(pad, bar_y, filled, 9, 4, fill=1, stroke=0)
        c.setFillColor(HexColor(DGREY))
        c.setFont("Helvetica", 7)
        c.drawRightString(pad + bar_w, bar_y - 8, f"{kpi_progress:.0f}% complete")

        # problem statement
        ps_y = bar_y - 26
        c.setFillColor(HexColor(DBLUE))
        c.setFont("Helvetica-Bold", 8)
        c.drawString(pad, ps_y, "PROBLEM STATEMENT")
        ps_y -= 14
        c.setFillColor(HexColor(BLACK))
        c.setFont("Helvetica", 9)
        prob = project.get("problem_statement", "")
        words = prob.split()
        line = ""
        for w in words:
            test = (line + " " + w).strip()
            if c.stringWidth(test, "Helvetica", 9) < SW - pad * 2:
                line = test
            else:
                c.drawString(pad, ps_y, line)
                ps_y -= 13
                line = w
                if ps_y < 30:
                    break
        if line and ps_y > 30:
            c.drawString(pad, ps_y, line)

        _bottom_bar(c,
                    f"Company KPI: {project.get('company_kpi_link', '')}",
                    f"Generated {date.today().strftime('%d %B %Y')}")

    # ── SLIDE 2: TEAM ─────────────────────────────────────────────────────────
    def slide_team(c):
        _slide_bg(c)
        _top_bar(c, "Project Team", f"{project.get('target_area', '')}   |   Week {cw} of 12")

        # build owner map from sub_comps
        owner_map = {}
        for s in sub_comps:
            ow = s.get("owner", "")
            if ow:
                owner_map.setdefault(ow, []).append(
                    f"{s.get('name', '')} ({s.get('baseline', '')} to {s.get('target', '')} {s.get('unit', '')})"
                )

        PAD = 20
        col_x = [PAD, 200, 310, 460, 680]
        col_w = [178, 108, 148, 218, 240]
        hdrs  = ["NAME", "ROLE", "DEPARTMENT", "KAI TARGET", "CONTRIBUTION"]
        row_h = 26
        hdr_y = SH - 78

        # header row
        c.setFillColor(HexColor(DBLUE))
        c.rect(PAD, hdr_y - 2, SW - PAD * 2, row_h, fill=1, stroke=0)
        c.setFillColor(HexColor(WHITE))
        c.setFont("Helvetica-Bold", 8.5)
        for x, hdr in zip(col_x, hdrs):
            c.drawString(x + 5, hdr_y + 8, hdr)

        row_y = hdr_y - row_h
        for ti, m in enumerate(team[:14]):
            bg = LGREY if ti % 2 == 0 else WHITE
            c.setFillColor(HexColor(bg))
            c.rect(PAD, row_y - 2, SW - PAD * 2, row_h, fill=1, stroke=0)
            mname = m.get("member_name", "")
            kai_str = "; ".join(owner_map.get(mname, ["—"]))
            vals = [
                mname,
                m.get("role", ""),
                m.get("department", ""),
                kai_str[:40],
                (m.get("contribution_target", "") or "—")[:30],
            ]
            c.setFillColor(HexColor(BLACK))
            c.setFont("Helvetica", 8.5)
            for xi, (x, val) in enumerate(zip(col_x, vals)):
                txt = str(val)
                while c.stringWidth(txt, "Helvetica", 8.5) > col_w[xi] - 8 and len(txt) > 3:
                    txt = txt[:-2] + "…"
                c.drawString(x + 5, row_y + 7, txt)
            row_y -= row_h

        # KPI summary box if space remains
        if kpi:
            bx = SW - 300
            by2 = 32
            bh2 = row_y - 32
            if bh2 > 60:
                c.setFillColor(HexColor(LGREY))
                c.roundRect(bx, by2, 280, bh2, 6, fill=1, stroke=0)
                c.setFillColor(HexColor(DBLUE))
                c.setFont("Helvetica-Bold", 9)
                c.drawString(bx + 10, by2 + bh2 - 18, "KPI SUMMARY")
                yy = by2 + bh2 - 36
                c.setFont("Helvetica-Bold", 11)
                c.setFillColor(HexColor(DBLUE))
                c.drawString(bx + 10, yy, kpi.get("kpi_name", "")[:30])
                yy -= 18
                for lbl, val in [
                    ("Baseline", f"{kpi.get('baseline_value', '')} {kpi.get('unit', '')}"),
                    ("Target",   f"{kpi.get('target_value', '')} {kpi.get('unit', '')}"),
                    ("Due Date", str(kpi.get("target_date", ""))[:10]),
                ]:
                    c.setFont("Helvetica-Bold", 8)
                    c.setFillColor(HexColor(DGREY))
                    c.drawString(bx + 10, yy, lbl + ":")
                    c.setFont("Helvetica", 9)
                    c.setFillColor(HexColor(BLACK))
                    c.drawString(bx + 80, yy, str(val))
                    yy -= 16

        _bottom_bar(c, f"{len(team)} team members", project.get("project_name", "")[:60])

    # ── SLIDE 3: MASTER PLAN ──────────────────────────────────────────────────
    def slide_gantt(c):
        _slide_bg(c)
        _top_bar(c, "Master Plan", f"Week {cw} of 12   |   {len(steps)} phases defined")
        if steps:
            fig = _gantt_chart(steps, weekly_updates, cw)
            if fig:
                ir = _fig_reader(fig)
                c.drawImage(ir, 12, 30, width=SW - 24, height=SH - 92, preserveAspectRatio=True)
        else:
            c.setFillColor(HexColor(MGREY))
            c.setFont("Helvetica", 14)
            c.drawCentredString(SW / 2, SH / 2, "No steps defined yet.")
        _bottom_bar(c, project.get("project_name", "")[:60], f"Generated {date.today().strftime('%d %b %Y')}")

    # ── SLIDE 4: KPI TREND ────────────────────────────────────────────────────
    def slide_kpi(c):
        _slide_bg(c)
        _top_bar(c, f"KPI Trend — {kpi.get('kpi_name', '')}", f"Week {cw}   |   Progress: {kpi_progress:.0f}%")
        fig, ax = plt.subplots(figsize=(13, 5), dpi=150)
        fig.patch.set_facecolor("#FAFAFA")
        ax.set_facecolor("#FAFAFA")
        weeks = [w["week_number"] for w in kpi_vals_s]
        values = [float(w["kpi_value"]) for w in kpi_vals_s]
        ax.axhline(kpi_baseline, color="#AAAAAA", linewidth=1.5, linestyle=":", label=f"Baseline {kpi_baseline}")
        ax.axhline(kpi_target,   color="#27AE60", linewidth=2.0, linestyle="--", label=f"Target {kpi_target}")
        if weeks:
            ax.fill_between(weeks, kpi_baseline, values, alpha=0.12, color="#006394")
            ax.plot(weeks, values, "o-", color="#006394", linewidth=2.5, markersize=8, zorder=5, label="Actual")
            for w, v in zip(weeks, values):
                ax.annotate(f"{v:.1f}", (w, v), textcoords="offset points", xytext=(0, 10),
                            fontsize=9, ha="center", color="#006394", fontweight="bold")
        ax.axvline(cw, color="#DE201B", linewidth=1.5, linestyle=":", alpha=0.7, label=f"Now W{cw}")
        ax.set_xlim(0.5, 12.5)
        mn = min(values + [kpi_baseline]) - 3 if values else 0
        mx = max(values + [kpi_target])   + 5 if values else 100
        ax.set_ylim(mn, mx)
        ax.set_xticks(range(1, 13))
        ax.set_xticklabels([f"W{i}" for i in range(1, 13)], fontsize=10)
        ax.set_ylabel(f"{kpi.get('kpi_name', '')} ({kpi.get('unit', '')})", fontsize=10, color="#566573")
        ax.tick_params(colors="#566573", labelsize=9)
        ax.legend(fontsize=9, frameon=False, loc="upper left")
        ax.grid(axis="y", color="#EEEEEE", linewidth=0.6)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        plt.tight_layout(pad=0.5)
        ir = _fig_reader(fig)
        c.drawImage(ir, 12, 30, width=SW - 24, height=SH - 92, preserveAspectRatio=True)
        _bottom_bar(c, f"Baseline: {kpi_baseline}  →  Target: {kpi_target} {kpi.get('unit', '')}", project.get("project_name", "")[:50])

    # ── SLIDE 5: KAI PROGRESS ─────────────────────────────────────────────────
    def slide_kai(c):
        _slide_bg(c)
        _top_bar(c, "KAI Progress", f"Week {cw}   |   {len(sub_comps)} Key Activity Indicators")

        if not sub_comps:
            c.setFillColor(HexColor(MGREY))
            c.setFont("Helvetica", 14)
            c.drawCentredString(SW / 2, SH / 2, "No KAIs defined yet.")
            _bottom_bar(c)
            return

        # get latest actual per KAI
        kai_actuals = {}
        for wu in sorted(weekly_updates, key=lambda x: x.get("week_number", 0)):
            notes = wu.get("kpi_notes") or ""
            try:
                nd = json.loads(notes) if notes.startswith("{") else {}
                for k, v in nd.items():
                    if isinstance(v, dict):
                        kai_actuals[k] = float(v.get("value", 0))
            except Exception:
                pass

        n = len(sub_comps)
        cols = min(n, 2)
        rows = math.ceil(n / cols)
        cell_w = (SW - 40) / cols
        cell_h = (SH - 85) / rows
        pad_c = 8

        for idx, kai in enumerate(sub_comps):
            col_i = idx % cols
            row_i = idx // cols
            bx = 20 + col_i * cell_w
            by = SH - 85 - (row_i + 1) * cell_h + pad_c

            kn = kai.get("name", "")
            b  = float(kai.get("baseline", 0) or 0)
            t  = float(kai.get("target",   0) or 0)
            cur = kai_actuals.get(kn, b)
            gap = abs(t - b)
            pct = abs(cur - b) / gap * 100 if gap > 0 else 0
            pct = min(100, max(0, pct))
            bc  = GREEN if pct >= 80 else AMBER if pct >= 40 else RED

            # card background
            c.setFillColor(HexColor(LGREY))
            c.roundRect(bx + 4, by, cell_w - 8, cell_h - pad_c, 8, fill=1, stroke=0)

            # top accent strip
            c.setFillColor(HexColor(bc))
            c.roundRect(bx + 4, by + cell_h - pad_c - 10, cell_w - 8, 10, 4, fill=1, stroke=0)

            # KAI name
            c.setFont("Helvetica-Bold", 10)
            c.setFillColor(HexColor(DBLUE))
            c.drawString(bx + 14, by + cell_h - pad_c - 24, kn[:45])

            # owner + baseline/target line
            c.setFont("Helvetica", 8)
            c.setFillColor(HexColor(DGREY))
            c.drawString(bx + 14, by + cell_h - pad_c - 36,
                         f"Owner: {kai.get('owner', '—')}   Baseline: {b} → Target: {t} {kai.get('unit', '')}")

            # progress bar
            bar_x = bx + 14
            bar_bw = cell_w - 36
            bar_y2 = by + cell_h - pad_c - 52
            c.setFillColor(HexColor(MGREY))
            c.roundRect(bar_x, bar_y2, bar_bw, 8, 4, fill=1, stroke=0)
            if pct > 0:
                c.setFillColor(HexColor(bc))
                c.roundRect(bar_x, bar_y2, bar_bw * pct / 100, 8, 4, fill=1, stroke=0)
            c.setFont("Helvetica-Bold", 9)
            c.setFillColor(HexColor(bc))
            c.drawString(bar_x + bar_bw + 6, bar_y2, f"{pct:.0f}%")

            # current value badge
            c.setFillColor(HexColor(bc))
            c.roundRect(bx + 14, by + 8, 80, 28, 5, fill=1, stroke=0)
            c.setFillColor(HexColor(WHITE))
            c.setFont("Helvetica-Bold", 14)
            c.drawCentredString(bx + 54, by + 18, f"{cur:.1f}")
            c.setFont("Helvetica", 7)
            c.drawCentredString(bx + 54, by + 9, "Current")

        _bottom_bar(c, f"KPI: {kpi.get('kpi_name', '')}", project.get("project_name", "")[:50])

    # ── SLIDE 6: ACTION PLAN ──────────────────────────────────────────────────
    def slide_actions(c):
        _slide_bg(c)
        _top_bar(c, "Action Plan",
                 f"{len(actions)} actions   |   {completed_cnt} completed   |   {in_prog_cnt} in progress   |   On-time: {ot_rate}%")

        STATUS_COL = {"Completed": GREEN, "In Progress": DBLUE, "Open": MGREY, "Overdue": RED}
        PAD = 16
        col_x = [PAD, 330, 460, 556, 668, 790]
        col_w = [312, 128, 94,  110, 120, 130]
        hdrs  = ["ACTION", "OWNER", "DUE DATE", "STATUS", "ROOT CAUSE", "EVIDENCE"]
        row_h = 25
        hdr_y = SH - 76

        c.setFillColor(HexColor(DBLUE))
        c.rect(PAD, hdr_y - 2, SW - PAD * 2, row_h, fill=1, stroke=0)
        c.setFillColor(HexColor(WHITE))
        c.setFont("Helvetica-Bold", 8.5)
        for x, hdr in zip(col_x, hdrs):
            c.drawString(x + 5, hdr_y + 8, hdr)

        row_y = hdr_y - row_h
        for ai, action in enumerate(actions[:15]):
            bg = LGREY if ai % 2 == 0 else WHITE
            c.setFillColor(HexColor(bg))
            c.rect(PAD, row_y - 2, SW - PAD * 2, row_h, fill=1, stroke=0)
            status = action.get("status", "Open")
            sc = STATUS_COL.get(status, MGREY)
            c.setFillColor(HexColor(sc))
            c.roundRect(col_x[3] + 3, row_y, 96, row_h - 5, 4, fill=1, stroke=0)
            cells = [
                str(action.get("description", ""))[:50],
                (str(action.get("owner", "")).split()[0] if action.get("owner") else "—")[:16],
                str(action.get("target_date", ""))[:10],
                "",
                str(action.get("root_cause_addressed", "—"))[:20],
                "Yes" if action.get("evidence_b64") else "—",
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

        _bottom_bar(c,
                    f"Completed: {completed_cnt}   In Progress: {in_prog_cnt}   Open: {len(actions)-completed_cnt-in_prog_cnt}",
                    project.get("project_name", "")[:50])

    # ── SLIDE 7: AUDIT SCORE ──────────────────────────────────────────────────
    def slide_audit(c):
        _slide_bg(c)
        tgt_this_week = TARGET_RAMP.get(cw, 100)
        gap_txt = f"+{int(total_score - tgt_this_week)} ahead" if total_score >= tgt_this_week else f"{int(total_score - tgt_this_week)} behind"
        _top_bar(c, "Audit Score", f"Week {cw}   |   Score: {int(total_score)}/100   |   Target: {tgt_this_week}   |   {gap_txt}")

        DIMS = [
            ("Involvement",   [1, 2, 34, 35, 36], 12),
            ("Method",        [8, 9, 10, 11, 12, 13, 14, 15, 16], 11),
            ("Action Plan",   [17, 18, 19, 20, 21, 22, 23], 21),
            ("Results",       [3, 4, 5, 6, 7, 24, 25], 41),
            ("Stabilisation", [26, 27, 28, 29, 30, 31, 32, 33], 15),
        ]

        # Score trajectory chart (left side)
        fig, ax = plt.subplots(figsize=(6.5, 3.8), dpi=150)
        fig.patch.set_facecolor("#FAFAFA")
        ax.set_facecolor("#FAFAFA")
        tw = sorted(TARGET_RAMP.keys())
        tt = [TARGET_RAMP[w] for w in tw]
        ax.fill_between(tw, 0, tt, alpha=0.05, color="#AAAAAA")
        ax.plot(tw, tt, "o--", color="#AAAAAA", lw=1.5, ms=4, label="Target ramp")
        scores_by_week = {ar["week_number"]: ar.get("total_score", 0) for ar in audit_records}
        aw = sorted(scores_by_week.keys())
        av = [scores_by_week[w] for w in aw]
        if aw:
            ax.fill_between(aw, 0, av, alpha=0.12, color="#006394")
            ax.plot(aw, av, "o-", color="#006394", lw=2.5, ms=7, label="Actual", zorder=5)
        ax.scatter([cw], [total_score], s=80, color="#006394", zorder=7)
        ax.scatter([cw], [tgt_this_week], s=60, color="#DE201B", marker="D", zorder=7, label=f"Target W{cw}")
        ax.axvline(cw, color="#DE201B", lw=1.5, ls=":", alpha=0.6)
        ax.set_xlim(0.5, 12.5)
        ax.set_ylim(0, 108)
        ax.set_xticks(range(1, 13))
        ax.set_xticklabels([f"W{i}" for i in range(1, 13)], fontsize=8)
        ax.set_ylabel("Score / 100", fontsize=9, color="#566573")
        ax.legend(fontsize=8, frameon=False, loc="upper left")
        ax.grid(axis="y", color="#EEEEEE", lw=0.5)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        plt.tight_layout(pad=0.4)
        ir = _fig_reader(fig)
        c.drawImage(ir, 12, 36, width=460, height=SH - 100, preserveAspectRatio=True)

        # Dimension cards (right side)
        rx = 482
        dw = (SW - rx - 14) / len(DIMS)
        for i, (dim, qs, possible) in enumerate(DIMS):
            achieved = sum(scores.get(q, 0) for q in qs)
            pct = achieved / possible * 100 if possible > 0 else 0
            dc = GREEN if pct >= 90 else AMBER if pct >= 60 else RED
            dx = rx + i * dw
            c.setFillColor(HexColor(LGREY))
            c.roundRect(dx + 2, SH - 200, dw - 4, 128, 6, fill=1, stroke=0)
            c.setFillColor(HexColor(dc))
            c.roundRect(dx + 2, SH - 82, dw - 4, 10, 3, fill=1, stroke=0)
            ddx = dx + dw / 2
            c.setFillColor(HexColor("#DDEAF5"))
            c.circle(ddx, SH - 148, 28, fill=1, stroke=0)
            c.setFillColor(HexColor(WHITE))
            c.circle(ddx, SH - 148, 20, fill=1, stroke=0)
            c.setFillColor(HexColor(dc))
            c.setFont("Helvetica-Bold", 12)
            c.drawCentredString(ddx, SH - 153, f"{int(pct)}%")
            c.setFillColor(HexColor(BLACK))
            c.setFont("Helvetica-Bold", 7)
            c.drawCentredString(ddx, SH - 188, dim[:12])
            c.setFillColor(HexColor(DGREY))
            c.setFont("Helvetica", 6.5)
            c.drawCentredString(ddx, SH - 198, f"{achieved}/{possible}")

        # Gap register
        gy = SH - 218
        c.setFillColor(HexColor(DBLUE))
        c.setFont("Helvetica-Bold", 8)
        c.drawString(rx, gy, "GAP REGISTER")
        gy -= 14
        gaps = [(qn, QUESTIONS[qn]) for qn in QUESTIONS
                if scores.get(qn, 0) < QUESTIONS[qn]["score"] and cw >= QUESTIONS[qn]["week"]]
        if not gaps:
            c.setFillColor(HexColor(GREEN))
            c.setFont("Helvetica-Bold", 9)
            c.drawString(rx, gy, "All due questions met!")
        else:
            for qn, qdata in gaps[:9]:
                c.setFillColor(HexColor("#FDECEA"))
                c.roundRect(rx, gy - 2, SW - rx - 14, 13, 3, fill=1, stroke=0)
                c.setFillColor(HexColor(RED))
                c.setFont("Helvetica-Bold", 7.5)
                c.drawString(rx + 4, gy, f"Q{qn}")
                c.setFillColor(HexColor(BLACK))
                c.setFont("Helvetica", 7.5)
                c.drawString(rx + 22, gy, qdata["text"][:60])
                c.setFillColor(HexColor(DGREY))
                c.drawRightString(SW - 18, gy, f"{qdata['score']}pt")
                gy -= 14

        ar = audit_records[-1] if audit_records else None
        if ar and ar.get("auditor_notes"):
            _bottom_bar(c, f"Notes: {ar['auditor_notes'][:80]}", f"Audited by: {ar.get('audited_by', '')} W{ar.get('week_number', '')}")
        else:
            _bottom_bar(c, project.get("project_name", "")[:60], f"Generated {date.today().strftime('%d %b %Y')}")

    # ── BUILD PDF ─────────────────────────────────────────────────────────────
    buf = io.BytesIO()
    c = _rl_canvas.Canvas(buf, pagesize=(SW, SH))

    slide_cover(c);  c.showPage()
    slide_team(c);   c.showPage()
    slide_gantt(c);  c.showPage()
    slide_kpi(c);    c.showPage()
    if sub_comps:
        slide_kai(c); c.showPage()
    slide_actions(c); c.showPage()
    slide_audit(c);   c.showPage()

    c.save()
    buf.seek(0)
    return buf


# ════════════════════════════════════════════════════════════════════════════
# MAIN RENDER FUNCTION
# ════════════════════════════════════════════════════════════════════════════

def render_fi_projects_tab(supabase, role, pillar, name):
    can_create = (role == "plant_manager") or (role == "pillar_leader" and pillar == "FI")
    is_auditor = role in ["plant_manager", "pillar_leader"]

    st.markdown("### FI Projects")

    try:
        proj_resp = supabase.table("fi_projects").select("*").order("created_at", desc=True).execute()
        all_projects = proj_resp.data or []
    except Exception as e:
        st.error(f"Could not load projects: {e}")
        return

    col_sel, col_new = st.columns([4, 1])
    if not all_projects:
        st.info("No FI projects yet. Create one to get started.")
        selected_project = None
    else:
        proj_names = {
            p["id"]: f"{p['project_name']} (Started: {str(p.get('launch_date', ''))[:10]})"
            for p in all_projects
        }
        sel_id = col_sel.selectbox("Select Project", list(proj_names.keys()),
                                   format_func=lambda x: proj_names[x], key="fi_proj_select")
        selected_project = next((p for p in all_projects if p["id"] == sel_id), None)

    if can_create:
        if col_new.button("New Project", key="fi_new_proj"):
            st.session_state["fi_creating_project"] = True

    if st.session_state.get("fi_creating_project") and can_create:
        with st.expander("Create New Project", expanded=True):
            with st.form("fi_new_proj_form"):
                np1, np2 = st.columns(2)
                new_name = np1.text_input("Project Name *")
                new_section = np2.selectbox("Section / Area *", list(PLANT_SECTIONS.keys()), key="fi_new_section_form")
                _nm = PLANT_SECTIONS.get(new_section, [])
                if _nm:
                    new_machine = st.selectbox("Machine", ["— All / General —"] + _nm, key="fi_new_machine_form")
                    new_area = f"{new_section} — {new_machine}" if new_machine != "— All / General —" else new_section
                else:
                    new_area = new_section
                new_problem = st.text_area("Problem Statement *", height=80)
                nd1, nd2 = st.columns(2)
                new_launch = nd1.date_input("Launch Date", value=date.today())
                new_expected = nd2.date_input("Expected Completion", value=date.today() + timedelta(weeks=12))
                new_kpi_link = st.text_input("Which company KPI does this project support?")
                if st.form_submit_button("Create Project", type="primary"):
                    if not new_name or not new_problem:
                        st.error("Project Name and Problem Statement are required.")
                    else:
                        supabase.table("fi_projects").insert({
                            "project_name": new_name,
                            "problem_statement": new_problem,
                            "target_area": new_area,
                            "launch_date": str(new_launch),
                            "expected_completion_date": str(new_expected),
                            "company_kpi_link": new_kpi_link,
                            "created_by": name,
                        }).execute()
                        st.success("Project created!")
                        st.session_state["fi_creating_project"] = False
                        st.rerun()
            if st.button("Cancel", key="fi_cancel_new"):
                st.session_state["fi_creating_project"] = False
                st.rerun()

    if not selected_project:
        return

    pid = selected_project["id"]
    current_week = _get_current_week(selected_project.get("launch_date"))

    try:
        team       = supabase.table("fi_project_team").select("*").eq("project_id", pid).execute().data or []
        kpi_rows   = supabase.table("fi_project_kpi").select("*").eq("project_id", pid).execute().data or []
        kpi        = kpi_rows[0] if kpi_rows else None
        steps      = supabase.table("fi_project_steps").select("*").eq("project_id", pid).order("sort_order").execute().data or []
        cost_rows  = supabase.table("fi_project_cost").select("*").eq("project_id", pid).execute().data or []
        cost       = cost_rows[0] if cost_rows else None
        wu_rows    = supabase.table("fi_weekly_updates").select("*").eq("project_id", pid).order("week_number").execute().data or []
        actions    = supabase.table("fi_actions").select("*").eq("project_id", pid).order("created_at").execute().data or []
        stab_rows  = supabase.table("fi_stabilisation").select("*").eq("project_id", pid).execute().data or []
        stab       = stab_rows[0] if stab_rows else None
        audit_records = supabase.table("fi_audit_records").select("*").eq("project_id", pid).order("week_number").execute().data or []
    except Exception as e:
        st.error(f"Error loading project data: {e}")
        return

    wu_by_week = {w["week_number"]: w for w in wu_rows}

    st.markdown(
        f"**Project:** {selected_project['project_name']} &nbsp;|&nbsp; "
        f"**Area:** {selected_project.get('target_area', '')} &nbsp;|&nbsp; "
        f"**Current Week:** <span style='background:#006394;color:white;padding:2px 10px;border-radius:10px;'>W{current_week}</span>",
        unsafe_allow_html=True
    )

    audit_tabs = ["Project Setup", "Weekly Update", "Results", "Stabilisation"]
    if is_auditor:
        audit_tabs.append("Audit View")
    subtabs = st.tabs(audit_tabs)

    # ════════════════════════════════
    # TAB 1 — PROJECT SETUP
    # ════════════════════════════════
    with subtabs[0]:
        st.markdown("#### Project Setup")

        with st.expander("Project Identity", expanded=True):
            _cur_area = selected_project.get("target_area", "") or ""
            _cur_section = _cur_area.split(" — ")[0] if " — " in _cur_area else list(PLANT_SECTIONS.keys())[0]
            _area_cols = st.columns(2)
            setup_section = _area_cols[0].selectbox(
                "Section / Area *", list(PLANT_SECTIONS.keys()),
                index=list(PLANT_SECTIONS.keys()).index(_cur_section) if _cur_section in PLANT_SECTIONS else 0,
                key="fi_setup_section"
            )
            _machines = PLANT_SECTIONS.get(setup_section, [])
            _cur_machine = _cur_area.split(" — ")[1] if " — " in _cur_area else ""
            if _machines:
                _mach_opts = ["— All / General —"] + _machines
                _mach_idx  = (_machines.index(_cur_machine) + 1) if _cur_machine in _machines else 0
                setup_machine = _area_cols[1].selectbox("Machine", _mach_opts, index=_mach_idx, key="fi_setup_machine")
                setup_area = f"{setup_section} — {setup_machine}" if setup_machine != "— All / General —" else setup_section
            else:
                _area_cols[1].caption("No sub-machines for this section")
                setup_area = setup_section

            with st.form("fi_setup_identity"):
                setup_name = st.text_input("Project Name *", value=selected_project.get("project_name", ""))
                st.info(f"Target Area: **{setup_area}**")
                setup_problem = st.text_area("Problem Statement *", value=selected_project.get("problem_statement", ""), height=80)
                d1, d2 = st.columns(2)
                setup_launch = d1.date_input("Launch Date",
                    value=date.fromisoformat(str(selected_project.get("launch_date", date.today()))[:10]))
                _auto_end = date.fromisoformat(str(selected_project.get("launch_date", date.today()))[:10]) + timedelta(weeks=12)
                setup_end = d2.date_input("Expected Completion",
                    value=date.fromisoformat(str(selected_project.get("expected_completion_date", _auto_end))[:10]))
                _company_kpis = DEFAULT_COMPANY_KPIS.copy()
                try:
                    _custom = supabase.table("fi_company_kpis").select("kpi_name").execute().data or []
                    _company_kpis += [k["kpi_name"] for k in _custom if k.get("kpi_name") not in _company_kpis]
                except Exception:
                    pass
                _cur_kpi_link = selected_project.get("company_kpi_link", "")
                _kpi_idx = _company_kpis.index(_cur_kpi_link) if _cur_kpi_link in _company_kpis else 0
                setup_kpi_link = st.selectbox("Which company KPI does this project support?", _company_kpis, index=_kpi_idx)
                setup_kpi_imp = st.text_area("How does this project impact that KPI?",
                    value=selected_project.get("kpi_impact", ""), height=60)
                if st.form_submit_button("Save Project Identity"):
                    supabase.table("fi_projects").update({
                        "project_name": setup_name, "target_area": setup_area,
                        "problem_statement": setup_problem,
                        "launch_date": str(setup_launch),
                        "expected_completion_date": str(setup_end),
                        "company_kpi_link": setup_kpi_link, "kpi_impact": setup_kpi_imp,
                    }).eq("id", pid).execute()
                    st.success("Saved")
                    st.rerun()

        with st.expander("Team Members", expanded=True):
            if team:
                _subs_list = []
                if kpi:
                    _raw = kpi.get("sub_components") or []
                    if isinstance(_raw, str):
                        try: _raw = json.loads(_raw)
                        except: _raw = []
                    _subs_list = [s for s in _raw if isinstance(s, dict)]
                _owner_map2 = {}
                for s in _subs_list:
                    ow = s.get("owner", "")
                    if ow:
                        _owner_map2.setdefault(ow, []).append(f"{s.get('name', '')} ({s.get('baseline', '')} to {s.get('target', '')} {s.get('unit', '')})")
                df_team = pd.DataFrame([{
                    "Name": m["member_name"],
                    "Role": m.get("role", ""),
                    "Department": m.get("department", ""),
                    "KPI Targets": " | ".join(_owner_map2.get(m["member_name"], ["—"]))
                } for m in team])
                st.dataframe(df_team, use_container_width=True, hide_index=True)
                if can_create:
                    del_name = st.selectbox("Remove member", ["—"] + [m["member_name"] for m in team], key="fi_del_member")
                    if st.button("Remove", key="fi_remove_member") and del_name != "—":
                        mid = next((m["id"] for m in team if m["member_name"] == del_name), None)
                        if mid:
                            supabase.table("fi_project_team").delete().eq("id", mid).execute()
                            st.rerun()
            with st.form("fi_add_member"):
                mc1, mc2, mc3 = st.columns(3)
                m_name = mc1.text_input("Name *")
                m_role = mc2.selectbox("Role", TEAM_ROLES)
                m_dept = mc3.text_input("Department")
                if st.form_submit_button("Add Team Member"):
                    if m_name:
                        supabase.table("fi_project_team").insert({
                            "project_id": pid, "member_name": m_name,
                            "role": m_role, "department": m_dept, "contribution_target": ""
                        }).execute()
                        st.rerun()

        with st.expander("KPI & KAI", expanded=True):
            kpi_vals_form = kpi or {}
            _launch_d = selected_project.get("launch_date", date.today())
            _auto_dt  = date.fromisoformat(str(_launch_d)[:10]) + timedelta(weeks=12)
            _existing_subs = kpi_vals_form.get("sub_components") or []
            if isinstance(_existing_subs, str):
                try: _existing_subs = json.loads(_existing_subs)
                except: _existing_subs = []
            _existing_subs = [s for s in _existing_subs if isinstance(s, dict)]

            _kpi_cats = list(KPI_TREE.keys())
            _cur_cat  = kpi_vals_form.get("kpi_category", "OEE Improvement")
            _cat_idx  = _kpi_cats.index(_cur_cat) if _cur_cat in _kpi_cats else 0
            st.markdown("**Step 1 — Select KPI:**")
            kpi_category = st.selectbox("KPI", _kpi_cats, index=_cat_idx, key="fi_kpi_cat")
            _def_unit  = KPI_TREE[kpi_category]["unit"]
            _kai_opts  = KPI_TREE[kpi_category]["kais"]
            st.markdown("**Step 2 — Select KAIs:**")
            _existing_kai_names = [s.get("name", "") for s in _existing_subs]
            _def_kais = [k for k in _existing_kai_names if k in _kai_opts]
            selected_kais = st.multiselect("KAIs", _kai_opts, default=_def_kais, key="fi_kai_select")

            with st.form("fi_kpi_kai_form"):
                st.markdown(f"##### {kpi_category}")
                _pk1, _pk2, _pk3, _pk4 = st.columns(4)
                _pk1.caption("Unit"); _pk2.caption("Baseline"); _pk3.caption("Target"); _pk4.caption("Target Date")
                _cur_unit = kpi_vals_form.get("unit", _def_unit)
                _u_idx = UNITS.index(_cur_unit) if _cur_unit in UNITS else 0
                kpi_unit     = _pk1.selectbox("Unit", UNITS, index=_u_idx, label_visibility="collapsed")
                kpi_baseline_v = _pk2.number_input("Baseline", label_visibility="collapsed", value=float(kpi_vals_form.get("baseline_value", 0) or 0))
                kpi_target_v   = _pk3.number_input("Target",   label_visibility="collapsed", value=float(kpi_vals_form.get("target_value", 0) or 0))
                kpi_tdate    = _pk4.date_input("Target Date", label_visibility="collapsed",
                    value=date.fromisoformat(str(kpi_vals_form.get("target_date", _auto_dt))[:10]) if kpi_vals_form.get("target_date") else _auto_dt)
                kai_rows = []
                if selected_kais:
                    st.divider()
                    _rh1, _rh2, _rh3, _rh4, _rh5 = st.columns([3, 1.5, 1.2, 1.2, 2])
                    _rh1.caption("KAI"); _rh2.caption("Unit"); _rh3.caption("Baseline"); _rh4.caption("Target"); _rh5.caption("Responsible")
                    _sub_by_name = {s.get("name", ""): s for s in _existing_subs}
                    _team_names  = [""] + [m["member_name"] for m in team]
                    for si, kai in enumerate(selected_kais):
                        ex = _sub_by_name.get(kai, {})
                        k1, k2, k3, k4, k5 = st.columns([3, 1.5, 1.2, 1.2, 2])
                        k1.markdown(f"**{kai}**")
                        _ku  = ex.get("unit", _def_unit)
                        _kui = UNITS.index(_ku) if _ku in UNITS else 0
                        _ko  = ex.get("owner", "")
                        _koi = _team_names.index(_ko) if _ko in _team_names else 0
                        kai_rows.append({
                            "name":     kai,
                            "unit":     k2.selectbox("Unit", UNITS, index=_kui, key=f"kai_u_{si}", label_visibility="collapsed"),
                            "baseline": k3.number_input("Baseline", value=float(ex.get("baseline", 0)), key=f"kai_b_{si}", label_visibility="collapsed"),
                            "target":   k4.number_input("Target",   value=float(ex.get("target", 0)),   key=f"kai_t_{si}", label_visibility="collapsed"),
                            "owner":    k5.selectbox("Responsible", _team_names, index=_koi, key=f"kai_o_{si}", label_visibility="collapsed"),
                        })
                if st.form_submit_button("Save KPI & KAIs", type="primary"):
                    _save = {
                        "project_id": pid, "kpi_name": kpi_category, "unit": kpi_unit,
                        "kpi_category": kpi_category, "sub_kpi_focus": ",".join(selected_kais),
                        "baseline_value": kpi_baseline_v, "target_value": kpi_target_v,
                        "target_date": str(kpi_tdate), "sub_components": json.dumps(kai_rows),
                    }
                    if kpi:
                        supabase.table("fi_project_kpi").update(_save).eq("id", kpi["id"]).execute()
                    else:
                        supabase.table("fi_project_kpi").insert(_save).execute()
                    st.success("Saved")
                    st.rerun()

        if current_week >= 5:
            with st.expander("Cost / Benefit", expanded=False):
                cost_vals = cost or {}
                with st.form("fi_cost_form"):
                    cb1, cb2 = st.columns(2)
                    cost_est  = cb1.number_input("Estimated Project Cost (K EUR)", value=float(cost_vals.get("estimated_cost", 0) or 0))
                    cost_save = cb2.number_input("Estimated Annual Savings (K EUR/yr)", value=float(cost_vals.get("estimated_savings", 0) or 0))
                    payback = round(cost_est / cost_save * 12, 1) if cost_save > 0 else 0
                    st.info(f"Payback Period: **{payback} months** (auto-calculated)")
                    if st.form_submit_button("Save Cost/Benefit"):
                        cb_data = {"project_id": pid, "estimated_cost": cost_est, "estimated_savings": cost_save, "payback_months": payback}
                        if cost:
                            supabase.table("fi_project_cost").update(cb_data).eq("id", cost["id"]).execute()
                        else:
                            supabase.table("fi_project_cost").insert(cb_data).execute()
                        st.success("Saved")
                        st.rerun()

        with st.expander("Master Plan & Gantt Chart", expanded=True):
            with st.form("fi_add_step"):
                ms1, ms2 = st.columns(2)
                step_name = ms1.text_input("Step / Phase Name *")
                step_desc = ms2.text_input("Description")
                mw1, mw2, mw3 = st.columns(3)
                step_start = mw1.number_input("Planned Start Week", min_value=1, max_value=12, value=1)
                step_end   = mw2.number_input("Planned End Week",   min_value=1, max_value=12, value=2)
                step_owner = mw3.selectbox("Owner", [""] + [m["member_name"] for m in team])
                if st.form_submit_button("Add Step"):
                    if step_name:
                        supabase.table("fi_project_steps").insert({
                            "project_id": pid, "step_name": step_name, "description": step_desc,
                            "planned_start_week": int(step_start), "planned_end_week": int(step_end),
                            "owner": step_owner, "sort_order": len(steps)
                        }).execute()
                        st.rerun()
            if steps:
                df_steps = pd.DataFrame(steps)[["step_name", "description", "planned_start_week", "planned_end_week", "owner"]]
                df_steps.columns = ["Step", "Description", "Start Week", "End Week", "Owner"]
                st.dataframe(df_steps, use_container_width=True, hide_index=True)
                gf = _gantt_chart(steps, wu_rows, current_week)
                if gf:
                    st.pyplot(gf)
                    plt.close(gf)
                if can_create:
                    del_step = st.selectbox("Remove step", ["—"] + [s["step_name"] for s in steps], key="fi_del_step")
                    if st.button("Remove Step", key="fi_remove_step") and del_step != "—":
                        sid = next((s["id"] for s in steps if s["step_name"] == del_step), None)
                        if sid:
                            supabase.table("fi_project_steps").delete().eq("id", sid).execute()
                            st.rerun()

    # ════════════════════════════════
    # TAB 2 — WEEKLY UPDATE
    # ════════════════════════════════
    with subtabs[1]:
        st.markdown(f"#### Weekly Update — Week {current_week}")
        sel_week = st.selectbox("View / edit week:", list(range(1, 13)), index=current_week - 1, key="fi_week_sel")
        wu = wu_by_week.get(sel_week, {})
        wu_id = wu.get("id")

        def _save_wu(updates):
            if wu_id:
                supabase.table("fi_weekly_updates").update(updates).eq("id", wu_id).execute()
            else:
                updates["project_id"] = pid
                updates["week_number"] = sel_week
                supabase.table("fi_weekly_updates").insert(updates).execute()
            st.rerun()

        with st.expander("Step Progress", expanded=True):
            if not steps:
                st.info("No steps defined yet.")
            else:
                step_progress = wu.get("step_progress") or []
                if isinstance(step_progress, str):
                    try: step_progress = json.loads(step_progress)
                    except: step_progress = []
                sp_by_id = {sp["step_id"]: sp for sp in step_progress if isinstance(sp, dict)}
                PCT_OPTS = ["0%", "25%", "50%", "75%", "100%"]
                _due_steps = [s for s in steps if s.get("planned_start_week", 1) <= sel_week]
                _future_steps = [s for s in steps if s.get("planned_start_week", 1) > sel_week]
                if _future_steps:
                    st.caption(f"{len(_future_steps)} step(s) not yet due.")
                if _due_steps:
                    with st.form(f"fi_step_progress_{sel_week}"):
                        new_sp = []
                        _chunks = [_due_steps[i:i+4] for i in range(0, len(_due_steps), 4)]
                        for chunk in _chunks:
                            cols = st.columns(len(chunk))
                            for ci, step in enumerate(chunk):
                                sid = str(step["id"])
                                ex  = sp_by_id.get(sid, {})
                                with cols[ci]:
                                    st.markdown(f"**{step['step_name']}**")
                                    st.caption(f"W{step.get('planned_start_week', '')}→W{step.get('planned_end_week', '')} | {step.get('owner', '') or '—'}")
                                    _pct_str = f"{int(ex.get('pct_complete', 0))}%"
                                    _pct_str = _pct_str if _pct_str in PCT_OPTS else "0%"
                                    pct_sel = st.selectbox("Progress", PCT_OPTS, index=PCT_OPTS.index(_pct_str), key=f"sp_p_{sid}", label_visibility="collapsed")
                                    notes = st.text_input("Notes", value=ex.get("notes", ""), key=f"sp_n_{sid}", placeholder="Notes...", label_visibility="collapsed")
                                    _pct_int = int(pct_sel.replace("%", ""))
                                    _auto_status = "Completed" if _pct_int == 100 else "In Progress" if _pct_int > 0 else "Not Started"
                                    new_sp.append({"step_id": sid, "status": _auto_status, "pct_complete": _pct_int, "notes": notes})
                        if st.form_submit_button("Save Step Progress", type="primary"):
                            _save_wu({"step_progress": json.dumps(new_sp), "updated_by": name})

                _wu_fresh = supabase.table("fi_weekly_updates").select("*").eq("project_id", pid).order("week_number").execute().data or []
                gf2 = _gantt_chart(steps, _wu_fresh, current_week)
                if gf2:
                    st.pyplot(gf2)
                    plt.close(gf2)

        with st.expander("KPI & KAI Update", expanded=True):
            if not kpi:
                st.info("No KPI defined yet.")
            else:
                _kai_subs = kpi.get("sub_components") or []
                if isinstance(_kai_subs, str):
                    try: _kai_subs = json.loads(_kai_subs)
                    except: _kai_subs = []
                _kai_subs = [s for s in _kai_subs if isinstance(s, dict)]
                _wu_kai = wu.get("kpi_notes") or ""
                try: _wu_kai_data = json.loads(_wu_kai) if _wu_kai.startswith("{") else {}
                except: _wu_kai_data = {}

                with st.form(f"fi_kpi_entry_{sel_week}"):
                    st.markdown(f"**{kpi.get('kpi_name', '')}** — Baseline: {kpi.get('baseline_value', '')} | Target: {kpi.get('target_value', '')} {kpi.get('unit', '')}")
                    ke1, ke2 = st.columns(2)
                    kpi_val_entry = ke1.number_input(f"This week's value ({kpi.get('unit', '')})", value=float(wu.get("kpi_value", 0) or 0))
                    collected_by  = ke2.multiselect("Collected by", [m["member_name"] for m in team],
                        default=[x for x in (wu.get("kpi_collected_by", "") or "").split(",") if x.strip() in [m["member_name"] for m in team]])
                    kai_weekly = {}
                    if _kai_subs:
                        st.divider()
                        st.caption("KAI updates this week:")
                        _kai_chunks2 = [_kai_subs[i:i+4] for i in range(0, len(_kai_subs), 4)]
                        for chunk2 in _kai_chunks2:
                            _kcols = st.columns(len(chunk2))
                            for ci2, kai2 in enumerate(chunk2):
                                kn = kai2.get("name", "")
                                _ex_val = float(_wu_kai_data.get(kn, {}).get("value", 0))
                                with _kcols[ci2]:
                                    st.markdown(f"**{kn}**")
                                    st.caption(f"Baseline: {kai2.get('baseline', '')} → Target: {kai2.get('target', '')} {kai2.get('unit', '')}")
                                    st.caption(f"Owner: {kai2.get('owner', '—')}")
                                    kai_weekly[kn] = {"value": st.number_input(f"{kn} value", value=_ex_val, key=f"kai_w_{sel_week}_{ci2}", label_visibility="collapsed")}
                    if st.form_submit_button("Save KPI & KAI Update", type="primary"):
                        _save_wu({"kpi_value": kpi_val_entry, "kpi_collected_by": ",".join(collected_by), "kpi_notes": json.dumps(kai_weekly), "updated_by": name})

                trend_wu = supabase.table("fi_weekly_updates").select("*").eq("project_id", pid).order("week_number").execute().data or []
                tf2 = _kpi_trend_chart(kpi, kpi.get("baseline_value"), kpi.get("target_value"), trend_wu)
                st.pyplot(tf2)
                plt.close(tf2)

        if sel_week >= 3:
            with st.expander("Root Cause Analysis", expanded=False):
                with st.form(f"fi_rca_{sel_week}"):
                    rca_done = st.checkbox("Has root cause analysis been performed this week?", value=bool(wu.get("rca_performed")))
                    rca_method = rca_findings = rca_file = None
                    causes_verified = causes_method = None
                    if rca_done:
                        r1, r2 = st.columns(2)
                        rca_method  = r1.selectbox("Method used", RCA_METHODS, index=RCA_METHODS.index(wu.get("rca_method", "5-Why")) if wu.get("rca_method") in RCA_METHODS else 0)
                        rca_file    = r2.file_uploader("Upload findings", type=["pdf", "png", "jpg"], key=f"rca_f_{sel_week}")
                        rca_findings = st.text_area("Describe findings *", value=wu.get("rca_findings", ""), height=100)
                        causes_verified = st.checkbox("Have causes been verified with data?", value=bool(wu.get("causes_verified")))
                        if causes_verified:
                            causes_method = st.text_input("Verification method", value=wu.get("causes_verification_method", ""))
                    reoc_before = reoc_desc = reoc_prev = single_pa = single_notes = None
                    if sel_week >= 7:
                        st.divider()
                        st.markdown("**Reoccurrence Analysis**")
                        reoc_before = st.checkbox("Has a similar problem occurred before?", value=bool(wu.get("reoccurrence_before")))
                        reoc_desc   = st.text_area("Describe previous occurrence", value=wu.get("reoccurrence_description", ""), height=60)
                        reoc_prev   = st.checkbox("Is reoccurrence prevention in place?", value=bool(wu.get("reoccurrence_prevention")))
                        single_pa   = st.checkbox("Single problem analysis applied?", value=bool(wu.get("single_problem_analysis")))
                        single_notes = st.text_input("Follow-up notes", value=wu.get("single_problem_notes", ""))
                    if st.form_submit_button("Save RCA"):
                        upd = {"rca_performed": rca_done, "updated_by": name}
                        if rca_done:
                            upd.update({"rca_method": rca_method, "rca_findings": rca_findings,
                                        "causes_verified": causes_verified or False,
                                        "causes_verification_method": causes_method or ""})
                            if rca_file: upd["rca_file_b64"] = _to_b64(rca_file)
                        if sel_week >= 7 and reoc_desc is not None:
                            upd.update({"reoccurrence_before": reoc_before, "reoccurrence_description": reoc_desc,
                                        "reoccurrence_prevention": reoc_prev, "single_problem_analysis": single_pa,
                                        "single_problem_notes": single_notes or ""})
                        _save_wu(upd)

        if sel_week >= 4:
            with st.expander("Action Plan", expanded=True):
                if actions:
                    df_act = pd.DataFrame(actions)[["description", "root_cause_addressed", "owner", "target_date", "status"]]
                    df_act.columns = ["Action", "Root Cause Addressed", "Owner", "Target Date", "Status"]
                    st.dataframe(df_act, use_container_width=True, hide_index=True)
                    if can_create:
                        act_sel = st.selectbox("Update action status", ["—"] + [a["description"][:50] for a in actions], key="fi_act_update")
                        if act_sel != "—":
                            act_obj = next((a for a in actions if a["description"][:50] == act_sel), None)
                            if act_obj:
                                nc1, nc2, nc3 = st.columns(3)
                                new_status = nc1.selectbox("New Status", ACTION_STATUSES,
                                    index=ACTION_STATUSES.index(act_obj.get("status", "Open")), key="fi_new_status")
                                new_ev = nc2.file_uploader("Upload Evidence", type=["pdf", "png", "jpg", "xlsx"], key="fi_ev_upload")
                                if nc3.button("Update Action", key="fi_update_act"):
                                    upd = {"status": new_status}
                                    if new_status == "Completed": upd["completed_date"] = str(date.today())
                                    if new_ev: upd["evidence_b64"] = _to_b64(new_ev); upd["evidence_filename"] = new_ev.name
                                    supabase.table("fi_actions").update(upd).eq("id", act_obj["id"]).execute()
                                    st.rerun()
                with st.form(f"fi_add_action_{sel_week}"):
                    ac1, ac2 = st.columns(2)
                    act_desc  = ac1.text_area("Action Description *", height=70)
                    act_rc    = ac2.text_area("Root Cause Addressed", height=70)
                    ac3, ac4, ac5 = st.columns(3)
                    act_owner = ac3.selectbox("Owner", [""] + [m["member_name"] for m in team], key=f"act_owner_{sel_week}")
                    act_date  = ac4.date_input("Target Date", value=date.today() + timedelta(weeks=2))
                    act_ev    = ac5.file_uploader("Evidence (optional)", type=["pdf", "png", "jpg"], key=f"act_ev_{sel_week}")
                    if st.form_submit_button("Add Action"):
                        if act_desc:
                            supabase.table("fi_actions").insert({
                                "project_id": pid, "description": act_desc, "root_cause_addressed": act_rc,
                                "owner": act_owner, "target_date": str(act_date), "status": "Open",
                                "created_week": sel_week,
                                "evidence_b64": _to_b64(act_ev) if act_ev else None,
                                "evidence_filename": act_ev.name if act_ev else None,
                            }).execute()
                            st.rerun()

        if sel_week >= 4:
            with st.expander("Basic Conditions", expanded=False):
                with st.form(f"fi_basic_{sel_week}"):
                    bc_done  = st.checkbox("Have critical areas been restored to basic conditions?", value=bool(wu.get("basic_conditions_restored")))
                    bc_desc  = bc_before = bc_after = bc_date = None
                    if bc_done:
                        bc_desc  = st.text_area("Describe what was restored", value=wu.get("basic_conditions_description", ""), height=60)
                        bc_b1, bc_b2 = st.columns(2)
                        bc_before = bc_b1.file_uploader("Before photo", type=["jpg", "jpeg", "png"], key=f"bc_before_{sel_week}")
                        bc_after  = bc_b2.file_uploader("After photo",  type=["jpg", "jpeg", "png"], key=f"bc_after_{sel_week}")
                        bc_date   = st.date_input("Date completed", value=date.today())
                    if st.form_submit_button("Save Basic Conditions"):
                        upd = {"basic_conditions_restored": bc_done, "updated_by": name}
                        if bc_done:
                            upd["basic_conditions_description"] = bc_desc or ""
                            upd["basic_conditions_date"] = str(bc_date) if bc_date else None
                            if bc_before: upd["basic_conditions_before_b64"] = _to_b64(bc_before)
                            if bc_after:  upd["basic_conditions_after_b64"]  = _to_b64(bc_after)
                        _save_wu(upd)

        with st.expander("Team Meeting Log", expanded=False):
            with st.form(f"fi_meeting_{sel_week}"):
                mt1, mt2 = st.columns(2)
                mtg_held  = mt1.checkbox("Meeting held this week?", value=bool(wu.get("meeting_held")))
                attendees = mt2.multiselect("Attendees", [m["member_name"] for m in team],
                    default=[x.strip() for x in (wu.get("meeting_attendees", "") or "").split(",")
                             if x.strip() in [m["member_name"] for m in team]])
                mtg_notes = st.text_area("Meeting notes", value=wu.get("meeting_notes", ""), height=60)
                if team:
                    att_rate = len(attendees) / len(team) * 100
                    st.info(f"Attendance: {att_rate:.0f}% ({len(attendees)}/{len(team)} members)")
                if st.form_submit_button("Save Meeting Log"):
                    _save_wu({"meeting_held": mtg_held, "meeting_attendees": ",".join(attendees), "meeting_notes": mtg_notes, "updated_by": name})

    # ════════════════════════════════
    # TAB 3 — RESULTS
    # ════════════════════════════════
    with subtabs[2]:
        st.markdown("#### Results Tracker")
        if not kpi:
            st.info("No KPI defined yet.")
        else:
            kpi_vals_list = [wu["kpi_value"] for wu in sorted(wu_rows, key=lambda x: x["week_number"]) if wu.get("kpi_value") is not None]
            baseline_v = kpi.get("baseline_value") or 0
            target_v   = kpi.get("target_value")   or 0
            if kpi_vals_list:
                current_val = kpi_vals_list[-1]
                total_gap2  = abs(target_v - baseline_v)
                progress2   = abs(current_val - baseline_v) / total_gap2 * 100 if total_gap2 > 0 else 0
                progress2   = min(100, progress2)
                rc1, rc2, rc3, rc4 = st.columns(4)
                rc1.metric("Baseline", f"{baseline_v} {kpi.get('unit', '')}")
                rc2.metric("Current",  f"{current_val} {kpi.get('unit', '')}")
                rc3.metric("Target",   f"{target_v} {kpi.get('unit', '')}")
                rc4.metric("Progress", f"{progress2:.0f}%")
                if len(kpi_vals_list) >= 2:
                    rate = (kpi_vals_list[-1] - kpi_vals_list[0]) / len(kpi_vals_list)
                    if rate != 0:
                        weeks_left = (target_v - kpi_vals_list[-1]) / rate
                        proj_week  = current_week + weeks_left
                        st.info(f"Projected completion: **Week {proj_week:.0f}** (linear extrapolation)")
            tf3 = _kpi_trend_chart(kpi, baseline_v, target_v, wu_rows)
            st.pyplot(tf3)
            plt.close(tf3)
            if cost:
                st.divider()
                cb_c1, cb_c2, cb_c3 = st.columns(3)
                cb_c1.metric("Estimated Cost",    f"{cost.get('estimated_cost', 0):.1f} K EUR")
                cb_c2.metric("Annual Savings",    f"{cost.get('estimated_savings', 0):.1f} K EUR/yr")
                cb_c3.metric("Payback Period",    f"{cost.get('payback_months', 0):.1f} months")

    # ════════════════════════════════
    # TAB 4 — STABILISATION
    # ════════════════════════════════
    with subtabs[3]:
        st.markdown("#### Stabilisation")
        stab_vals = stab or {}

        with st.expander(f"Standards & Procedures {'(Available from Week 10)' if current_week < 10 else ''}", expanded=current_week >= 10):
            with st.form("fi_stab_procedures"):
                proc_done = st.checkbox("Have procedures been created to hold the gains?",
                                        value=bool(stab_vals.get("procedures_created")), disabled=current_week < 10)
                procedures = []
                if proc_done:
                    n_procs = st.number_input("Number of procedures", min_value=1, max_value=10,
                                              value=max(1, len(stab_vals.get("procedures") or [])))
                    existing_procs = stab_vals.get("procedures") or []
                    if isinstance(existing_procs, str):
                        try: existing_procs = json.loads(existing_procs)
                        except: existing_procs = []
                    for pi in range(int(n_procs)):
                        ep = existing_procs[pi] if pi < len(existing_procs) else {}
                        pc1, pc2 = st.columns(2)
                        procedures.append({
                            "name": pc1.text_input(f"Procedure {pi+1} Name", value=ep.get("name", ""), key=f"proc_n_{pi}"),
                            "description": pc2.text_input("Description", value=ep.get("description", ""), key=f"proc_d_{pi}"),
                        })
                if st.form_submit_button("Save Procedures"):
                    sd = {"procedures_created": proc_done, "procedures": json.dumps(procedures)}
                    if stab: supabase.table("fi_stabilisation").update(sd).eq("id", stab["id"]).execute()
                    else: supabase.table("fi_stabilisation").insert({"project_id": pid, **sd}).execute()
                    st.rerun()

        with st.expander(f"CIL Standards {'(Available from Week 5)' if current_week < 5 else ''}", expanded=current_week >= 5):
            with st.form("fi_stab_cil"):
                cil_def   = st.checkbox("Are CIL standards defined?", value=bool(stab_vals.get("cil_standards_defined")), disabled=current_week < 5)
                cil_score = st.number_input("Latest CIL Audit Score (%)", min_value=0.0, max_value=100.0,
                                            value=float(stab_vals.get("cil_audit_score", 0) or 0))
                cil_file  = st.file_uploader("Upload CIL audit record", type=["pdf", "xlsx", "png"], key="fi_cil_file")
                if cil_score >= 90:
                    st.success("CIL score meets >= 90% target")
                elif cil_score > 0:
                    st.warning(f"CIL score {cil_score:.0f}% — target is >= 90%")
                if st.form_submit_button("Save CIL"):
                    cd = {"cil_standards_defined": cil_def, "cil_audit_score": cil_score}
                    if cil_file: cd["cil_file_b64"] = _to_b64(cil_file)
                    if stab: supabase.table("fi_stabilisation").update(cd).eq("id", stab["id"]).execute()
                    else: supabase.table("fi_stabilisation").insert({"project_id": pid, **cd}).execute()
                    st.rerun()

        with st.expander(f"Monitoring Systems {'(Available from Week 7)' if current_week < 7 else ''}", expanded=current_week >= 7):
            with st.form("fi_stab_monitoring"):
                mon_in    = st.checkbox("Are monitoring systems in place?", value=bool(stab_vals.get("monitoring_in_place")), disabled=current_week < 7)
                mon_types = st.multiselect("System types", MONITORING_TYPES,
                    default=[x.strip() for x in (stab_vals.get("monitoring_types", "") or "").split(",") if x.strip() in MONITORING_TYPES])
                mon_active = st.checkbox("Are they being actively used?", value=bool(stab_vals.get("monitoring_active")))
                mon_date   = st.date_input("Last update date",
                    value=date.fromisoformat(str(stab_vals["monitoring_last_update"])) if stab_vals.get("monitoring_last_update") else date.today())
                mon_ev = st.file_uploader("Upload evidence", type=["pdf", "png", "jpg"], key="fi_mon_ev")
                if st.form_submit_button("Save Monitoring"):
                    md = {"monitoring_in_place": mon_in, "monitoring_types": ",".join(mon_types),
                          "monitoring_active": mon_active, "monitoring_last_update": str(mon_date)}
                    if mon_ev: md["monitoring_evidence_b64"] = _to_b64(mon_ev)
                    if stab: supabase.table("fi_stabilisation").update(md).eq("id", stab["id"]).execute()
                    else: supabase.table("fi_stabilisation").insert({"project_id": pid, **md}).execute()
                    st.rerun()

        with st.expander(f"OPLs & Training {'(Available from Week 10)' if current_week < 10 else ''}", expanded=current_week >= 10):
            with st.form("fi_stab_opls"):
                existing_opls = stab_vals.get("opls") or []
                if isinstance(existing_opls, str):
                    try: existing_opls = json.loads(existing_opls)
                    except: existing_opls = []
                n_opls = st.number_input("Number of OPLs/SOPs created", min_value=0, max_value=20, value=max(0, len(existing_opls)))
                new_opls = []
                for oi in range(int(n_opls)):
                    eo = existing_opls[oi] if oi < len(existing_opls) else {}
                    oc1, oc2, oc3 = st.columns(3)
                    new_opls.append({
                        "title":    oc1.text_input(f"OPL/SOP {oi+1} Title",   value=eo.get("title", ""),    key=f"opl_t_{oi}"),
                        "covers":   oc2.text_input("Improvement it covers",    value=eo.get("covers", ""),   key=f"opl_c_{oi}"),
                        "created_by": oc3.selectbox("Created by", [""] + [m["member_name"] for m in team], key=f"opl_cb_{oi}"),
                    })
                st.divider()
                st.markdown("**Training Matrix**")
                existing_tm = stab_vals.get("training_matrix") or []
                if isinstance(existing_tm, str):
                    try: existing_tm = json.loads(existing_tm)
                    except: existing_tm = []
                new_tm = []
                for mi2, member in enumerate(team):
                    mname2 = member["member_name"]
                    em = next((x for x in existing_tm if x.get("member") == mname2), {})
                    tc1, tc2, tc3 = st.columns(3)
                    new_tm.append({
                        "member": mname2,
                        "opls_trained": tc1.multiselect(f"{mname2} — OPLs trained",
                            [o["title"] for o in new_opls if o["title"]],
                            default=[x for x in (em.get("opls_trained") or []) if x in [o["title"] for o in new_opls]],
                            key=f"tm_opl_{mi2}"),
                        "training_date":    str(tc2.date_input("Training date", value=date.today(), key=f"tm_date_{mi2}")),
                        "plan_confirmed":   tc3.checkbox("Training planned?", value=bool(em.get("plan_confirmed")), key=f"tm_plan_{mi2}"),
                    })
                if st.form_submit_button("Save OPLs & Training"):
                    od = {"opls": json.dumps(new_opls), "training_matrix": json.dumps(new_tm)}
                    if stab: supabase.table("fi_stabilisation").update(od).eq("id", stab["id"]).execute()
                    else: supabase.table("fi_stabilisation").insert({"project_id": pid, **od}).execute()
                    st.rerun()

        with st.expander(f"Workplace & 5S {'(Available from Week 5)' if current_week < 5 else ''}", expanded=current_week >= 5):
            with st.form("fi_stab_5s"):
                five_s       = st.slider("5S Rating", 1, 5, int(stab_vals.get("five_s_rating", 1) or 1))
                five_s_ph    = st.file_uploader("Upload 5S photos", type=["jpg", "jpeg", "png"], key="fi_5s_photo")
                five_s_notes = st.text_area("5S notes", value=stab_vals.get("five_s_notes", ""), height=60)
                imp_vis = st.checkbox("Are improvements evident?", value=bool(stab_vals.get("improvements_visible")))
                imp_ph  = st.file_uploader("Upload improvement photos", type=["jpg", "jpeg", "png"], key="fi_imp_photo")
                if st.form_submit_button("Save Workplace Evidence"):
                    wd = {"five_s_rating": five_s, "five_s_notes": five_s_notes, "improvements_visible": imp_vis}
                    if five_s_ph: wd["five_s_photos_b64"] = _to_b64(five_s_ph)
                    if imp_ph:    wd["improvements_photos_b64"] = _to_b64(imp_ph)
                    if stab: supabase.table("fi_stabilisation").update(wd).eq("id", stab["id"]).execute()
                    else: supabase.table("fi_stabilisation").insert({"project_id": pid, **wd}).execute()
                    st.rerun()

    # ════════════════════════════════
    # TAB 5 — AUDIT VIEW
    # ════════════════════════════════
    if is_auditor and len(subtabs) > 4:
        with subtabs[4]:
            st.markdown("#### Audit View")
            _stab_fresh = supabase.table("fi_stabilisation").select("*").eq("project_id", pid).execute().data or []
            stab_fresh  = _stab_fresh[0] if _stab_fresh else {}
            try:
                scores_a, total_a = _score_project(selected_project, team, kpi, steps, wu_rows, actions, stab_fresh, audit_records)
            except Exception as _e:
                st.error(f"Scoring error: {_e}")
                scores_a = {q: 0 for q in QUESTIONS}
                total_a  = 0

            sc_col = "green" if total_a >= 70 else "orange" if total_a >= 45 else "red"
            st.markdown(f"## Total Score: :{sc_col}[{int(total_a)} / 100]")
            tgt_wk = TARGET_RAMP.get(current_week, 100)
            gap_a  = total_a - tgt_wk
            if gap_a >= 0:
                st.success(f"On track — {int(total_a)}/100 vs target {tgt_wk}/100 (+{gap_a:.0f} pts ahead)")
            else:
                st.warning(f"Behind target — {int(total_a)}/100 vs target {tgt_wk}/100 ({gap_a:.0f} pts behind)")

            # Score trajectory
            fig_tr, ax_tr = plt.subplots(figsize=(10, 4), dpi=100)
            tw2 = sorted(TARGET_RAMP.keys())
            tt2 = [TARGET_RAMP[w] for w in tw2]
            ax_tr.plot(tw2, tt2, "o--", color="#888888", label="Target ramp", linewidth=1.5)
            sbw = {ar["week_number"]: ar.get("total_score", 0) for ar in audit_records}
            aw2 = sorted(sbw.keys())
            av2 = [sbw[w] for w in aw2]
            if aw2:
                ax_tr.plot(aw2, av2, "o-", color="#006394", label="Actual score", linewidth=2)
            ax_tr.axvline(current_week, color="#DE201B", linewidth=1.5, linestyle=":", label=f"Week {current_week}")
            ax_tr.set_xticks(range(1, 13))
            ax_tr.set_xticklabels([f"W{i}" for i in range(1, 13)])
            ax_tr.set_ylabel("Score / 100")
            ax_tr.set_ylim(0, 105)
            ax_tr.spines["top"].set_visible(False)
            ax_tr.spines["right"].set_visible(False)
            ax_tr.legend(fontsize=9, frameon=False)
            plt.tight_layout()
            st.pyplot(fig_tr)
            plt.close(fig_tr)

            # Dimension traffic lights
            st.markdown("#### Dimension Traffic Lights")
            dimensions = {
                "Involvement":   [1, 2, 34, 35, 36],
                "Method":        [8, 9, 10, 11, 12, 13, 14, 15, 16],
                "Action Plan":   [17, 18, 19, 20, 21, 22, 23],
                "Results":       [3, 4, 5, 6, 7, 24, 25],
                "Stabilisation": [26, 27, 28, 29, 30, 31, 32, 33],
            }
            dim_cols = st.columns(len(dimensions))
            for ci3, (dim, qs) in enumerate(dimensions.items()):
                achieved2 = sum(scores_a.get(q, 0) for q in qs)
                possible2 = sum(QUESTIONS[q]["score"] for q in qs)
                pct2 = achieved2 / possible2 * 100 if possible2 > 0 else 0
                color2 = "green" if pct2 >= 90 else "orange" if pct2 >= 70 else "red"
                icon = "🟢" if pct2 >= 90 else "🟡" if pct2 >= 70 else "🔴"
                dim_cols[ci3].metric(f"{icon} {dim}", f"{achieved2}/{possible2}", f"{pct2:.0f}%")

            # Per-question table
            st.divider()
            st.markdown("#### Per-Question Status")
            q_data = []
            for qn, qdata in QUESTIONS.items():
                achieved3 = scores_a.get(qn, 0)
                due = current_week >= qdata["week"]
                status3 = ("Met" if achieved3 >= qdata["score"] else "Not Met" if due else "Not Due Yet")
                q_data.append({
                    "Q#": qn, "Question": qdata["text"][:60],
                    "Weight": qdata["score"], "Due Week": qdata["week"],
                    "Status": status3, "Score": f"{achieved3}/{qdata['score']}"
                })
            st.dataframe(pd.DataFrame(q_data), use_container_width=True, hide_index=True)

            # Gap register
            gaps2 = [(qn, QUESTIONS[qn]) for qn in QUESTIONS
                     if scores_a.get(qn, 0) < QUESTIONS[qn]["score"] and current_week >= QUESTIONS[qn]["week"]]
            if gaps2:
                st.divider()
                st.markdown("#### Gap Register")
                for qn, qdata in gaps2:
                    st.markdown(f"**Q{qn}** ({qdata['score']} pts) — {qdata['text']}")

            # Audit form
            st.divider()
            st.markdown("#### Team Understanding Check")
            with st.form("fi_audit_understanding"):
                au1, au2 = st.columns(2)
                tested_member     = au1.text_input("Team member tested")
                understanding_pass = au2.checkbox("Did they explain correctly?")
                audit_notes2 = st.text_area("Auditor notes", height=80)
                if st.form_submit_button("Save Audit Record"):
                    q_scores2 = {str(qn): scores_a.get(qn, 0) for qn in QUESTIONS}
                    if understanding_pass:
                        q_scores2["34"] = True
                        q_scores2["35"] = True
                    ar_data = {
                        "project_id": pid, "week_number": current_week,
                        "question_scores": json.dumps(q_scores2),
                        "total_score": total_a,
                        "team_understanding_tested": bool(tested_member),
                        "member_tested": tested_member,
                        "understanding_pass": understanding_pass,
                        "auditor_notes": audit_notes2,
                        "audited_by": name,
                    }
                    existing_ar = next((ar for ar in audit_records if ar["week_number"] == current_week), None)
                    if existing_ar:
                        supabase.table("fi_audit_records").update(ar_data).eq("id", existing_ar["id"]).execute()
                    else:
                        supabase.table("fi_audit_records").insert(ar_data).execute()
                    st.success("Audit record saved")
                    st.rerun()

            # PDF Export
            st.divider()
            st.markdown("#### Generate PDF Report")
            if st.button("Generate Full Project Report PDF", type="primary", key="fi_pdf_gen"):
                with st.spinner("Generating PDF..."):
                    _stab_pdf = supabase.table("fi_stabilisation").select("*").eq("project_id", pid).execute().data
                    _stab_pdf = _stab_pdf[0] if _stab_pdf else {}
                    pdf_buf = _generate_project_pdf(
                        selected_project, team, kpi, steps,
                        wu_rows, actions, _stab_pdf,
                        audit_records, scores_a, total_a
                    )
                st.download_button(
                    "Download PDF", data=pdf_buf,
                    file_name=f"FI_{selected_project['project_name'].replace(' ', '_')}.pdf",
                    mime="application/pdf", key="fi_pdf_download"
                )
