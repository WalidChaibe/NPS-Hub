# fi_projects_tab.py  — Hub + Forms architecture
# Each checklist requirement links to a real form. Completing the form = requirement met.
import io, json, base64, math
from datetime import date, timedelta
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import streamlit as st
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib.colors import HexColor

# ─────────────────────────────────────────────────────────────────────────────
# REQUIREMENTS  — what each req needs + which form fulfils it
# ─────────────────────────────────────────────────────────────────────────────
# form keys: "team" | "kpi_setup" | "master_plan" | "meeting" | "rca" |
#            "actions" | "cil" | "five_s" | "monitoring" | "opls" | "training"
REQUIREMENTS = [
    {"id":1,  "week":1,  "pts":1,  "form":"team",
     "text":"Are team members listed?",
     "where":"Launch Report 2.1 · Team Members Sheet",
     "links":["Launch Report 2.1","Team Members"]},
    {"id":2,  "week":1,  "pts":1,  "form":"team",
     "text":"Have all team members been assigned clear roles?",
     "where":"Launch Report 2.1 · Team Members Sheet · Who Does What",
     "links":["Launch Report 2.1","Team Members","Who Does What"]},
    {"id":3,  "week":1,  "pts":4,  "form":"kpi_setup",
     "text":"Is there a clear link to a company/area KPI?",
     "where":"Launch Report 5",
     "links":["Launch Report 5"]},
    {"id":4,  "week":1,  "pts":1,  "form":"kpi_setup",
     "text":"Is the historical data (timeframe and value) clearly shown?",
     "where":"Launch Report 1.1–1.3",
     "links":["Launch Report 1.1-1.3"]},
    {"id":5,  "week":1,  "pts":1,  "form":"kpi_setup",
     "text":"Is the KPI subdivided into definable components?",
     "where":"KPI Breakdown Sheet",
     "links":["KPI Breakdown"]},
    {"id":6,  "week":1,  "pts":2,  "form":"master_plan",
     "text":"Are Route and Master Plan clearly visible and up-to-date?",
     "where":"Team Master Plan Sheet",
     "links":["Master Plan"]},
    {"id":7,  "week":1,  "pts":2,  "form":"meeting",
     "text":"Are meetings organized and attendance at expected levels?",
     "where":"Launch Report 1.4 · Meeting Attendance Sheet",
     "links":["Launch Report 1.4","Meeting Attendance"]},
    {"id":8,  "week":2,  "pts":1,  "form":"kpi_setup",
     "text":"Is the KPI target (timeframe and value) clearly shown?",
     "where":"Launch Report 4 & 5",
     "links":["Launch Report 4 & 5"]},
    {"id":9,  "week":2,  "pts":1,  "form":"master_plan",
     "text":"Is the target of each step clear?",
     "where":"KPI Breakdown · Team Master Plan",
     "links":["KPI Breakdown","Master Plan"]},
    {"id":10, "week":2,  "pts":1,  "form":"meeting",
     "text":"Is data collection consistent across all shifts?",
     "where":"Data Collection Plan Sheet",
     "links":["Data Collection Plan"]},
    {"id":11, "week":3,  "pts":1,  "form":"master_plan",
     "text":"Have step targets been subdivided into specific activities?",
     "where":"Who Does What · Team Master Plan",
     "links":["Who Does What","Master Plan"]},
    {"id":12, "week":3,  "pts":5,  "form":"team",
     "text":"Is the route/methodology well understood by all team members?",
     "where":"Launch Report 3",
     "links":["Launch Report 3"]},
    {"id":13, "week":3,  "pts":3,  "form":"team",
     "text":"Can a randomly picked team member explain the activity board?",
     "where":"Team Master Plan · Who Does What",
     "links":["Master Plan","Who Does What"]},
    {"id":14, "week":4,  "pts":4,  "form":"kpi_setup",
     "text":"Has a Cost/Benefit chart been introduced and kept up-to-date?",
     "where":"Financial KPI Tracking Sheet",
     "links":["Financial KPI"]},
    {"id":15, "week":4,  "pts":1,  "form":"rca",
     "text":"Have root-cause analyses been used and well documented?",
     "where":"Root Cause Analysis Sheet",
     "links":["Root Cause Analysis"]},
    {"id":16, "week":4,  "pts":3,  "form":"cil",
     "text":"Have critical areas been restored to basic conditions?",
     "where":"Restoration & Basic Conditions Sheet",
     "links":["Restoration"]},
    {"id":17, "week":4,  "pts":5,  "form":"actions",
     "text":"Are planned actions visible with target completion dates?",
     "where":"Who Does What · Team Master Plan",
     "links":["Who Does What","Master Plan"]},
    {"id":18, "week":5,  "pts":1,  "form":"rca",
     "text":"Have problem causes been verified and quantified with data?",
     "where":"Root Cause Analysis Sheet",
     "links":["Root Cause Analysis"]},
    {"id":19, "week":5,  "pts":2,  "form":"rca",
     "text":"Has the team used the route methods/tools to attack problems?",
     "where":"Launch Report 3 · Root Cause Analysis",
     "links":["Launch Report 3","Root Cause Analysis"]},
    {"id":20, "week":5,  "pts":2,  "form":"actions",
     "text":"Is there an owner assigned to each action?",
     "where":"Who Does What Sheet",
     "links":["Who Does What"]},
    {"id":21, "week":5,  "pts":2,  "form":"actions",
     "text":"Is the action plan up-to-date?",
     "where":"Who Does What Sheet",
     "links":["Who Does What"]},
    {"id":22, "week":5,  "pts":3,  "form":"actions",
     "text":"Is the majority of actions completed on time?",
     "where":"Who Does What Sheet",
     "links":["Who Does What"]},
    {"id":23, "week":5,  "pts":4,  "form":"opls",
     "text":"Is there evidence of implemented actions (OPLs, pictures, standards)?",
     "where":"OPL & SOP Register Sheet",
     "links":["OPL & SOP"]},
    {"id":24, "week":5,  "pts":3,  "form":"cil",
     "text":"Are CIL standards clear? Do CIL audits achieve at least 90%?",
     "where":"CIL Standards & Audits Sheet",
     "links":["CIL Standards"]},
    {"id":25, "week":5,  "pts":1,  "form":"five_s",
     "text":"Is the workplace well organized (5S)?",
     "where":"5S Audit Sheet",
     "links":["5S Audit"]},
    {"id":26, "week":5,  "pts":1,  "form":"opls",
     "text":"Are improvements on the targeted areas evident?",
     "where":"Restoration · OPL & SOP Register",
     "links":["Restoration","OPL & SOP"]},
    {"id":27, "week":6,  "pts":2,  "form":"rca",
     "text":"Has the group found logical countermeasures with sound logic?",
     "where":"Root Cause Analysis · Who Does What",
     "links":["Root Cause Analysis","Who Does What"]},
    {"id":28, "week":7,  "pts":1,  "form":"rca",
     "text":"Is the reoccurrence analysis present and updated?",
     "where":"Root Cause Analysis Sheet",
     "links":["Root Cause Analysis"]},
    {"id":29, "week":7,  "pts":1,  "form":"rca",
     "text":"Is the single problem analysis applied and followed up?",
     "where":"Root Cause Analysis Sheet",
     "links":["Root Cause Analysis"]},
    {"id":30, "week":7,  "pts":15, "form":"kpi_setup",
     "text":"Is the trend of the performance indicator positive?",
     "where":"KPI Count · Financial KPI Tracking",
     "links":["KPI Count","Financial KPI"]},
    {"id":31, "week":7,  "pts":2,  "form":"monitoring",
     "text":"Are monitoring systems for key actions in place and visible?",
     "where":"Monitoring & Controls Sheet",
     "links":["Monitoring"]},
    {"id":32, "week":8,  "pts":2,  "form":"monitoring",
     "text":"Are monitoring system devices used and up-to-date?",
     "where":"Monitoring & Controls Sheet",
     "links":["Monitoring"]},
    {"id":33, "week":10, "pts":2,  "form":"opls",
     "text":"Are there procedures in place to hold the gains achieved?",
     "where":"OPL & SOP · Monitoring & Controls",
     "links":["OPL & SOP","Monitoring"]},
    {"id":34, "week":10, "pts":2,  "form":"opls",
     "text":"Have OPLs/SOPs been created for every significant improvement?",
     "where":"OPL & SOP Register Sheet",
     "links":["OPL & SOP"]},
    {"id":35, "week":10, "pts":2,  "form":"training",
     "text":"Is there a training matrix for OPLs/SOPs with a training plan?",
     "where":"Training Matrix Sheet",
     "links":["Training Matrix"]},
    {"id":36, "week":11, "pts":15, "form":"kpi_setup",
     "text":"Has the team achieved its goal or made substantial progress?",
     "where":"KPI Count · Financial KPI Tracking",
     "links":["KPI Count","Financial KPI"]},
]

from collections import defaultdict
REQ_BY_WEEK   = defaultdict(list)
REQ_BY_FORM   = defaultdict(list)
REQ_BY_ID     = {}
for r in REQUIREMENTS:
    REQ_BY_WEEK[r["week"]].append(r)
    REQ_BY_FORM[r["form"]].append(r)
    REQ_BY_ID[r["id"]] = r

ACTIVE_WEEKS  = sorted(REQ_BY_WEEK.keys())
TOTAL_POINTS  = sum(r["pts"] for r in REQUIREMENTS)   # 100

TARGET_RAMP   = {1:12,2:15,3:24,4:33,5:56,6:58,
                 7:77,8:79,9:79,10:85,11:100,12:100}

FORM_META = {
    "team":        {"icon":"👥","label":"Team Setup",         "color":"#0C5595"},
    "kpi_setup":   {"icon":"📈","label":"KPI & Results",      "color":"#1E8449"},
    "master_plan": {"icon":"🗓","label":"Master Plan",         "color":"#0E5E86"},
    "meeting":     {"icon":"📋","label":"Meeting Minutes",     "color":"#8E44AD"},
    "rca":         {"icon":"🔍","label":"Root Cause Analysis", "color":"#D68910"},
    "actions":     {"icon":"✅","label":"Action Plan",         "color":"#006394"},
    "cil":         {"icon":"🔧","label":"CIL & Restoration",   "color":"#DE201B"},
    "five_s":      {"icon":"⭐","label":"5S Audit",            "color":"#D68910"},
    "monitoring":  {"icon":"📊","label":"Monitoring",          "color":"#566573"},
    "opls":        {"icon":"📝","label":"OPLs & SOPs",         "color":"#17375E"},
    "training":    {"icon":"🎓","label":"Training Matrix",     "color":"#8E44AD"},
}

PLANT_SECTIONS = {
    "Corrugator":    ["BHS","Fosber"],
    "Die-Cut":       ["BOBST 160-II","BOBST 203","BOBST MASTERCUT 1","BOBST MASTERCUT 2"],
    "FFG":           ["LMC FFG","MARTIN 616","924","SATURN"],
    "Folder Gluers": ["Bahmüller TURBOX","VEGA 2"],
    "Stitcher":      ["BAHMULLER STITCHER"],
    "Printer":       ["IPACK"],
    "Pre-Print":     ["CI4","CI6"],
    "QuickSet":      ["QuickSet"],
    "Jumbo":         ["JUMBO"],
    "RM Warehouse":  [],
    "FG Warehouse":  [],
    "Maintenance":   [],
}

TEAM_ROLES    = ["Team Leader","Analyst","Operator","Maintenance","Quality","Other"]
ACTION_STATUS = ["Open","In Progress","Completed","Overdue"]

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def _cw(launch_date):
    if not launch_date: return 1
    if isinstance(launch_date, str):
        try: launch_date = date.fromisoformat(launch_date)
        except: return 1
    delta = (date.today() - launch_date).days
    return max(1, min(12, math.ceil(delta/7) if delta>0 else 1))

def _b64(f):
    if f is None: return None
    return base64.b64encode(f.getvalue()).decode()

def _safe_date(d):
    try: return date.fromisoformat(str(d)[:10])
    except: return None

def _pj(v, fb):
    if isinstance(v, list): return v
    if isinstance(v, str) and v.strip():
        try: return json.loads(v)
        except: pass
    return fb

def _load_checklist(supabase, pid):
    try:
        rows = supabase.table("fi_project_checklist").select("*")\
            .eq("project_id",pid).execute().data or []
        return {r["req_id"]: r for r in rows}
    except: return {}

def _sync_checklist_from_data(supabase, pid, name):
    """
    Re-evaluate every requirement from live Supabase data and write
    the result to fi_project_checklist. Called on every page load.
    This ensures requirements show as done as soon as the data exists,
    regardless of whether the user explicitly submitted a form.
    """
    try:
        team    = supabase.table("fi_project_team").select("*").eq("project_id",pid).execute().data or []
        kpi_r   = supabase.table("fi_project_kpi").select("*").eq("project_id",pid).execute().data or []
        kpi     = kpi_r[0] if kpi_r else {}
        steps   = supabase.table("fi_project_steps").select("*").eq("project_id",pid).execute().data or []
        wu_rows = supabase.table("fi_weekly_updates").select("*").eq("project_id",pid).execute().data or []
        actions = supabase.table("fi_actions").select("*").eq("project_id",pid).execute().data or []
        stab_r  = supabase.table("fi_stabilisation").select("*").eq("project_id",pid).execute().data or []
        stab    = stab_r[0] if stab_r else {}
        proj_r  = supabase.table("fi_projects").select("*").eq("id",pid).execute().data or []
        proj    = proj_r[0] if proj_r else {}
        meets   = supabase.table("fi_meetings").select("*").eq("project_id",pid).execute().data or []
        analyses= supabase.table("fi_project_analysis").select("*").eq("project_id",pid).execute().data or []
    except:
        return  # silently skip if tables missing

    subs      = _pj(kpi.get("sub_components"),[])
    readings  = sorted([w for w in wu_rows if w.get("kpi_value") is not None], key=lambda x:x["week_number"])
    base      = float(kpi.get("baseline_value") or 0)
    tgt       = float(kpi.get("target_value")   or 0)
    has_rca   = any(a.get("analysis_type") in ("5why","fishbone") for a in analyses)
    has_root  = any((_pj(a.get("data"),{}) if isinstance(a.get("data"),str) else a.get("data",{})).get("root_cause")
                    for a in analyses if a.get("analysis_type")=="5why")
    past_due  = [a for a in actions if _safe_date(a.get("target_date")) and _safe_date(a["target_date"])<date.today()]
    on_time   = [a for a in past_due if a.get("status")=="Completed"]
    opls      = _pj(stab.get("opls"),[])

    # trend positive
    trend_ok = False
    if len(readings)>=3 and base!=tgt:
        trend_ok = (tgt>base and readings[-1]["kpi_value"]>readings[-3]["kpi_value"]) or                    (tgt<base and readings[-1]["kpi_value"]<readings[-3]["kpi_value"])

    # goal achieved
    goal_ok = False
    if readings and base!=0 and tgt!=0:
        gap  = abs(tgt-base)
        prog = abs(readings[-1]["kpi_value"]-base)/gap if gap else 0
        goal_ok = prog>=0.80

    # Helper to safely parse analysis data
    def _adata(a):
        d = a.get("data", {})
        return _pj(d, {}) if isinstance(d, str) else (d if isinstance(d, dict) else {})

    has_subs = len([s for s in subs if isinstance(s, dict) and s.get("name")]) > 0

    results = {
        # ── WEEK 1 ──────────────────────────────────────────────────────────
        # Req 1: team members listed
        1:  len(team) >= 1,
        # Req 2: all members have roles
        2:  len(team) >= 1 and all(m.get("role") for m in team),
        # Req 3: clear link to company KPI
        3:  bool(proj.get("company_kpi_link")) or bool(kpi.get("kpi_name")),
        # Req 4: historical data shown — baseline set OR historical_context filled
        4:  base != 0 or bool(kpi.get("historical_context")),
        # Req 5: KPI subdivided into components
        5:  has_subs,
        # Req 6: master plan visible — at least 1 step
        6:  len(steps) >= 1,
        # Req 7: meetings held
        7:  len(meets) >= 1,
        # ── WEEK 2 ──────────────────────────────────────────────────────────
        # Req 8: target clearly shown
        8:  tgt != 0 and bool(kpi.get("target_date")),
        # Req 9: target of each step clear
        9:  len(steps) >= 1 and all(s.get("planned_end_week") for s in steps),
        # Req 10: data collection consistent — meeting with >=80% attendance
        10: any(int(m.get("attendance_pct") or 0) >= 80 for m in meets) or len(meets) >= 2,
        # ── WEEK 3 ──────────────────────────────────────────────────────────
        # Req 11: step targets subdivided into activities — any step progress logged
        11: any(bool(_pj(wu.get("step_progress"), [])) for wu in wu_rows),
        # Req 12: methodology understood — team has roles + at least 1 meeting
        12: len(team) >= 1 and all(m.get("role") for m in team) and len(meets) >= 1,
        # Req 13: random member can explain board — same proxy
        13: len(team) >= 1 and len(meets) >= 1,
        # ── WEEK 4 ──────────────────────────────────────────────────────────
        # Req 14: cost/benefit — baseline AND target set
        14: base != 0 and tgt != 0,
        # Req 15: RCA documented
        15: has_rca,
        # Req 16: basic conditions restored
        16: bool(stab.get("basic_conditions_done")),
        # Req 17: planned actions visible with dates
        17: len(actions) > 0 and any(a.get("target_date") for a in actions),
        # ── WEEK 5 ──────────────────────────────────────────────────────────
        # Req 18: causes verified — root cause filled in 5-Why
        18: has_root,
        # Req 19: used route methods/tools — any RCA done
        19: has_rca,
        # Req 20: owner on each action
        20: len(actions) > 0 and all(a.get("owner") for a in actions),
        # Req 21: action plan up-to-date — actions exist
        21: len(actions) > 0,
        # Req 22: majority completed on time
        22: (bool(past_due) and len(on_time) / len(past_due) >= 0.5) if past_due else False,
        # Req 23: evidence of actions — OPLs or evidence text
        23: len(opls) > 0 or bool(stab.get("opl_evidence")),
        # Req 24: CIL standards + >=90% audit
        24: bool(stab.get("cil_standards_defined")) and float(stab.get("cil_audit_score") or 0) >= 90,
        # Req 25: 5S organized
        25: int(stab.get("five_s_rating") or 0) >= 3,
        # Req 26: improvements evident
        26: bool(stab.get("improvements_visible")) or len(opls) > 0,
        # ── WEEK 6 ──────────────────────────────────────────────────────────
        # Req 27: logical countermeasures — fishbone or 5-why with root cause
        27: has_rca and has_root,
        # ── WEEK 7 ──────────────────────────────────────────────────────────
        # Req 28: reoccurrence analysis
        28: any(_adata(a).get("reoccurrence") for a in analyses),
        # Req 29: single problem analysis
        29: has_root,
        # Req 30: KPI trend positive
        30: trend_ok,
        # Req 31: monitoring in place
        31: bool(stab.get("monitoring_in_place")),
        # ── WEEK 8 ──────────────────────────────────────────────────────────
        # Req 32: monitoring active and up-to-date
        32: bool(stab.get("monitoring_active")),
        # ── WEEK 10 ─────────────────────────────────────────────────────────
        # Req 33: procedures to hold gains
        33: bool(stab.get("procedures_created")),
        # Req 34: OPLs/SOPs created
        34: len(opls) > 0,
        # Req 35: training matrix in place
        35: bool(_pj(stab.get("training_matrix"), [])),
        # ── WEEK 11 ─────────────────────────────────────────────────────────
        # Req 36: goal achieved or substantial progress
        36: goal_ok,
    }

    # Batch upsert — suppress individual errors, show nothing to user
    rows = [
        {"project_id": pid, "req_id": req_id,
         "done": bool(done), "updated_by": "system",
         "updated_at": date.today().isoformat()}
        for req_id, done in results.items()
    ]
    try:
        supabase.table("fi_project_checklist").upsert(
            rows, on_conflict="project_id,req_id"
        ).execute()
    except Exception:
        # RLS or other error — fall back to one-by-one silently
        for row in rows:
            try:
                supabase.table("fi_project_checklist").upsert(
                    row, on_conflict="project_id,req_id"
                ).execute()
            except Exception:
                pass


def _mark_req(supabase, pid, req_id, done, user):
    try:
        supabase.table("fi_project_checklist").upsert({
            "project_id":pid,"req_id":req_id,
            "done":done,"updated_by":user,
            "updated_at":date.today().isoformat(),
        }, on_conflict="project_id,req_id").execute()
    except Exception as e:
        st.error(f"Could not save: {e}")

def _auto_mark(supabase, pid, form_key, done, user):
    """Auto-mark all requirements of a form when form is saved."""
    for r in REQ_BY_FORM.get(form_key,[]):
        _mark_req(supabase, pid, r["id"], done, user)

def _req_done_count(checklist, form_key):
    reqs = REQ_BY_FORM.get(form_key,[])
    done = sum(1 for r in reqs if checklist.get(r["id"],{}).get("done"))
    return done, len(reqs)

def _score(checklist):
    return sum(r["pts"] for r in REQUIREMENTS if checklist.get(r["id"],{}).get("done"))


# ─────────────────────────────────────────────────────────────────────────────
# CSS  (injected once)
# ─────────────────────────────────────────────────────────────────────────────
_CSS = """
<style>
.req-card {
    border-radius:10px; padding:14px 16px 10px; margin-bottom:6px;
    border-left:4px solid; background:#fff;
    box-shadow:0 1px 5px rgba(0,0,0,0.06); position:relative;
}
.req-done  { border-left-color:#1E8449!important; background:#F0FFF4!important; }
.req-open  { border-left-color:#0C5595!important; }
.req-late  { border-left-color:#DE201B!important; background:#FFF5F5!important; }
.req-future{ border-left-color:#BDC3C7!important; background:#F8FAFC!important; }
.pts-pill  { display:inline-block;padding:2px 9px;border-radius:12px;
             font-size:11px;font-weight:700;background:#EBF8FF;color:#0C5595;margin-left:6px;}
.link-chip { display:inline-block;padding:3px 9px;border-radius:10px;font-size:11px;
             font-weight:600;background:#F4F6F8;color:#0C5595;margin:2px;
             border:1px solid #D0D9E8; }
.form-btn  { display:inline-block;padding:5px 14px;border-radius:8px;font-size:12px;
             font-weight:700;cursor:pointer;border:none;margin-top:8px; }
.week-dot  { display:inline-block;width:7px;height:7px;border-radius:50%;margin:0 1px;vertical-align:middle; }
.section-hdr { font-size:17px;font-weight:800;color:#0C5595;
               border-left:4px solid #DE201B;padding-left:10px;
               margin:22px 0 14px; }
</style>
"""

# ─────────────────────────────────────────────────────────────────────────────
# SCORE + RAMP CHART
# ─────────────────────────────────────────────────────────────────────────────
def _ramp_fig(checklist, cw):
    wk_score = {}
    cum = 0
    for w in range(1,13):
        for r in REQ_BY_WEEK.get(w,[]):
            if checklist.get(r["id"],{}).get("done"): cum += r["pts"]
        if w in ACTIVE_WEEKS: wk_score[w] = cum

    fig, ax = plt.subplots(figsize=(10,3))
    fig.patch.set_facecolor("#fff"); ax.set_facecolor("#FAFAFA")
    tw = list(range(1,13))
    ax.fill_between(tw,[TARGET_RAMP[w] for w in tw],alpha=0.05,color="#566573")
    ax.plot(tw,[TARGET_RAMP[w] for w in tw],"s--",color="#BDC3C7",lw=1.4,ms=4,label="Target")
    if wk_score:
        ws=sorted(wk_score); wv=[wk_score[w] for w in ws]
        ax.fill_between(ws,wv,alpha=0.10,color="#0C5595")
        ax.plot(ws,wv,"o-",color="#0C5595",lw=2.5,ms=7,label="Actual",zorder=5)
        for w,v in zip(ws,wv):
            ax.annotate(str(v),(w,v),textcoords="offset points",xytext=(0,8),
                        fontsize=8,ha="center",color="#0C5595",fontweight="bold")
    ax.axvline(cw,color="#DE201B",lw=1.5,ls=":",alpha=0.7)
    ax.set_xlim(0.5,12.5); ax.set_ylim(0,110)
    ax.set_xticks(range(1,13)); ax.set_xticklabels([f"W{i}" for i in range(1,13)],fontsize=8.5)
    ax.set_ylabel("Score",fontsize=9,color="#566573")
    ax.legend(fontsize=9,frameon=False)
    ax.grid(axis="y",color="#eee",lw=0.5)
    for s in ["top","right"]: ax.spines[s].set_visible(False)
    fig.tight_layout(pad=0.5); return fig


# ─────────────────────────────────────────────────────────────────────────────
# MEETING MINUTES REPORT (PDF)
# ─────────────────────────────────────────────────────────────────────────────
def _meeting_pdf(meetings, project):
    W, H = 595, 842   # A4
    buf  = io.BytesIO()
    c    = rl_canvas.Canvas(buf, pagesize=(W, H))

    def _page_hdr(c, proj_name, week):
        c.setFillColor(HexColor("#0C5595")); c.rect(0,H-60,W,60,fill=1,stroke=0)
        c.setFillColor(HexColor("#ffffff")); c.setFont("Helvetica-Bold",16)
        c.drawString(24, H-38, f"Meeting Minutes — {proj_name}")
        c.setFont("Helvetica",10); c.drawString(24, H-54, f"Week {week} Report")
        c.setFillColor(HexColor("#DE201B")); c.rect(0,H-64,W,4,fill=1,stroke=0)

    for mi, mtg in enumerate(meetings):
        if mi > 0: c.showPage()
        _page_hdr(c, project.get("project_name","FI Project"), mtg.get("week_number","—"))

        y = H - 90
        def _row(label, value, indent=24):
            nonlocal y
            c.setFont("Helvetica-Bold",9); c.setFillColor(HexColor("#64748B"))
            c.drawString(indent, y, label.upper())
            c.setFont("Helvetica",10); c.setFillColor(HexColor("#1E293B"))
            c.drawString(indent+120, y, str(value)[:90])
            y -= 18

        _row("Date", mtg.get("meeting_date","—"))
        _row("Week", mtg.get("week_number","—"))
        _row("Attendance", mtg.get("attendees","—"))
        _row("Attendance %", f"{mtg.get('attendance_pct','—')}%")
        y -= 8
        c.setFillColor(HexColor("#E2E8F0")); c.rect(24,y,W-48,1,fill=1,stroke=0); y -= 16

        # Agenda
        c.setFont("Helvetica-Bold",10); c.setFillColor(HexColor("#0C5595"))
        c.drawString(24,y,"Agenda Items"); y -= 16
        for item in _pj(mtg.get("agenda"),[]):
            c.setFont("Helvetica",9); c.setFillColor(HexColor("#334155"))
            c.drawString(36, y, f"• {str(item)[:90]}"); y -= 14
        y -= 8

        # Notes
        c.setFont("Helvetica-Bold",10); c.setFillColor(HexColor("#0C5595"))
        c.drawString(24,y,"Discussion Notes"); y -= 14
        notes = str(mtg.get("notes","—"))
        import textwrap
        for line in textwrap.wrap(notes, width=95)[:20]:
            c.setFont("Helvetica",9); c.setFillColor(HexColor("#334155"))
            c.drawString(36,y,line); y -= 13
        y -= 8

        # Action items
        c.setFont("Helvetica-Bold",10); c.setFillColor(HexColor("#0C5595"))
        c.drawString(24,y,"Actions Raised"); y -= 14
        for act in _pj(mtg.get("actions_raised"),[]):
            c.setFont("Helvetica",9); c.setFillColor(HexColor("#334155"))
            c.drawString(36,y,f"• {str(act)[:90]}"); y -= 13
        y -= 8

        # Next steps
        c.setFont("Helvetica-Bold",10); c.setFillColor(HexColor("#0C5595"))
        c.drawString(24,y,"Next Steps"); y -= 14
        for ns in _pj(mtg.get("next_steps"),[]):
            c.setFont("Helvetica",9); c.setFillColor(HexColor("#334155"))
            c.drawString(36,y,f"• {str(ns)[:90]}"); y -= 13

        # Footer
        c.setFillColor(HexColor("#94A3B8")); c.setFont("Helvetica",8)
        c.drawString(24,28,f"Generated {date.today().isoformat()} · NPS Hub · FI Pillar")
        c.setFillColor(HexColor("#E2E8F0")); c.rect(0,20,W,1,fill=1,stroke=0)

    c.save(); buf.seek(0); return buf


# ─────────────────────────────────────────────────────────────────────────────
# FORM RENDERERS  — each opens inline below its requirement card
# ─────────────────────────────────────────────────────────────────────────────

def _form_team(supabase, pid, project, checklist, cw, name, can_edit):
    """Team setup — list, roles, methodology understanding."""
    st.markdown('<div class="section-hdr">👥 Team Setup</div>', unsafe_allow_html=True)
    try:
        team = supabase.table("fi_project_team").select("*").eq("project_id",pid).execute().data or []
    except: team = []

    if team:
        df = pd.DataFrame([{"Name":m["member_name"],"Role":m.get("role",""),"Department":m.get("department","")} for m in team])
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("No team members yet. Add below.")

    if can_edit:
        with st.expander("➕ Add team member"):
            with st.form("fi_team_add"):
                c1,c2,c3 = st.columns(3)
                t_name = c1.text_input("Name *")
                t_role = c2.selectbox("Role", TEAM_ROLES)
                t_dept = c3.text_input("Department")
                if st.form_submit_button("Add", type="primary"):
                    if t_name:
                        supabase.table("fi_project_team").insert({
                            "project_id":pid,"member_name":t_name,
                            "role":t_role,"department":t_dept,"contribution_target":""
                        }).execute()
                        _auto_mark(supabase,pid,"team",True,name); st.rerun()
                    else: st.error("Name required.")

        if team:
            del_m = st.selectbox("Remove member",["—"]+[m["member_name"] for m in team],key="fi_del_m_team")
            if st.button("Remove",key="fi_rem_team") and del_m!="—":
                mid = next((m["id"] for m in team if m["member_name"]==del_m),None)
                if mid:
                    supabase.table("fi_project_team").delete().eq("id",mid).execute()
                    st.rerun()

    # Methodology understanding — req 12,13 (auditor confirms)
    if is_auditor := True:
        st.markdown("**Methodology Understanding** *(req. 12 & 13)*")
        with st.form("fi_meth"):
            m_ok  = st.checkbox("Team demonstrated understanding of methodology",
                                value=bool(checklist.get(12,{}).get("done")))
            mb_ok = st.checkbox("Random member can explain the activity board",
                                value=bool(checklist.get(13,{}).get("done")))
            m_note = st.text_area("Notes", height=60,
                                  value=checklist.get(12,{}).get("note",""))
            if st.form_submit_button("Save", type="primary") and can_edit:
                _mark_req(supabase,pid,12,m_ok,name)
                _mark_req(supabase,pid,13,mb_ok,name)
                st.success("Saved"); st.rerun()


def _form_kpi_setup(supabase, pid, project, checklist, cw, name, can_edit):
    """KPI baseline, target, components, trend."""
    st.markdown('<div class="section-hdr">📈 KPI & Results</div>', unsafe_allow_html=True)
    try:
        kpi_rows = supabase.table("fi_project_kpi").select("*").eq("project_id",pid).execute().data or []
        kpi = kpi_rows[0] if kpi_rows else {}
    except: kpi={}

    KPI_CATS=["OEE Improvement","Quality Defect Reduction","Waste Reduction","Cost Reduction",
              "Safety Improvement","Delivery Performance","5S Score Improvement","Throughput/Productivity"]
    UNITS=[ "%","MT","SAR","K SAR","LM","SQM","Count","Score","Hours","Mins"]

    with st.form("fi_kpi_form"):
        k1,k2 = st.columns(2)
        kpi_cat  = k1.selectbox("KPI Category", KPI_CATS,
            index=KPI_CATS.index(kpi.get("kpi_category",KPI_CATS[0])) if kpi.get("kpi_category") in KPI_CATS else 0)
        kpi_unit = k2.selectbox("Unit", UNITS,
            index=UNITS.index(kpi.get("unit","%")) if kpi.get("unit") in UNITS else 0)
        k3,k4,k5 = st.columns(3)
        kpi_base = k3.number_input("Baseline Value", value=float(kpi.get("baseline_value",0) or 0))
        kpi_tgt  = k4.number_input("Target Value",   value=float(kpi.get("target_value",0)   or 0))
        kpi_date = k5.date_input("Target Date",
            value=date.fromisoformat(str(kpi.get("target_date",date.today()+timedelta(weeks=12)))[:10])
                  if kpi.get("target_date") else date.today()+timedelta(weeks=12))
        kpi_link = st.text_input("Company KPI Link",value=project.get("company_kpi_link",""))
        kpi_hist = st.text_area("Historical context (timeframe & prior values)",
            value=kpi.get("historical_context",""), height=60,
            placeholder="e.g. OEE was 63% in Jan, 62% Feb, 61% Mar — trending down over 3 months")
        # Sub-components
        st.caption("**KPI Components** (one per line, e.g. 'Availability, Performance, Quality')")
        _ex_subs = _pj(kpi.get("sub_components"),[])
        _ex_str  = "\n".join(s.get("name","") if isinstance(s,dict) else str(s) for s in _ex_subs)
        subs_raw = st.text_area("Components", value=_ex_str, height=70,
                                placeholder="Availability\nPerformance\nQuality")

        # KPI weekly readings
        st.divider()
        st.caption("**Weekly KPI Readings** — enter the latest value for the current week")
        try:
            wu_rows = supabase.table("fi_weekly_updates").select("*").eq("project_id",pid).order("week_number").execute().data or []
        except: wu_rows=[]
        wu_by_week = {w["week_number"]:w for w in wu_rows}
        cur_wu = wu_by_week.get(cw,{})
        kv = st.number_input(f"W{cw} KPI Value ({kpi_unit if kpi.get('unit') else ''})",
                             value=float(cur_wu.get("kpi_value",kpi_base) or kpi_base))

        if st.form_submit_button("Save KPI Setup", type="primary") and can_edit:
            subs = [{"name":s.strip()} for s in subs_raw.split("\n") if s.strip()]
            payload={"project_id":pid,"kpi_name":kpi_cat,"kpi_category":kpi_cat,
                     "unit":kpi_unit,"baseline_value":kpi_base,"target_value":kpi_tgt,
                     "target_date":str(kpi_date),"sub_components":json.dumps(subs),
                     "historical_context":kpi_hist}
            if kpi.get("id"):
                update_payload = {k:v for k,v in payload.items() if k != "project_id"}
                supabase.table("fi_project_kpi").update(update_payload).eq("id",kpi["id"]).execute()
            else:
                supabase.table("fi_project_kpi").insert(payload).execute()
            # save KPI link on project
            supabase.table("fi_projects").update({"company_kpi_link":kpi_link}).eq("id",pid).execute()
            # save weekly reading
            wu_payload={"project_id":pid,"week_number":cw,"kpi_value":kv,"updated_by":name}
            if cur_wu.get("id"):
                wu_update = {k:v for k,v in wu_payload.items() if k != "project_id"}
                supabase.table("fi_weekly_updates").update(wu_update).eq("id",cur_wu["id"]).execute()
            else:
                supabase.table("fi_weekly_updates").insert(wu_payload).execute()
            # auto-mark related requirements
            _mark_req(supabase,pid,3, bool(kpi_link), name)
            _mark_req(supabase,pid,4, bool(kpi_hist) and kpi_base!=0, name)
            _mark_req(supabase,pid,5, len(subs)>0, name)
            _mark_req(supabase,pid,8, kpi_tgt!=0, name)
            # trend positive? (need 3+ readings going right way)
            readings = sorted([w for w in wu_rows if w.get("kpi_value") is not None],
                              key=lambda x:x["week_number"])
            trend_ok = len(readings)>=3 and (
                (kpi_tgt>kpi_base and readings[-1]["kpi_value"]>readings[-3]["kpi_value"]) or
                (kpi_tgt<kpi_base and readings[-1]["kpi_value"]<readings[-3]["kpi_value"])
            )
            _mark_req(supabase,pid,30,trend_ok,name)
            # goal achieved?
            goal_ok = False
            if readings and kpi_base!=0 and kpi_tgt!=0:
                gap  = abs(kpi_tgt-kpi_base)
                prog = abs(readings[-1]["kpi_value"]-kpi_base)/gap if gap else 0
                goal_ok = prog >= 0.80
            _mark_req(supabase,pid,36,goal_ok,name)
            st.success("Saved ✓"); st.rerun()

    # Show trend chart if data exists
    try:
        wu_rows2 = supabase.table("fi_weekly_updates").select("*").eq("project_id",pid)\
            .order("week_number").execute().data or []
    except: wu_rows2=[]
    if wu_rows2 and kpi.get("baseline_value") is not None:
        sorted_wu = [w for w in wu_rows2 if w.get("kpi_value") is not None]
        if sorted_wu:
            fig,ax = plt.subplots(figsize=(9,3))
            fig.patch.set_facecolor("#fff"); ax.set_facecolor("#FAFAFA")
            wks = [w["week_number"] for w in sorted_wu]
            vals= [float(w["kpi_value"]) for w in sorted_wu]
            base= float(kpi.get("baseline_value",0))
            tgt = float(kpi.get("target_value",0))
            ax.axhline(base,color="#BDC3C7",lw=1.5,ls=":",label=f"Baseline {base}")
            ax.axhline(tgt, color="#1E8449",lw=1.5,ls="--",label=f"Target {tgt}")
            ax.fill_between(wks,base,vals,alpha=0.10,color="#0C5595")
            ax.plot(wks,vals,"o-",color="#0C5595",lw=2.5,ms=8,label="Actual",zorder=5)
            for w,v in zip(wks,vals):
                ax.annotate(f"{v:.1f}",(w,v),textcoords="offset points",
                            xytext=(0,9),fontsize=8,ha="center",color="#0C5595",fontweight="bold")
            ax.set_xlim(0.5,12.5); ax.set_xticks(range(1,13))
            ax.set_xticklabels([f"W{i}" for i in range(1,13)],fontsize=8)
            ax.legend(fontsize=9,frameon=False)
            ax.grid(axis="y",color="#eee",lw=0.5)
            for s in ["top","right"]: ax.spines[s].set_visible(False)
            fig.tight_layout(pad=0.4)
            st.pyplot(fig,use_container_width=True); plt.close(fig)


def _form_master_plan(supabase, pid, project, checklist, cw, name, can_edit):
    """Steps / Gantt."""
    st.markdown('<div class="section-hdr">🗓 Master Plan</div>', unsafe_allow_html=True)
    try:
        steps   = supabase.table("fi_project_steps").select("*").eq("project_id",pid).order("sort_order").execute().data or []
        wu_rows = supabase.table("fi_weekly_updates").select("*").eq("project_id",pid).execute().data or []
        team    = supabase.table("fi_project_team").select("*").eq("project_id",pid).execute().data or []
    except: steps=[]; wu_rows=[]; team=[]

    if steps:
        # Gantt
        n=len(steps); fig_h=max(3.0,n*0.52+1.0)
        fig,ax=plt.subplots(figsize=(12,fig_h)); fig.patch.set_facecolor("#fff"); ax.set_facecolor("#FAFAFA")
        sp_map={}
        for wu in wu_rows:
            for sp in _pj(wu.get("step_progress"),[]):
                if isinstance(sp,dict):
                    sid=sp.get("step_id",""); sp_map[sid]=max(sp_map.get(sid,0),sp.get("pct_complete",0))
        for i,step in enumerate(steps):
            ps=max(1,step.get("planned_start_week",1)); pe=min(12,step.get("planned_end_week",ps)); dur=pe-ps+1
            ax.barh(i,dur,left=ps-1,height=0.55,color="#d6e8f7",zorder=2)
            ax.barh(i,dur,left=ps-1,height=0.55,color="none",edgecolor="#0C5595",linewidth=1.0,zorder=3)
            pct=sp_map.get(str(step.get("id","")),0)
            if pct>0:
                ax.barh(i,dur*pct/100,left=ps-1,height=0.55,
                        color="#1E8449" if pct==100 else "#0C5595",alpha=0.85,zorder=4)
                if pct>=15:
                    ax.text(ps-1+dur*pct/200,i,f"{pct}%",ha="center",va="center",
                            fontsize=7.5,color="white",fontweight="bold",zorder=5)
            if step.get("owner"):
                ax.text(pe+0.1,i,step["owner"],ha="left",va="center",fontsize=7,color="#566573",zorder=5)
        for w in range(13): ax.axvline(w,color="#eee",linewidth=0.6,zorder=1)
        ax.axvline(cw-1,color="#DE201B",linewidth=2,linestyle="--",zorder=6,alpha=0.9)
        ax.set_yticks(range(n)); ax.set_yticklabels([s.get("step_name","") for s in steps],fontsize=9)
        ax.set_xticks(range(13)); ax.set_xticklabels([""]+[f"W{i}" for i in range(1,13)],fontsize=8)
        ax.set_xlim(-0.1,12.5); ax.set_ylim(-0.8,n-0.2); ax.invert_yaxis()
        for sp2 in ["top","right","left","bottom"]: ax.spines[sp2].set_visible(False)
        ax.tick_params(left=False,bottom=False)
        fig.tight_layout(pad=0.6)
        st.pyplot(fig,use_container_width=True); plt.close(fig)
    else:
        st.info("No steps yet. Add below.")

    if can_edit:
        with st.expander("➕ Add step"):
            with st.form("fi_step_add"):
                s1,s2=st.columns(2)
                s_name=s1.text_input("Step Name *"); s_desc=s2.text_input("Description")
                sw1,sw2,sw3=st.columns(3)
                s_start=sw1.number_input("Start Week",min_value=1,max_value=12,value=1)
                s_end  =sw2.number_input("End Week",  min_value=1,max_value=12,value=2)
                s_owner=sw3.selectbox("Owner",[""]+[m["member_name"] for m in team])
                if st.form_submit_button("Add Step",type="primary"):
                    if s_name:
                        supabase.table("fi_project_steps").insert({
                            "project_id":pid,"step_name":s_name,"description":s_desc,
                            "planned_start_week":int(s_start),"planned_end_week":int(s_end),
                            "owner":s_owner,"sort_order":len(steps)
                        }).execute()
                        _mark_req(supabase,pid,6,True,name)
                        _mark_req(supabase,pid,9,True,name)
                        st.rerun()

        if steps:
            del_s=st.selectbox("Remove step",["—"]+[s["step_name"] for s in steps],key="fi_del_step")
            if st.button("Remove",key="fi_rem_step") and del_s!="—":
                sid=next((s["id"] for s in steps if s["step_name"]==del_s),None)
                if sid: supabase.table("fi_project_steps").delete().eq("id",sid).execute(); st.rerun()

    # Step progress update
    if steps:
        with st.expander(f"📊 Update progress — Week {cw}"):
            try:
                cur_wu = (supabase.table("fi_weekly_updates").select("*").eq("project_id",pid)
                          .eq("week_number",cw).execute().data or [{}])[0]
            except: cur_wu={}
            due_steps=[s for s in steps if s.get("planned_start_week",1)<=cw]
            sp_list=_pj(cur_wu.get("step_progress"),[])
            sp_map2={sp["step_id"]:sp for sp in sp_list if isinstance(sp,dict)}
            with st.form(f"fi_sp_w{cw}"):
                new_sp=[]; PCT=["0%","25%","50%","75%","100%"]
                for step in due_steps:
                    sid=str(step["id"]); ex=sp_map2.get(sid,{})
                    sa,sb=st.columns([2,1])
                    sa.markdown(f"**{step['step_name']}**")
                    cur_p=f"{int(ex.get('pct_complete',0))}%"
                    psel=sb.selectbox("",PCT,index=PCT.index(cur_p) if cur_p in PCT else 0,
                                      key=f"spp_{sid}",label_visibility="collapsed")
                    new_sp.append({"step_id":sid,"pct_complete":int(psel.replace("%","")),
                                   "status":"Completed" if psel=="100%" else "In Progress" if psel!="0%" else "Not Started"})
                if st.form_submit_button("Save Progress",type="primary") and can_edit:
                    wu_p={"project_id":pid,"week_number":cw,"step_progress":json.dumps(new_sp),"updated_by":name}
                    if cur_wu.get("id"):
                        supabase.table("fi_weekly_updates").update(wu_p).eq("id",cur_wu["id"]).execute()
                    else:
                        supabase.table("fi_weekly_updates").insert(wu_p).execute()
                    _mark_req(supabase,pid,11,True,name)
                    st.success("Saved"); st.rerun()


def _form_meeting(supabase, pid, project, checklist, cw, name, can_edit):
    """Meeting minutes — every week, printable."""
    st.markdown('<div class="section-hdr">📋 Meeting Minutes</div>', unsafe_allow_html=True)
    try:
        team     = supabase.table("fi_project_team").select("*").eq("project_id",pid).execute().data or []
        meetings = supabase.table("fi_meetings").select("*").eq("project_id",pid)\
            .order("week_number",desc=True).execute().data or []
    except: team=[]; meetings=[]

    # Past meetings list
    if meetings:
        for m in meetings[:5]:
            att_pct = m.get("attendance_pct","?")
            with st.expander(f"📋 Week {m.get('week_number','?')} — {m.get('meeting_date','')[:10]} — {m.get('attendees','?')[:40]}"):
                st.markdown(f"**Attendance:** {m.get('attendees','—')} ({att_pct}%)")
                st.markdown(f"**Notes:** {m.get('notes','—')}")
                for label, key in [("Agenda","agenda"),("Actions Raised","actions_raised"),("Next Steps","next_steps")]:
                    items=_pj(m.get(key),[])
                    if items:
                        st.markdown(f"**{label}:**")
                        for it in items: st.markdown(f"  - {it}")
        # Download all meetings as PDF
        pdf_buf = _meeting_pdf(meetings, project)
        st.download_button(
            "📥 Download All Meeting Minutes (PDF)",
            data=pdf_buf,
            file_name=f"MeetingMinutes_{project.get('project_name','FI').replace(' ','_')}.pdf",
            mime="application/pdf"
        )
    else:
        st.info("No meetings logged yet.")

    # Log new meeting
    if can_edit:
        with st.expander(f"➕ Log Week {cw} Meeting", expanded=not bool(meetings)):
            with st.form("fi_meeting_form"):
                m1,m2=st.columns(2)
                m_date=m1.date_input("Meeting Date",value=date.today())
                m_attendees=m2.multiselect("Attendees",[t["member_name"] for t in team])
                m_agenda_raw=st.text_area("Agenda (one item per line)",height=80)
                m_notes=st.text_area("Discussion Notes",height=100)
                m_actions_raw=st.text_area("Actions Raised (one per line)",height=70)
                m_next_raw=st.text_area("Next Steps (one per line)",height=70)
                if st.form_submit_button("Save Meeting Minutes",type="primary"):
                    att_pct=round(len(m_attendees)/max(len(team),1)*100)
                    agenda=[l.strip() for l in m_agenda_raw.split("\n") if l.strip()]
                    actions=[l.strip() for l in m_actions_raw.split("\n") if l.strip()]
                    nexts=[l.strip() for l in m_next_raw.split("\n") if l.strip()]
                    supabase.table("fi_meetings").insert({
                        "project_id":pid,"week_number":cw,
                        "meeting_date":str(m_date),
                        "attendees":", ".join(m_attendees),
                        "attendance_pct":att_pct,
                        "notes":m_notes,
                        "agenda":json.dumps(agenda),
                        "actions_raised":json.dumps(actions),
                        "next_steps":json.dumps(nexts),
                        "created_by":name,
                    }).execute()
                    _mark_req(supabase,pid,7,True,name)
                    _mark_req(supabase,pid,10,att_pct>=80,name)
                    st.success("Saved ✓"); st.rerun()


def _form_rca(supabase, pid, project, checklist, cw, name, can_edit):
    """Root Cause Analysis — 5-Why + Fishbone, saved as JSON."""
    st.markdown('<div class="section-hdr">🔍 Root Cause Analysis</div>', unsafe_allow_html=True)
    try:
        analyses = supabase.table("fi_project_analysis").select("*")\
            .eq("project_id",pid).in_("analysis_type",["5why","fishbone"])\
            .order("created_at",desc=True).execute().data or []
    except: analyses=[]

    if analyses:
        for a in analyses:
            d=_pj(a["data"],{}) if isinstance(a["data"],str) else a["data"]
            atype=a.get("analysis_type",""); wk=a.get("week_number","?")
            with st.expander(f"{'5-Why' if atype=='5why' else 'Fishbone'} — W{wk} · {a.get('created_by','')} · {str(a.get('created_at',''))[:10]}"):
                if atype=="5why":
                    st.markdown(f"**Problem:** {d.get('problem','')}")
                    for i,w in enumerate(d.get("whys",[]),1):
                        if w: st.markdown(f"**Why {i}:** {w}")
                    if d.get("root_cause"): st.markdown(f"**Root Cause:** {d.get('root_cause')}")
                else:
                    st.markdown(f"**Effect:** {d.get('problem','')}")
                    for cat,causes in d.get("categories",{}).items():
                        if causes: st.markdown(f"**{cat}:** {', '.join(causes)}")
                if can_edit:
                    if st.button("🗑 Delete",key=f"del_rca_{a['id']}"):
                        supabase.table("fi_project_analysis").delete().eq("id",a["id"]).execute(); st.rerun()

    if can_edit:
        tool=st.radio("Analysis Tool",["5-Why","Fishbone"],horizontal=True,key="rca_tool_sel")
        prob_default=project.get("problem_statement","")

        if tool=="5-Why":
            with st.form("fi_5why"):
                prob=st.text_area("Problem Statement",value=prob_default,height=70)
                whys=[]
                for i in range(5):
                    whys.append(st.text_input(f"Why {i+1}",key=f"why_{i}"))
                root=st.text_input("Root Cause (summary)")
                if st.form_submit_button("Save 5-Why",type="primary"):
                    active=[w for w in whys if w.strip()]
                    if prob and active:
                        supabase.table("fi_project_analysis").insert({
                            "project_id":pid,"week_number":cw,"analysis_type":"5why",
                            "data":json.dumps({"problem":prob,"whys":whys,"root_cause":root}),
                            "created_by":name,
                        }).execute()
                        _mark_req(supabase,pid,15,True,name)
                        _mark_req(supabase,pid,18,bool(root),name)
                        _mark_req(supabase,pid,19,True,name)
                        _mark_req(supabase,pid,27,bool(root),name)
                        _mark_req(supabase,pid,28,True,name)
                        _mark_req(supabase,pid,29,True,name)
                        st.success("Saved ✓"); st.rerun()
                    else: st.warning("Fill problem and at least one Why.")
        else:
            with st.form("fi_fishbone"):
                prob=st.text_area("Effect / Problem",value=prob_default[:80],height=60)
                st.caption("Enter causes per category (one per line)")
                cats={}
                DEFAULT_CATS=["Man","Machine","Method","Material","Measurement","Environment"]
                cols2=st.columns(3)
                for ci,cat in enumerate(DEFAULT_CATS):
                    raw=cols2[ci%3].text_area(cat,height=90,key=f"fb_{cat}",
                                              placeholder="One cause per line")
                    cats[cat]=[l.strip() for l in raw.split("\n") if l.strip()]
                if st.form_submit_button("Save Fishbone",type="primary"):
                    filled={k:v for k,v in cats.items() if v}
                    if prob and filled:
                        supabase.table("fi_project_analysis").insert({
                            "project_id":pid,"week_number":cw,"analysis_type":"fishbone",
                            "data":json.dumps({"problem":prob,"categories":cats}),
                            "created_by":name,
                        }).execute()
                        _mark_req(supabase,pid,15,True,name)
                        _mark_req(supabase,pid,27,True,name)
                        st.success("Saved ✓"); st.rerun()
                    else: st.warning("Fill problem and at least one category.")


def _form_actions(supabase, pid, project, checklist, cw, name, can_edit):
    """Action plan."""
    st.markdown('<div class="section-hdr">✅ Action Plan</div>', unsafe_allow_html=True)
    try:
        actions = supabase.table("fi_actions").select("*").eq("project_id",pid).order("created_at").execute().data or []
        team    = supabase.table("fi_project_team").select("*").eq("project_id",pid).execute().data or []
    except: actions=[]; team=[]

    if actions:
        SCOLS={"Open":"#0C5595","In Progress":"#D68910","Completed":"#1E8449","Overdue":"#DE201B"}
        for a in actions:
            is_od=(a.get("target_date") and _safe_date(a["target_date"]) and
                   _safe_date(a["target_date"])<date.today() and a.get("status")!="Completed")
            status="Overdue" if is_od else a.get("status","Open")
            sc=SCOLS.get(status,"#566573")
            st.markdown(f"""
            <div class="req-card" style="border-left-color:{sc}">
              <b>{a.get('description','')}</b>
              <span style="float:right;font-size:11px;font-weight:700;color:{sc}">{status}</span><br>
              <span style="font-size:12px;color:#64748B">👤 {a.get('owner','—')} &nbsp; 📅 {str(a.get('target_date',''))[:10]}</span>
            </div>""", unsafe_allow_html=True)
            if can_edit:
                uc1,uc2=st.columns([2,3])
                new_st=uc1.selectbox("",ACTION_STATUS,
                    index=ACTION_STATUS.index(a.get("status","Open")) if a.get("status") in ACTION_STATUS else 0,
                    key=f"act_st_{a['id']}",label_visibility="collapsed")
                if uc2.button("Update",key=f"act_upd_{a['id']}"):
                    supabase.table("fi_actions").update({"status":new_st}).eq("id",a["id"]).execute()
                    _auto_mark_actions(supabase,pid,actions,name); st.rerun()
    else:
        st.info("No actions yet.")

    if can_edit:
        with st.expander("➕ Add action"):
            with st.form("fi_act_add"):
                a1,a2=st.columns(2)
                a_desc=a1.text_area("Description *",height=70)
                a_rca =a2.text_area("Root Cause Addressed",height=70)
                a3,a4=st.columns(2)
                a_own=a3.selectbox("Owner",[""]+[m["member_name"] for m in team])
                a_date=a4.date_input("Target Date",value=date.today()+timedelta(weeks=2))
                if st.form_submit_button("Add Action",type="primary"):
                    if a_desc:
                        supabase.table("fi_actions").insert({
                            "project_id":pid,"description":a_desc,
                            "root_cause_addressed":a_rca,"owner":a_own,
                            "target_date":str(a_date),"status":"Open","created_week":cw,
                        }).execute()
                        _auto_mark_actions(supabase,pid,
                            supabase.table("fi_actions").select("*").eq("project_id",pid).execute().data or [],
                            name)
                        st.rerun()
                    else: st.error("Description required.")

def _auto_mark_actions(supabase, pid, actions, name):
    has_actions  = len(actions) > 0
    has_owners   = has_actions and all(a.get("owner") for a in actions)
    has_dates    = has_actions and any(a.get("target_date") for a in actions)
    past_due     = [a for a in actions if _safe_date(a.get("target_date")) and _safe_date(a["target_date"])<date.today()]
    on_time      = [a for a in past_due if a.get("status")=="Completed"]
    majority_ok  = bool(past_due) and len(on_time)/len(past_due)>=0.5
    _mark_req(supabase,pid,17,has_actions and has_dates,name)
    _mark_req(supabase,pid,20,has_owners,name)
    _mark_req(supabase,pid,21,has_actions,name)
    _mark_req(supabase,pid,22,majority_ok,name)


def _form_cil(supabase, pid, project, checklist, cw, name, can_edit):
    st.markdown('<div class="section-hdr">🔧 CIL & Restoration</div>', unsafe_allow_html=True)
    try:
        stab_rows=supabase.table("fi_stabilisation").select("*").eq("project_id",pid).execute().data or []
        stab=stab_rows[0] if stab_rows else {}
    except: stab={}
    with st.form("fi_cil_form"):
        basic_done=st.checkbox("Critical areas restored to basic conditions",value=bool(stab.get("basic_conditions_done")))
        basic_desc=st.text_area("Describe what was done",value=stab.get("basic_conditions_desc",""),height=70)
        cil_def=st.checkbox("CIL standards defined",value=bool(stab.get("cil_standards_defined")))
        cil_score=st.number_input("Latest CIL Audit Score (%)",min_value=0.0,max_value=100.0,
                                  value=float(stab.get("cil_audit_score",0) or 0))
        if cil_score>=90: st.success("✓ Meets ≥90% target")
        elif cil_score>0: st.warning(f"{cil_score:.0f}% — target ≥90%")
        if st.form_submit_button("Save",type="primary") and can_edit:
            payload={"project_id":pid,"basic_conditions_done":basic_done,
                     "basic_conditions_desc":basic_desc,
                     "cil_standards_defined":cil_def,"cil_audit_score":cil_score}
            if stab.get("id"):
                supabase.table("fi_stabilisation").update(payload).eq("id",stab["id"]).execute()
            else:
                supabase.table("fi_stabilisation").insert(payload).execute()
            _mark_req(supabase,pid,16,basic_done,name)
            _mark_req(supabase,pid,24,cil_def and cil_score>=90,name)
            st.success("Saved ✓"); st.rerun()


def _form_five_s(supabase, pid, project, checklist, cw, name, can_edit):
    st.markdown('<div class="section-hdr">⭐ 5S Audit</div>', unsafe_allow_html=True)
    try:
        stab_rows=supabase.table("fi_stabilisation").select("*").eq("project_id",pid).execute().data or []
        stab=stab_rows[0] if stab_rows else {}
    except: stab={}
    with st.form("fi_5s_form"):
        rating=st.slider("5S Rating",1,5,int(stab.get("five_s_rating",1) or 1))
        notes =st.text_area("5S observations",value=stab.get("five_s_notes",""),height=70)
        imp   =st.checkbox("Improvements on targeted areas evident",value=bool(stab.get("improvements_visible")))
        if st.form_submit_button("Save",type="primary") and can_edit:
            payload={"project_id":pid,"five_s_rating":rating,"five_s_notes":notes,"improvements_visible":imp}
            if stab.get("id"):
                supabase.table("fi_stabilisation").update(payload).eq("id",stab["id"]).execute()
            else:
                supabase.table("fi_stabilisation").insert(payload).execute()
            _mark_req(supabase,pid,25,rating>=3,name)
            _mark_req(supabase,pid,26,imp,name)
            st.success("Saved ✓"); st.rerun()


def _form_monitoring(supabase, pid, project, checklist, cw, name, can_edit):
    st.markdown('<div class="section-hdr">📊 Monitoring & Controls</div>', unsafe_allow_html=True)
    try:
        stab_rows=supabase.table("fi_stabilisation").select("*").eq("project_id",pid).execute().data or []
        stab=stab_rows[0] if stab_rows else {}
    except: stab={}
    with st.form("fi_mon_form"):
        mon_in  =st.checkbox("Monitoring systems in place and visible",value=bool(stab.get("monitoring_in_place")))
        mon_types=st.multiselect("Types of monitoring",
            ["Checklist","Audit","Form","Visual Board","SPC Chart","Other"],
            default=[x.strip() for x in (stab.get("monitoring_types","") or "").split(",") if x.strip()])
        mon_act =st.checkbox("Monitoring actively used and up-to-date",value=bool(stab.get("monitoring_active")))
        mon_date=st.date_input("Last updated",
            value=date.fromisoformat(str(stab["monitoring_last_update"])) if stab.get("monitoring_last_update") else date.today())
        mon_note=st.text_area("Evidence / description",value=stab.get("monitoring_notes",""),height=60)
        if st.form_submit_button("Save",type="primary") and can_edit:
            payload={"project_id":pid,"monitoring_in_place":mon_in,
                     "monitoring_types":",".join(mon_types),
                     "monitoring_active":mon_act,"monitoring_last_update":str(mon_date),
                     "monitoring_notes":mon_note}
            if stab.get("id"):
                supabase.table("fi_stabilisation").update(payload).eq("id",stab["id"]).execute()
            else:
                supabase.table("fi_stabilisation").insert(payload).execute()
            _mark_req(supabase,pid,31,mon_in,name)
            _mark_req(supabase,pid,32,mon_act,name)
            _mark_req(supabase,pid,33,mon_act,name)
            st.success("Saved ✓"); st.rerun()


def _form_opls(supabase, pid, project, checklist, cw, name, can_edit):
    st.markdown('<div class="section-hdr">📝 OPLs & SOPs</div>', unsafe_allow_html=True)
    try:
        stab_rows=supabase.table("fi_stabilisation").select("*").eq("project_id",pid).execute().data or []
        stab=stab_rows[0] if stab_rows else {}
    except: stab={}
    ex_opls=_pj(stab.get("opls"),[])
    with st.form("fi_opls_form"):
        n_opls=st.number_input("Number of OPLs/SOPs",min_value=0,max_value=20,value=max(0,len(ex_opls)))
        new_opls=[]
        for oi in range(int(n_opls)):
            eo=ex_opls[oi] if oi<len(ex_opls) else {}
            o1,o2=st.columns(2)
            new_opls.append({"title":o1.text_input(f"OPL {oi+1} Title",value=eo.get("title",""),key=f"ot_{oi}"),
                              "description":o2.text_input("Description",value=eo.get("description",""),key=f"od_{oi}")})
        proc_done=st.checkbox("Procedures in place to hold gains",value=bool(stab.get("procedures_created")))
        evidence =st.text_area("Evidence of implemented actions (links, descriptions)",
                               value=stab.get("opl_evidence",""),height=70)
        if st.form_submit_button("Save",type="primary") and can_edit:
            payload={"project_id":pid,"opls":json.dumps(new_opls),
                     "procedures_created":proc_done,"opl_evidence":evidence}
            if stab.get("id"):
                supabase.table("fi_stabilisation").update(payload).eq("id",stab["id"]).execute()
            else:
                supabase.table("fi_stabilisation").insert(payload).execute()
            _mark_req(supabase,pid,23,len(new_opls)>0 or bool(evidence),name)
            _mark_req(supabase,pid,26,len(new_opls)>0,name)
            _mark_req(supabase,pid,33,proc_done,name)
            _mark_req(supabase,pid,34,len(new_opls)>0,name)
            st.success("Saved ✓"); st.rerun()


def _form_training(supabase, pid, project, checklist, cw, name, can_edit):
    st.markdown('<div class="section-hdr">🎓 Training Matrix</div>', unsafe_allow_html=True)
    try:
        stab_rows=supabase.table("fi_stabilisation").select("*").eq("project_id",pid).execute().data or []
        stab=stab_rows[0] if stab_rows else {}
        team=supabase.table("fi_project_team").select("*").eq("project_id",pid).execute().data or []
    except: stab={}; team=[]
    ex_tm=_pj(stab.get("training_matrix"),[])
    ex_opls=_pj(stab.get("opls"),[])
    opl_titles=[o.get("title","") for o in ex_opls if o.get("title")]
    with st.form("fi_training_form"):
        new_tm=[]
        if team:
            for mi,m in enumerate(team):
                mn=m["member_name"]; em=next((x for x in ex_tm if x.get("member")==mn),{})
                tc1,tc2,tc3=st.columns(3)
                tc1.markdown(f"**{mn}**")
                trained=tc2.multiselect("Trained on",opl_titles,
                    default=[x for x in _pj(em.get("opls_trained"),[]) if x in opl_titles],
                    key=f"tr_{mi}")
                planned=tc3.checkbox("Training planned?",value=bool(em.get("plan_confirmed")),key=f"tp_{mi}")
                new_tm.append({"member":mn,"opls_trained":trained,"plan_confirmed":planned})
        else:
            st.info("Add team members first.")
        if st.form_submit_button("Save Training Matrix",type="primary") and can_edit:
            payload={"project_id":pid,"training_matrix":json.dumps(new_tm)}
            if stab.get("id"):
                supabase.table("fi_stabilisation").update(payload).eq("id",stab["id"]).execute()
            else:
                supabase.table("fi_stabilisation").insert(payload).execute()
            _mark_req(supabase,pid,35,len(new_tm)>0,name)
            st.success("Saved ✓"); st.rerun()

# Map form keys to their render functions
FORM_RENDERERS = {
    "team":        _form_team,
    "kpi_setup":   _form_kpi_setup,
    "master_plan": _form_master_plan,
    "meeting":     _form_meeting,
    "rca":         _form_rca,
    "actions":     _form_actions,
    "cil":         _form_cil,
    "five_s":      _form_five_s,
    "monitoring":  _form_monitoring,
    "opls":        _form_opls,
    "training":    _form_training,
}


# ─────────────────────────────────────────────────────────────────────────────
# MAIN RENDER
# ─────────────────────────────────────────────────────────────────────────────
def render_fi_projects_tab(supabase, role, pillar, name):
    can_edit   = (role=="plant_manager") or (role=="pillar_leader" and pillar=="FI")

    st.markdown(_CSS, unsafe_allow_html=True)

    # ── project selector ──────────────────────────────────────────────────────
    try:
        all_projects = supabase.table("fi_projects").select("*")\
            .order("created_at",desc=True).execute().data or []
    except Exception as e:
        st.error(f"Could not load projects: {e}"); return

    h1,h2,h3 = st.columns([4,1,1])
    selected_project = None

    if not all_projects:
        st.info("No FI projects yet.")
    else:
        proj_map={p["id"]:p["project_name"] for p in all_projects}
        sel_id=h1.selectbox("Project",list(proj_map.keys()),
            format_func=lambda x:proj_map[x],key="fi_sel",label_visibility="collapsed")
        selected_project=next((p for p in all_projects if p["id"]==sel_id),None)

    if can_edit and h2.button("＋ New",key="fi_new_btn"):
        st.session_state["fi_creating"]=True

    # delete (plant manager only)
    if role=="plant_manager" and selected_project:
        if h3.button("🗑 Delete",key="fi_del_btn"):
            st.session_state["fi_confirm_delete"]=selected_project["id"]
        if st.session_state.get("fi_confirm_delete")==selected_project["id"]:
            st.warning(f"⚠️ Delete **{selected_project['project_name']}** and all data?")
            dc1,dc2,_=st.columns([1,1,5])
            if dc1.button("Yes, delete",type="primary",key="fi_del_confirm"):
                for tbl in ["fi_weekly_updates","fi_project_steps","fi_project_team",
                            "fi_project_kpi","fi_project_cost","fi_actions",
                            "fi_audit_records","fi_stabilisation","fi_project_analysis",
                            "fi_project_checklist","fi_meetings"]:
                    try: supabase.table(tbl).delete().eq("project_id",sel_id).execute()
                    except: pass
                supabase.table("fi_projects").delete().eq("id",sel_id).execute()
                st.session_state.pop("fi_confirm_delete",None); st.success("Deleted"); st.rerun()
            if dc2.button("Cancel",key="fi_del_cancel"):
                st.session_state.pop("fi_confirm_delete",None); st.rerun()

    # create project
    if st.session_state.get("fi_creating") and can_edit:
        with st.expander("Create New Project",expanded=True):
            ns=st.selectbox("Section",list(PLANT_SECTIONS.keys()),key="fi_ns")
            nm=PLANT_SECTIONS.get(ns,[])
            area_new=f"{ns} — {st.selectbox('Machine',['— All —']+nm,key='fi_nm')}" if nm else ns
            with st.form("fi_create_form"):
                n_name   =st.text_input("Project Name *")
                n_problem=st.text_area("Problem Statement *",height=70)
                nc1,nc2  =st.columns(2)
                n_launch =nc1.date_input("Launch Date",value=date.today())
                n_end    =nc2.date_input("Expected Completion",value=date.today()+timedelta(weeks=12))
                n_kpi    =st.text_input("Company KPI this project supports")
                if st.form_submit_button("Create Project",type="primary"):
                    if n_name and n_problem:
                        supabase.table("fi_projects").insert({
                            "project_name":n_name,"problem_statement":n_problem,
                            "target_area":area_new,"launch_date":str(n_launch),
                            "expected_completion_date":str(n_end),
                            "company_kpi_link":n_kpi,"created_by":name,
                        }).execute()
                        st.session_state["fi_creating"]=False; st.rerun()
                    else: st.error("Name and Problem Statement required.")
        if st.button("Cancel",key="fi_cancel"): st.session_state["fi_creating"]=False; st.rerun()

    if not selected_project: return

    pid = selected_project["id"]
    cw  = _cw(selected_project.get("launch_date"))
    _sync_checklist_from_data(supabase, pid, name)
    checklist = _load_checklist(supabase, pid)
    total_score = _score(checklist)
    target_now  = TARGET_RAMP.get(cw,100)

    # ── PROJECT HEADER ────────────────────────────────────────────────────────
    pct = total_score / TOTAL_POINTS * 100
    on_track = total_score >= target_now
    st.markdown(f"""
    <div style="background:linear-gradient(135deg,#17375E,#0C5595);border-radius:14px;
         padding:20px 26px;margin-bottom:18px;color:white">
      <div style="font-size:21px;font-weight:800">{selected_project['project_name']}</div>
      <div style="font-size:12px;opacity:.70;margin-bottom:12px">
        {selected_project.get('target_area','')} &nbsp;·&nbsp;
        Week {cw} of 12 &nbsp;·&nbsp;
        Launch: {str(selected_project.get('launch_date',''))[:10]}
      </div>
      <div style="display:flex;align-items:center;gap:20px">
        <div style="font-size:38px;font-weight:900;line-height:1">{total_score}</div>
        <div style="flex:1">
          <div style="font-size:11px;opacity:.65;margin-bottom:4px;display:flex;justify-content:space-between">
            <span>Score vs target</span><span>{total_score} / {target_now} pts (W{cw} target)</span>
          </div>
          <div style="height:8px;border-radius:4px;background:rgba(255,255,255,.2);overflow:hidden">
            <div style="height:100%;border-radius:4px;width:{min(pct,100):.0f}%;
                 background:{'#4ade80' if on_track else '#fbbf24'}"></div>
          </div>
        </div>
        <div style="font-size:13px;font-weight:700;
             color:{'#4ade80' if on_track else '#fbbf24'}">
          {'✓ On track' if on_track else '⚠ Below target'}
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── WEEK NAVIGATION ───────────────────────────────────────────────────────
    if "fi_sel_week" not in st.session_state:
        st.session_state["fi_sel_week"] = min(cw, max(ACTIVE_WEEKS))

    sel_week = st.session_state["fi_sel_week"]

    # 12-column week pills
    wcols = st.columns(12)
    for wi, wcol in enumerate(wcols, start=1):
        has_r = wi in ACTIVE_WEEKS
        if not has_r:
            wcol.markdown(f"<div style='text-align:center;color:#CBD5E1;font-size:11px;padding:6px 0'>W{wi}</div>",
                          unsafe_allow_html=True)
            continue
        reqs_w = REQ_BY_WEEK[wi]
        done_w = sum(1 for r in reqs_w if checklist.get(r["id"],{}).get("done"))
        all_done = done_w==len(reqs_w)
        is_sel   = wi==sel_week

        if all_done:   bg,tc="#1E8449","white"
        elif wi==cw:   bg,tc="#0C5595","white"
        elif wi<cw:    bg,tc="#FEF2F2","#DE201B"
        else:          bg,tc="#F4F6F8","#566573"

        border = f"3px solid #D68910" if is_sel else "3px solid transparent"
        if wcol.button(f"W{wi}",key=f"wk_btn_{wi}",
                       help=f"W{wi}: {done_w}/{len(reqs_w)} done"):
            st.session_state["fi_sel_week"]=wi; st.rerun()

        # dots under button
        dots="".join(
            f"<span class='week-dot' style='background:{'#1E8449' if checklist.get(r['id'],{}).get('done') else '#E2E8F0'}'></span>"
            for r in reqs_w)
        wcol.markdown(f"<div style='text-align:center;margin-top:-2px'>{dots}</div>",unsafe_allow_html=True)

    st.divider()

    # ── WEEK REQUIREMENTS ─────────────────────────────────────────────────────
    reqs_w = REQ_BY_WEEK.get(sel_week,[])
    done_w = sum(1 for r in reqs_w if checklist.get(r["id"],{}).get("done"))
    is_past   = sel_week < cw
    is_future = sel_week > cw

    # Week heading
    wc1,wc2 = st.columns([3,1])
    sc = "#1E8449" if done_w==len(reqs_w) else ("#DE201B" if is_past else "#0C5595")
    wc1.markdown(f"""
    <div style="font-size:20px;font-weight:800;color:#0C5595">
      Week {sel_week}
      <span style="font-size:13px;font-weight:600;color:{sc};
            background:#F4F6F8;padding:3px 10px;border-radius:12px;margin-left:8px">
        {done_w}/{len(reqs_w)} done
      </span>
    </div>
    <div style="font-size:12px;color:#94A3B8;margin-top:2px">
      {sum(r['pts'] for r in reqs_w if checklist.get(r['id'],{}).get('done'))} /
      {sum(r['pts'] for r in reqs_w)} pts this week
      {'  ·  <b style="color:#DE201B">Past week — not fully complete</b>' if is_past and done_w<len(reqs_w) else ''}
    </div>
    """, unsafe_allow_html=True)

    # ── Track which form is open ──────────────────────────────────────────────
    open_form_key = st.session_state.get(f"fi_open_form_w{sel_week}")

    # ── Requirement cards ──────────────────────────────────────────────────────
    # Group by form so one form serves multiple requirements
    seen_forms = set()

    for r in reqs_w:
        is_done  = bool(checklist.get(r["id"],{}).get("done"))
        card_cls = "req-done" if is_done else ("req-future" if is_future else ("req-late" if is_past else "req-open"))
        icon = "✅" if is_done else ("🔒" if is_future else ("🔴" if is_past else "⬜"))
        fm   = FORM_META[r["form"]]

        r_text  = r["text"]
        r_pts   = r["pts"]
        r_pts_s = "pts" if r_pts > 1 else "pt"
        fm_col  = fm["color"]
        fm_icon = fm["icon"]
        fm_lbl  = fm["label"]
        st.markdown(f"""
        <div class="req-card {card_cls}">
          <div style="display:flex;align-items:flex-start;gap:10px">
            <span style="font-size:18px;flex-shrink:0">{icon}</span>
            <div style="flex:1">
              <div style="font-size:14px;font-weight:600;color:#1E293B">
                {r_text}
                <span class="pts-pill">{r_pts} {r_pts_s}</span>
              </div>
            </div>
            <div style="flex-shrink:0;text-align:right;font-size:11px;color:{fm_col};
                 font-weight:600;min-width:90px">
              {fm_icon} {fm_lbl}
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        # "Complete it" button — opens the form inline
        fk = r["form"]
        if fk not in seen_forms and can_edit:
            is_form_open = open_form_key == fk
            btn_label = f"▼ {fm['icon']} {fm['label']}" if is_form_open else f"► Complete it  — {fm['icon']} {fm['label']}"
            if st.button(btn_label, key=f"open_form_{sel_week}_{fk}"):
                st.session_state[f"fi_open_form_w{sel_week}"] = None if is_form_open else fk
                st.rerun()

            if is_form_open:
                with st.container():
                    st.markdown(f"""<div style="border-left:3px solid {fm['color']};
                        padding-left:16px;margin:4px 0 16px">""", unsafe_allow_html=True)
                    FORM_RENDERERS[fk](supabase, pid, selected_project, checklist, cw, name, can_edit)
                    st.markdown("</div>", unsafe_allow_html=True)
            seen_forms.add(fk)

        st.markdown("<div style='height:2px'></div>", unsafe_allow_html=True)

    # ── Score ramp chart ──────────────────────────────────────────────────────
    st.divider()
    st.markdown("#### Score Progress")
    fig = _ramp_fig(checklist, cw)
    st.pyplot(fig, use_container_width=True)
    plt.close(fig)

    # ── Edit project details ──────────────────────────────────────────────────
    if can_edit:
        with st.expander("⚙️ Edit Project Details"):
            with st.form("fi_edit_proj"):
                ep1,ep2=st.columns(2)
                e_name   =ep1.text_input("Project Name",value=selected_project.get("project_name",""))
                e_kpi    =ep2.text_input("Company KPI",value=selected_project.get("company_kpi_link",""))
                e_problem=st.text_area("Problem Statement",value=selected_project.get("problem_statement",""),height=70)
                ed1,ed2  =st.columns(2)
                e_launch =ed1.date_input("Launch Date",
                    value=date.fromisoformat(str(selected_project.get("launch_date",date.today()))[:10]))
                e_end    =ed2.date_input("Expected End",
                    value=date.fromisoformat(str(selected_project.get("expected_completion_date",date.today()+timedelta(weeks=12)))[:10]))
                if st.form_submit_button("Save Changes"):
                    supabase.table("fi_projects").update({
                        "project_name":e_name,"problem_statement":e_problem,
                        "company_kpi_link":e_kpi,"launch_date":str(e_launch),
                        "expected_completion_date":str(e_end),
                    }).eq("id",pid).execute()
                    st.success("Saved"); st.rerun()
