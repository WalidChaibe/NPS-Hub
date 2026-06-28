# fi_projects_tab.py  — v2  Week-by-Week Checklist Interface
import io, json, base64, math
from datetime import date, timedelta
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import streamlit as st

# ─────────────────────────────────────────────
# REQUIREMENTS  (36 audit criteria)
# ─────────────────────────────────────────────
REQUIREMENTS = [
    {"id":1,  "week":1,  "points":1,  "text":"Are team members listed?",
     "where":"Imp. Project Launch Report (2.1) · Team Members Sheet",
     "links":["Launch Report 2.1","Team Members"]},
    {"id":2,  "week":1,  "points":1,  "text":"Have all team members been assigned clear roles?",
     "where":"Imp. Project Launch Report (2.1) · Team Members Sheet · Who Does What Sheet",
     "links":["Launch Report 2.1","Team Members","Who Does What"]},
    {"id":3,  "week":1,  "points":4,  "text":"Is there a clear link to a company/area KPI?",
     "where":"Imp. Project Launch Report (5)",
     "links":["Launch Report 5"]},
    {"id":4,  "week":1,  "points":1,  "text":"Is the historical data (timeframe & value) clearly shown?",
     "where":"Imp. Project Launch Report (1.1, 1.2, 1.3)",
     "links":["Launch Report 1.1-1.3"]},
    {"id":5,  "week":1,  "points":1,  "text":"Is the performance indicator subdivided into definable components?",
     "where":"KPI Breakdown Sheet",
     "links":["KPI Breakdown"]},
    {"id":6,  "week":1,  "points":2,  "text":"Are Route and Master Plan clearly visible and up-to-date?",
     "where":"Team Master Plan Sheet",
     "links":["Master Plan"]},
    {"id":7,  "week":1,  "points":2,  "text":"Are meetings organized and attendance at expected levels?",
     "where":"Imp. Project Launch Report (1.4) · Meeting Attendance Sheet",
     "links":["Launch Report 1.4","Meeting Attendance"]},
    {"id":8,  "week":2,  "points":1,  "text":"Is the performance indicator's target (timeframe & value) clearly shown?",
     "where":"Imp. Project Launch Report (4, 5)",
     "links":["Launch Report 4 & 5"]},
    {"id":9,  "week":2,  "points":1,  "text":"Is the target of each step clear?",
     "where":"KPI Breakdown Sheet · Team Master Plan Sheet",
     "links":["KPI Breakdown","Master Plan"]},
    {"id":10, "week":2,  "points":1,  "text":"Is data collection consistent across all shifts? All details captured?",
     "where":"Data Collection Plan Sheet",
     "links":["Data Collection Plan"]},
    {"id":11, "week":3,  "points":1,  "text":"Have step targets been subdivided into specific activities?",
     "where":"Who Does What Sheet · Team Master Plan Sheet",
     "links":["Who Does What","Master Plan"]},
    {"id":12, "week":3,  "points":5,  "text":"Is the route/methodology well understood by all team members?",
     "where":"Imp. Project Launch Report (3)",
     "links":["Launch Report 3"]},
    {"id":13, "week":3,  "points":3,  "text":"Can a randomly picked team member explain the activity board?",
     "where":"Team Master Plan Sheet · Who Does What Sheet",
     "links":["Master Plan","Who Does What"]},
    {"id":14, "week":4,  "points":4,  "text":"Has a Cost/Benefit chart been introduced and kept up-to-date?",
     "where":"Financial KPI Tracking Sheet",
     "links":["Financial KPI"]},
    {"id":15, "week":4,  "points":1,  "text":"Have root-cause analyses been used and well documented?",
     "where":"Root Cause Analysis Sheet",
     "links":["Root Cause Analysis"]},
    {"id":16, "week":4,  "points":3,  "text":"Have the critical areas been restored to basic conditions?",
     "where":"Restoration & Basic Conditions Sheet",
     "links":["Restoration"]},
    {"id":17, "week":4,  "points":5,  "text":"Are planned actions visible with target completion dates?",
     "where":"Who Does What Sheet · Team Master Plan Sheet",
     "links":["Who Does What","Master Plan"]},
    {"id":18, "week":5,  "points":1,  "text":"Have problem causes been verified and quantified with data?",
     "where":"Root Cause Analysis Sheet",
     "links":["Root Cause Analysis"]},
    {"id":19, "week":5,  "points":2,  "text":"Has the team used the route methods/tools to attack problems?",
     "where":"Imp. Project Launch Report (3) · Root Cause Analysis Sheet",
     "links":["Launch Report 3","Root Cause Analysis"]},
    {"id":20, "week":5,  "points":2,  "text":"Is there an owner assigned to each action?",
     "where":"Who Does What Sheet",
     "links":["Who Does What"]},
    {"id":21, "week":5,  "points":2,  "text":"Is the action plan up-to-date?",
     "where":"Who Does What Sheet",
     "links":["Who Does What"]},
    {"id":22, "week":5,  "points":3,  "text":"Is the majority of actions completed on time?",
     "where":"Who Does What Sheet",
     "links":["Who Does What"]},
    {"id":23, "week":5,  "points":4,  "text":"Is there evidence of implemented actions (OPLs, pictures, standards)?",
     "where":"OPL & SOP Register Sheet",
     "links":["OPL & SOP"]},
    {"id":24, "week":5,  "points":3,  "text":"Are CIL standards clear? Do CIL audits achieve at least 90%?",
     "where":"CIL Standards & Audits Sheet",
     "links":["CIL Standards"]},
    {"id":25, "week":5,  "points":1,  "text":"Is the workplace well organized (5S)?",
     "where":"5S Audit Sheet",
     "links":["5S Audit"]},
    {"id":26, "week":5,  "points":1,  "text":"Are improvements on the targeted machine areas evident?",
     "where":"Restoration & Basic Conditions Sheet · OPL & SOP Register Sheet",
     "links":["Restoration","OPL & SOP"]},
    {"id":27, "week":6,  "points":2,  "text":"Has the group found logical countermeasures with sound logic?",
     "where":"Root Cause Analysis Sheet · Who Does What Sheet",
     "links":["Root Cause Analysis","Who Does What"]},
    {"id":28, "week":7,  "points":1,  "text":"Is the reoccurrence analysis present and updated?",
     "where":"Root Cause Analysis Sheet",
     "links":["Root Cause Analysis"]},
    {"id":29, "week":7,  "points":1,  "text":"Is the single problem analysis applied and followed up?",
     "where":"Root Cause Analysis Sheet",
     "links":["Root Cause Analysis"]},
    {"id":30, "week":7,  "points":15, "text":"Is the trend of the performance indicator positive?",
     "where":"KPI Count Tracking Sheet · Financial KPI Tracking Sheet",
     "links":["KPI Count","Financial KPI"]},
    {"id":31, "week":7,  "points":2,  "text":"Are monitoring systems for key actions in place and visible?",
     "where":"Monitoring & Controls Sheet",
     "links":["Monitoring"]},
    {"id":32, "week":8,  "points":2,  "text":"Are monitoring system devices used and up-to-date?",
     "where":"Monitoring & Controls Sheet",
     "links":["Monitoring"]},
    {"id":33, "week":10, "points":2,  "text":"Are there procedures in place to hold the gains achieved?",
     "where":"OPL & SOP Register Sheet · Monitoring & Controls Sheet",
     "links":["OPL & SOP","Monitoring"]},
    {"id":34, "week":10, "points":2,  "text":"Have OPLs/SOPs been created for every significant improvement?",
     "where":"OPL & SOP Register Sheet",
     "links":["OPL & SOP"]},
    {"id":35, "week":10, "points":2,  "text":"Is there a training matrix for OPLs/SOPs with a training plan?",
     "where":"Training Matrix Sheet",
     "links":["Training Matrix"]},
    {"id":36, "week":11, "points":15, "text":"Has the team achieved its goal or made substantial progress?",
     "where":"KPI Count Tracking Sheet · Financial KPI Tracking Sheet",
     "links":["KPI Count","Financial KPI"]},
]

# Group by week
from collections import defaultdict
REQ_BY_WEEK = defaultdict(list)
for r in REQUIREMENTS:
    REQ_BY_WEEK[r["week"]].append(r)

ACTIVE_WEEKS = sorted(REQ_BY_WEEK.keys())   # [1,2,3,4,5,6,7,8,10,11]
TOTAL_POINTS = sum(r["points"] for r in REQUIREMENTS)   # 100

TARGET_RAMP = {1:12,2:15,3:24,4:33,5:56,6:58,7:77,8:79,9:79,10:85,11:100,12:100}

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

# ─────────────────────────────────────────────
# COLOURS
# ─────────────────────────────────────────────
C = {
    "blue":   "#0C5595",
    "red":    "#DE201B",
    "green":  "#1E8449",
    "amber":  "#D68910",
    "grey":   "#566573",
    "lgrey":  "#F4F6F8",
    "mgrey":  "#BDC3C7",
    "white":  "#FFFFFF",
    "teal":   "#0E5E86",
    "navy":   "#17375E",
}

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
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

def _parse_json(v, fallback):
    if isinstance(v, list): return v
    if isinstance(v, str) and v.strip():
        try: return json.loads(v)
        except: pass
    return fallback


# ─────────────────────────────────────────────
# LOAD / SAVE CHECKLIST STATE
# ─────────────────────────────────────────────
def _load_checklist(supabase, project_id):
    """Returns dict {req_id: {"done": bool, "note": str, "updated_at": str}}"""
    try:
        rows = supabase.table("fi_project_checklist").select("*")\
            .eq("project_id", project_id).execute().data or []
        return {r["req_id"]: r for r in rows}
    except:
        return {}

def _save_req(supabase, project_id, req_id, done, note, user_name):
    try:
        supabase.table("fi_project_checklist").upsert({
            "project_id": project_id,
            "req_id":     req_id,
            "done":       done,
            "note":       note or "",
            "updated_by": user_name,
            "updated_at": date.today().isoformat(),
        }, on_conflict="project_id,req_id").execute()
        return True
    except Exception as e:
        st.error(f"Save failed: {e}"); return False


# ─────────────────────────────────────────────
# SCORE CHART
# ─────────────────────────────────────────────
def _score_chart(checklist, cur_week):
    # cumulative score up to each week
    cum = 0
    week_scores = {}
    for w in range(1, 13):
        for r in REQ_BY_WEEK.get(w, []):
            if checklist.get(r["id"], {}).get("done"):
                cum += r["points"]
        if w in ACTIVE_WEEKS or w <= cur_week:
            week_scores[w] = cum

    fig, ax = plt.subplots(figsize=(10, 3.2))
    fig.patch.set_facecolor("#ffffff"); ax.set_facecolor("#fafafa")

    tw = list(range(1, 13))
    tt = [TARGET_RAMP[w] for w in tw]
    ax.fill_between(tw, 0, tt, alpha=0.06, color="#566573")
    ax.plot(tw, tt, "s--", color="#BDC3C7", lw=1.5, ms=4, label="Target ramp", zorder=2)

    if week_scores:
        ws = sorted(week_scores.keys())
        wv = [week_scores[w] for w in ws]
        ax.fill_between(ws, 0, wv, alpha=0.12, color=C["blue"])
        ax.plot(ws, wv, "o-", color=C["blue"], lw=2.5, ms=7, label="Score", zorder=5)
        for w, v in zip(ws, wv):
            ax.annotate(f"{v}", (w,v), textcoords="offset points",
                        xytext=(0,8), fontsize=8, ha="center",
                        color=C["blue"], fontweight="bold")

    ax.axvline(cur_week, color=C["red"], lw=1.5, ls=":", alpha=0.7)
    ax.set_xlim(0.5, 12.5); ax.set_ylim(0, 110)
    ax.set_xticks(range(1,13))
    ax.set_xticklabels([f"W{i}" for i in range(1,13)], fontsize=9)
    ax.set_ylabel("Score / 100", fontsize=9, color=C["grey"])
    ax.legend(fontsize=9, frameon=False)
    ax.grid(axis="y", color="#eeeeee", lw=0.6)
    for sp in ["top","right"]: ax.spines[sp].set_visible(False)
    fig.tight_layout(pad=0.5)
    return fig


# ─────────────────────────────────────────────
# MAIN RENDER
# ─────────────────────────────────────────────
def render_fi_projects_tab(supabase, role, pillar, name):
    can_edit   = (role == "plant_manager") or (role == "pillar_leader" and pillar == "FI")
    is_auditor = role in ["plant_manager","pillar_leader"]

    # ── CSS ──────────────────────────────────────────────────────────────────
    st.markdown("""
    <style>
    .week-pill {
        display:inline-block; padding:6px 14px; border-radius:20px;
        font-size:13px; font-weight:700; cursor:pointer;
        margin:3px; border:none; text-align:center; min-width:52px;
    }
    .req-card {
        border-radius:10px; padding:14px 16px; margin-bottom:8px;
        border-left:4px solid; background:#ffffff;
        box-shadow:0 1px 4px rgba(0,0,0,0.06);
    }
    .req-done  { border-left-color:#1E8449; background:#F0FFF4; }
    .req-open  { border-left-color:#0C5595; background:#FFFFFF; }
    .req-late  { border-left-color:#DE201B; background:#FFF5F5; }
    .pts-badge {
        display:inline-block; padding:2px 8px; border-radius:12px;
        font-size:11px; font-weight:700; background:#EBF8FF; color:#0C5595;
        margin-left:8px;
    }
    .link-btn {
        display:inline-block; padding:3px 10px; border-radius:10px;
        font-size:11px; font-weight:600; background:#F4F6F8;
        color:#0C5595; margin:2px; text-decoration:none; border:1px solid #d0d9e8;
    }
    .week-header {
        font-size:22px; font-weight:800; color:#0C5595; margin-bottom:2px;
    }
    .week-sub {
        font-size:13px; color:#94A3B8; margin-bottom:16px;
    }
    .progress-bar-bg {
        height:8px; border-radius:4px; background:#E2E8F0; overflow:hidden;
    }
    .progress-bar-fill {
        height:100%; border-radius:4px;
        background:linear-gradient(90deg,#0C5595,#1E8449);
    }
    </style>
    """, unsafe_allow_html=True)

    # ── Project selector ──────────────────────────────────────────────────────
    try:
        all_projects = supabase.table("fi_projects").select("*")\
            .order("created_at", desc=True).execute().data or []
    except Exception as e:
        st.error(f"Could not load projects: {e}"); return

    top1, top2, top3 = st.columns([4,1,1])

    if not all_projects:
        st.info("No FI projects yet."); selected_project = None
    else:
        proj_map = {p["id"]: p["project_name"] for p in all_projects}
        sel_id = top1.selectbox("Project", list(proj_map.keys()),
            format_func=lambda x: proj_map[x], key="fi_sel",
            label_visibility="collapsed")
        selected_project = next((p for p in all_projects if p["id"]==sel_id), None)

    if can_edit and top2.button("＋ New Project", key="fi_new_btn"):
        st.session_state["fi_creating"] = True

    # Plant manager: delete button
    if role=="plant_manager" and selected_project:
        if top3.button("🗑 Delete", key="fi_del_btn"):
            st.session_state["fi_confirm_delete"] = selected_project["id"]
        if st.session_state.get("fi_confirm_delete") == selected_project["id"]:
            st.warning(f"⚠️ Delete **{selected_project['project_name']}** and all its data? This cannot be undone.")
            c1,c2,_ = st.columns([1,1,5])
            if c1.button("Yes, delete", type="primary", key="fi_del_confirm"):
                for tbl in ["fi_weekly_updates","fi_project_steps","fi_project_team",
                            "fi_project_kpi","fi_project_cost","fi_actions",
                            "fi_audit_records","fi_stabilisation","fi_company_kpis",
                            "fi_project_analysis","fi_project_checklist"]:
                    try: supabase.table(tbl).delete().eq("project_id",sel_id).execute()
                    except: pass
                supabase.table("fi_projects").delete().eq("id",sel_id).execute()
                st.session_state.pop("fi_confirm_delete",None)
                st.success("Deleted."); st.rerun()
            if c2.button("Cancel", key="fi_del_cancel"):
                st.session_state.pop("fi_confirm_delete",None); st.rerun()

    # ── Create project ────────────────────────────────────────────────────────
    if st.session_state.get("fi_creating") and can_edit:
        with st.expander("Create New Project", expanded=True):
            ns = st.selectbox("Section", list(PLANT_SECTIONS.keys()), key="fi_ns")
            nm = PLANT_SECTIONS.get(ns,[])
            area_new = f"{ns} — {st.selectbox('Machine',['— All —']+nm,key='fi_nm')}" if nm else ns
            with st.form("fi_create_form"):
                n_name    = st.text_input("Project Name *")
                n_problem = st.text_area("Problem Statement *", height=70)
                c1,c2 = st.columns(2)
                n_launch = c1.date_input("Launch Date", value=date.today())
                n_end    = c2.date_input("Expected Completion", value=date.today()+timedelta(weeks=12))
                n_kpi    = st.text_input("Company KPI this project supports")
                if st.form_submit_button("Create Project", type="primary"):
                    if n_name and n_problem:
                        supabase.table("fi_projects").insert({
                            "project_name":n_name,"problem_statement":n_problem,
                            "target_area":area_new,"launch_date":str(n_launch),
                            "expected_completion_date":str(n_end),
                            "company_kpi_link":n_kpi,"created_by":name,
                        }).execute()
                        st.session_state["fi_creating"]=False; st.rerun()
                    else: st.error("Name and Problem Statement required.")
        if st.button("Cancel", key="fi_cancel_create"):
            st.session_state["fi_creating"]=False; st.rerun()

    if not selected_project: return

    pid = selected_project["id"]
    cw  = _cw(selected_project.get("launch_date"))

    # ── Load checklist state ──────────────────────────────────────────────────
    checklist = _load_checklist(supabase, pid)

    # ── Compute scores ────────────────────────────────────────────────────────
    total_done  = sum(r["points"] for r in REQUIREMENTS if checklist.get(r["id"],{}).get("done"))
    total_pts   = TOTAL_POINTS   # 100
    target_now  = TARGET_RAMP.get(cw, 100)
    pct_overall = total_done / total_pts * 100

    # Per-week stats
    week_stats = {}
    for w in ACTIVE_WEEKS:
        reqs = REQ_BY_WEEK[w]
        done = [r for r in reqs if checklist.get(r["id"],{}).get("done")]
        week_stats[w] = {
            "total": len(reqs),
            "done":  len(done),
            "pts_done": sum(r["points"] for r in done),
            "pts_total": sum(r["points"] for r in reqs),
        }

    # ── TOP HEADER ────────────────────────────────────────────────────────────
    st.markdown(f"""
    <div style="background:linear-gradient(135deg,{C['navy']},{C['blue']});
         border-radius:14px;padding:20px 24px;margin-bottom:20px;color:white;">
      <div style="font-size:22px;font-weight:800;margin-bottom:4px">
        {selected_project['project_name']}
      </div>
      <div style="font-size:13px;opacity:0.75;margin-bottom:14px">
        {selected_project.get('target_area','')} &nbsp;·&nbsp;
        Launch: {str(selected_project.get('launch_date',''))[:10]} &nbsp;·&nbsp;
        Currently Week {cw} of 12
      </div>
      <div style="display:flex;gap:20px;align-items:center">
        <div>
          <div style="font-size:36px;font-weight:900;line-height:1">{int(total_done)}</div>
          <div style="font-size:11px;opacity:0.65">pts achieved</div>
        </div>
        <div style="flex:1">
          <div style="display:flex;justify-content:space-between;font-size:11px;opacity:0.75;margin-bottom:4px">
            <span>Progress</span><span>{pct_overall:.0f}% (target: {target_now})</span>
          </div>
          <div class="progress-bar-bg">
            <div class="progress-bar-fill" style="width:{min(pct_overall,100):.0f}%"></div>
          </div>
        </div>
        <div style="text-align:right">
          <div style="font-size:28px;font-weight:900;line-height:1;
               color:{'#4ade80' if total_done>=target_now else '#fbbf24' if total_done>=target_now*0.7 else '#f87171'}">
            {'✓' if total_done>=target_now else '!'}</div>
          <div style="font-size:11px;opacity:0.65">
            {'On track' if total_done>=target_now else 'Below target'}</div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── WEEK NAVIGATION ───────────────────────────────────────────────────────
    # Show weeks 1–12 as pills; weeks with requirements are clickable
    sel_week = st.session_state.get("fi_sel_week", min(cw, max(ACTIVE_WEEKS)))

    # Build the week picker as columns
    cols = st.columns(12)
    for wi, col in enumerate(cols, start=1):
        wn = wi
        has_reqs = wn in ACTIVE_WEEKS
        if not has_reqs:
            col.markdown(f"""<div style="text-align:center;padding:6px 0;font-size:12px;
                color:#CBD5E1;font-weight:500">W{wn}</div>""", unsafe_allow_html=True)
            continue
        ws = week_stats[wn]
        all_done = ws["done"] == ws["total"]
        is_cur   = wn == cw
        is_sel   = wn == sel_week

        if all_done:
            bg, tc = C["green"], "white"
        elif is_cur:
            bg, tc = C["blue"], "white"
        elif wn < cw:
            bg, tc = "#FEF2F2", C["red"]  # past week not complete
        else:
            bg, tc = "#F4F6F8", C["grey"]  # future

        border = f"3px solid {C['amber']}" if is_sel else "3px solid transparent"

        if col.button(f"W{wn}", key=f"wk_{wn}",
                      help=f"Week {wn}: {ws['done']}/{ws['total']} done"):
            st.session_state["fi_sel_week"] = wn
            st.rerun()

        # mini progress dots below button
        dots = "".join(
            f"<span style='display:inline-block;width:5px;height:5px;border-radius:50%;"
            f"background:{'#1E8449' if checklist.get(r['id'],{}).get('done') else '#E2E8F0'};"
            f"margin:0 1px'></span>"
            for r in REQ_BY_WEEK[wn]
        )
        col.markdown(f"<div style='text-align:center;margin-top:-4px'>{dots}</div>",
                     unsafe_allow_html=True)

    st.divider()

    # ── SELECTED WEEK CONTENT ─────────────────────────────────────────────────
    reqs_this_week = REQ_BY_WEEK.get(sel_week, [])
    ws = week_stats.get(sel_week, {"done":0,"total":0,"pts_done":0,"pts_total":0})
    is_past   = sel_week < cw
    is_future = sel_week > cw

    # Week heading
    status_txt = "✓ Complete" if ws["done"]==ws["total"] else \
                 f"{ws['done']}/{ws['total']} done"
    status_col = C["green"] if ws["done"]==ws["total"] else \
                 (C["red"] if is_past else C["blue"])

    lc, rc = st.columns([3,1])
    lc.markdown(f"""
    <div class="week-header">Week {sel_week}
      <span style="font-size:14px;font-weight:600;color:{status_col};
            background:{'#F0FFF4' if ws['done']==ws['total'] else '#F4F6F8'};
            padding:3px 10px;border-radius:12px;margin-left:10px">{status_txt}</span>
    </div>
    <div class="week-sub">{ws['pts_done']} / {ws['pts_total']} pts this week
      {'&nbsp;·&nbsp; <b style=\"color:#DE201B\">Past — not all requirements met</b>' if is_past and ws['done']<ws['total'] else ''}
      {'&nbsp;·&nbsp; <i>Future week — prepare ahead</i>' if is_future else ''}
    </div>
    """, unsafe_allow_html=True)

    # Quick "mark all done" for auditors
    if can_edit and reqs_this_week:
        all_done_now = all(checklist.get(r["id"],{}).get("done") for r in reqs_this_week)
        if rc.button(
            "✓ Mark all done" if not all_done_now else "✗ Mark all open",
            key=f"bulk_{sel_week}", type="primary" if not all_done_now else "secondary"
        ):
            for r in reqs_this_week:
                _save_req(supabase, pid, r["id"], not all_done_now, "", name)
            st.rerun()

    # ── Requirement cards ──────────────────────────────────────────────────────
    for r in reqs_this_week:
        state    = checklist.get(r["id"], {})
        is_done  = bool(state.get("done"))
        note_val = state.get("note","")

        card_class = "req-done" if is_done else \
                     ("req-late" if is_past else "req-open")
        icon = "✅" if is_done else ("🔴" if is_past else "⬜")

        # Links as inline pills
        link_html = " ".join(
            f'<span class="link-btn">📋 {lnk}</span>' for lnk in r["links"]
        )

        st.markdown(f"""
        <div class="req-card {card_class}">
          <div style="display:flex;align-items:flex-start;gap:10px">
            <span style="font-size:18px;margin-top:1px">{icon}</span>
            <div style="flex:1">
              <div style="font-size:14px;font-weight:600;color:#1E293B;margin-bottom:4px">
                {r['text']}
                <span class="pts-badge">{r['points']} pt{'s' if r['points']>1 else ''}</span>
              </div>
              <div style="font-size:11px;color:#94A3B8;margin-bottom:6px">
                📁 {r['where']}
              </div>
              <div>{link_html}</div>
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        # Toggle + note (inline, tight)
        if can_edit:
            tc1, tc2 = st.columns([1, 4])
            new_done = tc1.checkbox(
                "Done" if not is_done else "Undo",
                value=is_done,
                key=f"chk_{r['id']}"
            )
            new_note = tc2.text_input(
                "Note / evidence",
                value=note_val,
                key=f"note_{r['id']}",
                placeholder="Optional note or evidence link…",
                label_visibility="collapsed"
            )
            if new_done != is_done or new_note != note_val:
                _save_req(supabase, pid, r["id"], new_done, new_note, name)
                st.rerun()

            if note_val and not is_done:
                pass  # already shown in text_input
            if is_done and state.get("updated_at"):
                st.caption(f"✓ Marked done · {state.get('updated_by','')} · {state.get('updated_at','')[:10]}")
        else:
            # Read-only view
            if is_done:
                st.caption(f"✓ {state.get('updated_by','')} · {state.get('updated_at','')[:10]}"
                           + (f" — {note_val}" if note_val else ""))

        st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

    # ── Score chart at bottom ──────────────────────────────────────────────────
    st.divider()
    st.markdown("#### Score Progress")
    fig = _score_chart(checklist, cw)
    st.pyplot(fig, use_container_width=True)
    plt.close(fig)

    # ── Edit project details (collapsed) ──────────────────────────────────────
    if can_edit:
        with st.expander("⚙️ Edit Project Details"):
            with st.form("fi_edit_project"):
                ea1,ea2 = st.columns(2)
                e_name    = ea1.text_input("Project Name", value=selected_project.get("project_name",""))
                e_kpi     = ea2.text_input("Company KPI", value=selected_project.get("company_kpi_link",""))
                e_problem = st.text_area("Problem Statement", value=selected_project.get("problem_statement",""), height=70)
                ed1,ed2   = st.columns(2)
                e_launch  = ed1.date_input("Launch Date",
                    value=date.fromisoformat(str(selected_project.get("launch_date",date.today()))[:10]))
                e_end     = ed2.date_input("Expected End",
                    value=date.fromisoformat(str(selected_project.get("expected_completion_date",date.today()+timedelta(weeks=12)))[:10]))
                if st.form_submit_button("Save Changes"):
                    supabase.table("fi_projects").update({
                        "project_name":e_name,"problem_statement":e_problem,
                        "company_kpi_link":e_kpi,"launch_date":str(e_launch),
                        "expected_completion_date":str(e_end),
                    }).eq("id",pid).execute()
                    st.success("Saved"); st.rerun()
