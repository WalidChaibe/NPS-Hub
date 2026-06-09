# fi_projects_tab.py
# Focused Improvement – full redesign with analysis saving + PDF report
import io, json, base64, math, tempfile
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

# ─────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────
QUESTIONS = {
    1:  {"text": "Team members listed",                              "score": 1,  "week": 1,  "dim": "Involvement"},
    2:  {"text": "All members assigned clear roles",                 "score": 1,  "week": 1,  "dim": "Involvement"},
    3:  {"text": "Problem linked to company KPI",                    "score": 4,  "week": 1,  "dim": "Results"},
    4:  {"text": "Cost/Benefit chart introduced & up-to-date",       "score": 4,  "week": 5,  "dim": "Results"},
    5:  {"text": "Historical data (timeframe & value) shown",        "score": 1,  "week": 1,  "dim": "Results"},
    6:  {"text": "Performance indicator target clearly shown",       "score": 1,  "week": 2,  "dim": "Results"},
    7:  {"text": "KPI subdivided into components",                   "score": 1,  "week": 1,  "dim": "Results"},
    8:  {"text": "Route & Master Plan visible and up-to-date",       "score": 2,  "week": 1,  "dim": "Method"},
    9:  {"text": "Target of each step clear",                        "score": 1,  "week": 2,  "dim": "Method"},
    10: {"text": "Step targets subdivided into activities",          "score": 1,  "week": 3,  "dim": "Method"},
    11: {"text": "Root-cause analysis documented",                   "score": 1,  "week": 4,  "dim": "Method"},
    12: {"text": "Causes verified and quantified with data",         "score": 1,  "week": 5,  "dim": "Method"},
    13: {"text": "Data collection consistent across all shifts",     "score": 1,  "week": 2,  "dim": "Method"},
    14: {"text": "Route methods/tools used to attack problems",      "score": 2,  "week": 5,  "dim": "Method"},
    15: {"text": "Reoccurrence analysis present & updated",          "score": 1,  "week": 7,  "dim": "Method"},
    16: {"text": "Single problem analysis applied & followed up",    "score": 1,  "week": 7,  "dim": "Method"},
    17: {"text": "Logical countermeasures defined",                  "score": 2,  "week": 6,  "dim": "Action Plan"},
    18: {"text": "Critical areas restored to basic conditions",      "score": 3,  "week": 4,  "dim": "Action Plan"},
    19: {"text": "Actions visible with target dates",                "score": 5,  "week": 4,  "dim": "Action Plan"},
    20: {"text": "Owner assigned to each action",                    "score": 2,  "week": 5,  "dim": "Action Plan"},
    21: {"text": "Action plan up-to-date",                           "score": 2,  "week": 5,  "dim": "Action Plan"},
    22: {"text": "Majority of actions completed on time",            "score": 3,  "week": 5,  "dim": "Action Plan"},
    23: {"text": "Evidence of implemented actions",                  "score": 4,  "week": 5,  "dim": "Action Plan"},
    24: {"text": "KPI trend is positive",                            "score": 15, "week": 7,  "dim": "Results"},
    25: {"text": "Goal achieved or substantial progress made",       "score": 15, "week": 11, "dim": "Results"},
    26: {"text": "Procedures in place to hold gains",                "score": 2,  "week": 10, "dim": "Stabilisation"},
    27: {"text": "Monitoring systems in place and visible",          "score": 2,  "week": 7,  "dim": "Stabilisation"},
    28: {"text": "Monitoring devices used and up-to-date",           "score": 2,  "week": 8,  "dim": "Stabilisation"},
    29: {"text": "OPLs/SOPs created for improvements",              "score": 2,  "week": 10, "dim": "Stabilisation"},
    30: {"text": "Training matrix for OPLs/SOPs in place",          "score": 2,  "week": 10, "dim": "Stabilisation"},
    31: {"text": "CIL standards clear, audits >= 90%",              "score": 3,  "week": 5,  "dim": "Stabilisation"},
    32: {"text": "Workplace well organised (5S)",                    "score": 1,  "week": 5,  "dim": "Stabilisation"},
    33: {"text": "Improvements on machine/area evident",             "score": 1,  "week": 5,  "dim": "Stabilisation"},
    34: {"text": "Methodology understood by all team members",       "score": 5,  "week": 3,  "dim": "Involvement"},
    35: {"text": "Random member can explain the activity board",     "score": 3,  "week": 3,  "dim": "Involvement"},
    36: {"text": "Meetings organised, attendance at expected level", "score": 2,  "week": 1,  "dim": "Involvement"},
}

TARGET_RAMP = {1:12,2:15,3:24,4:33,5:56,6:58,7:77,8:79,9:79,10:85,11:100,12:100}

TEAM_ROLES       = ["Team Leader","Analyst","Operator","Maintenance","Quality","Other"]
RCA_METHODS      = ["5-Why","Fishbone","Pareto","FMEA","Other"]
ACTION_STATUSES  = ["Open","In Progress","Completed","Overdue"]
MONITORING_TYPES = ["Checklist","Audit","Form","Visual Board","Other"]

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

KPI_TREE = {
    "OEE Improvement":         {"unit": "%",     "kais": ["Reduce Breakdown Time","Reduce Minor Stoppages","Reduce Setup & Changeover Time","Reduce Speed Losses","Increase Availability Rate","Increase Performance Rate","Increase Quality Rate"]},
    "Quality Defect Reduction":{"unit": "%",     "kais": ["Reduce Customer Complaints","Reduce Internal Defects (PPM)","Reduce Rework Rate","Reduce Scrap Rate","Improve First Pass Yield","Reduce NCR Count"]},
    "Waste Reduction":         {"unit": "%",     "kais": ["Reduce Paper/Board Waste %","Reduce Trim Waste","Reduce Ink & Chemical Waste","Reduce Energy Consumption","Reduce Sheet Waste"]},
    "Cost Reduction":          {"unit": "K SAR", "kais": ["Reduce Maintenance Spend","Reduce Material Costs","Reduce Labour Overtime","Reduce Energy Costs","Reduce Rework & Scrap Costs"]},
    "Safety Improvement":      {"unit": "Count", "kais": ["Reduce Near Miss Incidents","Reduce Lost Time Accidents","Improve Safety Audit Score","Eliminate Unsafe Conditions"]},
    "Delivery Performance":    {"unit": "%",     "kais": ["Reduce Order Lead Time","Improve Schedule Adherence","Improve OTIF Rate","Reduce Order Backlog"]},
    "5S Score Improvement":    {"unit": "Score", "kais": ["Improve Sort Score","Improve Set-in-Order Score","Improve Shine Score","Improve Standardise Score","Improve Sustain Score"]},
    "Throughput/Productivity": {"unit": "MT",    "kais": ["Increase Net Run Time","Increase Average Speed","Reduce Idle Time","Improve Capacity Utilisation"]},
}

UNITS = ["%","MT","SAR","K SAR","LM","SQM","Hits","Hits/Hour","LM/Min","Hours","Mins","Count","Score"]

C = {
    "blue":  "#0C5595", "red":   "#DE201B", "green": "#1E8449",
    "amber": "#D68910", "grey":  "#566573", "lgrey": "#F4F6F8",
    "mgrey": "#BDC3C7", "white": "#FFFFFF", "black": "#1A1A2E",
    "navy":  "#17375E", "teal":  "#0E5E86",
}


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def _cw(launch_date):
    if not launch_date:
        return 1
    if isinstance(launch_date, str):
        try: launch_date = date.fromisoformat(launch_date)
        except: return 1
    delta = (date.today() - launch_date).days
    return max(1, min(12, math.ceil(delta / 7) if delta > 0 else 1))

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

def _card(title, value, sub="", color="#0C5595", width=None):
    w = f"width:{width};" if width else "flex:1;"
    st.markdown(f"""
    <div style="{w}background:{color};border-radius:10px;padding:14px 18px;margin:4px;">
      <div style="color:rgba(255,255,255,.7);font-size:11px;font-weight:600;letter-spacing:.5px">{title.upper()}</div>
      <div style="color:#fff;font-size:22px;font-weight:700;margin:4px 0 2px">{value}</div>
      <div style="color:rgba(255,255,255,.65);font-size:11px">{sub}</div>
    </div>""", unsafe_allow_html=True)

def _fig_to_bytes(fig, dpi=150):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight")
    buf.seek(0)
    return buf

# ─────────────────────────────────────────────
# PDF SLIDE HELPERS  (matches QM pillar exactly)
# ─────────────────────────────────────────────
def _pdf_draw_cover(c, W, H, title, subtitle, date_str, logo_reader=None):
    c.setFillColorRGB(1,1,1); c.rect(0,0,W,H,fill=1,stroke=0)
    if logo_reader:
        try: c.drawImage(logo_reader, 18, H-115, width=280, height=110, preserveAspectRatio=True, mask="auto")
        except: pass
    line_y = H - 125
    c.setFillColor(HexColor("#0C5595")); c.rect(18, line_y, W-18, 4, fill=1, stroke=0)
    c.setFillColor(HexColor("#DE201B")); c.rect(0, line_y, 4, H-line_y, fill=1, stroke=0)
    c.setFillColor(HexColor("#0E5E86")); c.setFont("Helvetica-Bold", 38)
    tw = c.stringWidth(title,"Helvetica-Bold",38)
    c.drawString((W-tw)/2, H*0.46, title)
    c.setFillColor(HexColor("#DE201B")); c.rect(80, H*0.43, W-160, 2, fill=1, stroke=0)
    c.setFillColor(HexColor("#555555")); c.setFont("Helvetica-Oblique", 18)
    sw = c.stringWidth(subtitle,"Helvetica-Oblique",18)
    c.drawString((W-sw)/2, H*0.38, subtitle)
    c.setFillColor(HexColor("#888888")); c.setFont("Helvetica-Oblique", 13)
    dw = c.stringWidth(date_str,"Helvetica-Oblique",13)
    c.drawString(W-dw-40, 28, date_str)

def _pdf_draw_section(c, W, H, text):
    c.setFillColorRGB(1,1,1); c.rect(0,0,W,H,fill=1,stroke=0)
    c.setFillColor(HexColor("#DE201B")); c.rect(40, H*0.58, 110, 4, fill=1, stroke=0)
    c.setFillColor(HexColor("#0C5595")); c.rect(155, H*0.58, W-195, 4, fill=1, stroke=0)
    c.setFillColor(HexColor("#0E5E86")); c.setFont("Helvetica-Bold", 38)
    tw = c.stringWidth(text,"Helvetica-Bold",38)
    c.drawString((W-tw)/2, H*0.58-60, text)

def _pdf_title_bar(c, title, W, H):
    c.setFillColor(HexColor("#0E5E86")); c.setFont("Helvetica-Bold", 20)
    c.drawString(40, H-50, title)
    c.setFillColor(HexColor("#DE201B")); c.rect(40, H-66, 110, 4, fill=1, stroke=0)
    c.setFillColor(HexColor("#0C5595")); c.rect(155, H-66, W-195, 4, fill=1, stroke=0)

def _build_fi_pdf(slides, logo_reader=None):
    W, H = 960, 540
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(W, H))
    for s in slides:
        if s.get("cover"):
            _pdf_draw_cover(c, W, H, s["title"], s.get("subtitle",""), s.get("date",""), logo_reader)
        elif s.get("section"):
            _pdf_draw_section(c, W, H, s["title"])
        else:
            _pdf_title_bar(c, s["title"], W, H)
            fig = s.get("fig")
            if fig:
                img_buf = _fig_to_bytes(fig, dpi=150)
                img = ImageReader(img_buf)
                c.drawImage(img, 40, 40, width=W-80, height=H-110, preserveAspectRatio=True, anchor="c")
                plt.close(fig)
            elif s.get("text_lines"):
                c.setFont("Helvetica", 13)
                y = H - 90
                for line in s["text_lines"]:
                    if y < 50: break
                    c.setFillColor(HexColor("#1A1A2E"))
                    c.drawString(50, y, str(line)[:110])
                    y -= 22
        c.showPage()
    c.save(); buf.seek(0)
    return buf


# ─────────────────────────────────────────────
# SCORING ENGINE
# ─────────────────────────────────────────────
def _compute_scores(project, team, kpi, steps, wu_rows, actions, stab, audit_records):
    team=team or []; steps=steps or []; actions=actions or []
    wu_rows=wu_rows or []; stab=stab or {}; audit_records=audit_records or []; kpi=kpi or {}
    kpi_vals=[w["kpi_value"] for w in sorted(wu_rows,key=lambda x:x["week_number"]) if w.get("kpi_value") is not None]

    def _parse_sp(wu): return _parse_json(wu.get("step_progress"),[])

    def _met(qn):
        if qn==1: return len(team)>=1
        if qn==2: return bool(team) and all(m.get("role") for m in team)
        if qn==3: return bool(project.get("problem_statement")) and bool(project.get("company_kpi_link"))
        if qn==4: return bool(project.get("cost_benefit_done"))
        if qn==5: return bool(kpi.get("kpi_name")) and kpi.get("baseline_value") is not None
        if qn==6: return kpi.get("target_value") is not None and bool(kpi.get("target_date"))
        if qn==7: subs=_parse_json(kpi.get("sub_components"),[]); return len([s for s in subs if isinstance(s,dict)])>0
        if qn==8: return len(steps)>=2 and all(s.get("planned_start_week") and s.get("planned_end_week") for s in steps)
        if qn==9: return bool(steps) and all(s.get("planned_end_week") for s in steps)
        if qn==10: return any(_parse_sp(wu) for wu in wu_rows)
        if qn==11: return any(wu.get("rca_performed") and wu.get("rca_findings") for wu in wu_rows)
        if qn==12: return any(wu.get("causes_verified") for wu in wu_rows)
        if qn==13: return any(wu.get("shifts_covered") and "All" in (wu.get("shifts_covered") or "") for wu in wu_rows)
        if qn==14: return any(wu.get("rca_method") in {"5-Why","Fishbone","Pareto","FMEA"} for wu in wu_rows)
        if qn==15: return any(wu.get("reoccurrence_description") for wu in wu_rows)
        if qn==16: return any(wu.get("single_problem_analysis") and wu.get("single_problem_notes") for wu in wu_rows)
        if qn==17: return bool(actions) and all(a.get("root_cause_addressed") for a in actions)
        if qn==18: return any(wu.get("basic_conditions_restored") for wu in wu_rows)
        if qn==19: return bool(actions) and any(a.get("target_date") for a in actions)
        if qn==20: return bool(actions) and all(a.get("owner") for a in actions)
        if qn==21:
            cur=_cw(project.get("launch_date"))
            return any(a.get("created_week") and int(a.get("created_week",0))>=cur-1 for a in actions)
        if qn==22:
            past_due=[a for a in actions if _safe_date(a.get("target_date")) and _safe_date(a["target_date"])<date.today()]
            on_time=[a for a in past_due if a.get("status")=="Completed"]
            return bool(past_due) and len(on_time)/len(past_due)>=0.5
        if qn==23: return any(a.get("evidence_b64") for a in actions if a.get("status")=="Completed")
        if qn==24:
            if len(kpi_vals)<3: return False
            b=kpi.get("baseline_value",kpi_vals[0]); t=kpi.get("target_value",b)
            return (t>b and kpi_vals[-1]>kpi_vals[-3]) or (t<b and kpi_vals[-1]<kpi_vals[-3])
        if qn==25:
            if not kpi_vals or kpi.get("target_value") is None or kpi.get("baseline_value") is None: return False
            b=float(kpi["baseline_value"]); t=float(kpi["target_value"]); gap=abs(t-b)
            return gap>0 and abs(kpi_vals[-1]-b)/gap>=0.80
        if qn==26: return bool(stab.get("procedures_created")) and bool(_parse_json(stab.get("procedures"),[]))
        if qn==27: return bool(stab.get("monitoring_in_place"))
        if qn==28: return bool(stab.get("monitoring_active")) and bool(stab.get("monitoring_last_update"))
        if qn==29: return bool(_parse_json(stab.get("opls"),[]))
        if qn==30: return bool(_parse_json(stab.get("training_matrix"),[]))
        if qn==31: return bool(stab.get("cil_standards_defined")) and float(stab.get("cil_audit_score") or 0)>=90
        if qn==32: return int(stab.get("five_s_rating") or 0)>=3
        if qn==33: return bool(stab.get("improvements_visible"))
        if qn==34: return any((ar.get("question_scores") or {}).get("34") for ar in audit_records)
        if qn==35: return any((ar.get("question_scores") or {}).get("35") for ar in audit_records)
        if qn==36: return any(wu.get("meeting_held") for wu in wu_rows)
        return False

    cur_week=_cw(project.get("launch_date"))
    q_status={}
    for qn,qdata in QUESTIONS.items():
        met=_met(qn); due=cur_week>=qdata["week"]
        q_status[qn]={"met":met,"due":due,"score":qdata["score"] if met else 0,
                      "week_due":qdata["week"],"text":qdata["text"],"dim":qdata["dim"],"max":qdata["score"]}
    total=sum(v["score"] for v in q_status.values())
    ar_by_week={ar["week_number"]:ar for ar in audit_records}
    weekly_scores={}
    for w in range(1,13):
        if w in ar_by_week: weekly_scores[w]=float(ar_by_week[w].get("total_score",0))
        elif w==cur_week: weekly_scores[w]=total
    return q_status, weekly_scores, total


# ─────────────────────────────────────────────
# CHARTS
# ─────────────────────────────────────────────
def _gantt(steps, wu_rows, cur_week):
    if not steps: return None
    n=len(steps); fig_h=max(3.5,n*0.52+1.2)
    fig,ax=plt.subplots(figsize=(13,fig_h)); fig.patch.set_facecolor("#ffffff"); ax.set_facecolor("#fafafa")
    sp_map={}
    for wu in wu_rows:
        for sp in _parse_json(wu.get("step_progress"),[]):
            if isinstance(sp,dict):
                sid=sp.get("step_id",""); sp_map[sid]=max(sp_map.get(sid,0),sp.get("pct_complete",0))
    bar_h=0.55
    for i,step in enumerate(steps):
        ps=max(1,step.get("planned_start_week",1)); pe=min(12,step.get("planned_end_week",ps)); dur=pe-ps+1
        ax.barh(i,dur,left=ps-1,height=bar_h,color="#d6e8f7",zorder=2)
        ax.barh(i,dur,left=ps-1,height=bar_h,color="none",edgecolor="#0C5595",linewidth=1.0,zorder=3)
        pct=sp_map.get(str(step.get("id","")),0); bar_col="#1E8449" if pct==100 else "#0C5595"
        if pct>0:
            ax.barh(i,dur*pct/100,left=ps-1,height=bar_h,color=bar_col,alpha=0.85,zorder=4)
            if pct>=15:
                ax.text(ps-1+dur*pct/200,i,f"{pct}%",ha="center",va="center",fontsize=7.5,color="white",fontweight="bold",zorder=5)
        if step.get("owner"):
            ax.text(pe+0.1,i,step["owner"],ha="left",va="center",fontsize=7,color="#566573",zorder=5)
        if pct<100 and pe<cur_week:
            ax.barh(i,dur,left=ps-1,height=bar_h,color="none",edgecolor="#DE201B",linewidth=1.5,linestyle="--",zorder=5,alpha=0.7)
    for w in range(13): ax.axvline(w,color="#e5e5e5",linewidth=0.6,zorder=1)
    ax.axvline(cur_week-1,color="#DE201B",linewidth=2.0,linestyle="--",zorder=6,alpha=0.9)
    ax.text(cur_week-0.85,-0.55,f"W{cur_week}",color="#DE201B",fontsize=8,fontweight="bold",zorder=7)
    ax.set_yticks(range(n)); ax.set_yticklabels([s.get("step_name","") for s in steps],fontsize=9,color="#1A1A2E")
    ax.set_xticks(range(13)); ax.set_xticklabels([""]+[f"W{i}" for i in range(1,13)],fontsize=8.5)
    ax.set_xlim(-0.1,12.5); ax.set_ylim(-0.8,n-0.2); ax.invert_yaxis()
    for sp in ["top","right","left","bottom"]: ax.spines[sp].set_visible(False)
    ax.tick_params(left=False,bottom=False)
    patches=[mpatches.Patch(facecolor="#d6e8f7",edgecolor="#0C5595",label="Planned"),
             mpatches.Patch(facecolor="#0C5595",label="In Progress"),
             mpatches.Patch(facecolor="#1E8449",label="Complete"),
             plt.Line2D([0],[0],color="#DE201B",lw=2,ls="--",label=f"Now W{cur_week}")]
    ax.legend(handles=patches,loc="lower right",fontsize=8,frameon=True,framealpha=0.95,edgecolor="#e0e0e0",ncol=4)
    fig.tight_layout(pad=0.8); return fig

def _kpi_chart(kpi, wu_rows):
    baseline=float(kpi.get("baseline_value") or 0); target=float(kpi.get("target_value") or 0)
    wu_sorted=sorted([w for w in wu_rows if w.get("kpi_value") is not None],key=lambda x:x["week_number"])
    weeks=[w["week_number"] for w in wu_sorted]; values=[float(w["kpi_value"]) for w in wu_sorted]
    fig,ax=plt.subplots(figsize=(11,4)); fig.patch.set_facecolor("#ffffff"); ax.set_facecolor("#fafafa")
    ax.axhline(baseline,color="#BDC3C7",lw=1.5,ls=":",label=f"Baseline {baseline}")
    ax.axhline(target,color="#1E8449",lw=2.0,ls="--",label=f"Target {target}")
    if weeks:
        ax.fill_between(weeks,baseline,values,alpha=0.10,color="#0C5595")
        ax.plot(weeks,values,"o-",color="#0C5595",lw=2.5,ms=8,zorder=5,label="Actual")
        for w,v in zip(weeks,values):
            ax.annotate(f"{v:.1f}",(w,v),textcoords="offset points",xytext=(0,10),fontsize=9,ha="center",color="#0C5595",fontweight="bold")
    ax.set_xlim(0.5,12.5); all_vals=values+[baseline,target]; spread=max(abs(target-baseline),1)
    ax.set_ylim(min(all_vals)-spread*0.15,max(all_vals)+spread*0.25)
    ax.set_xticks(range(1,13)); ax.set_xticklabels([f"W{i}" for i in range(1,13)],fontsize=9)
    ax.set_ylabel(f"{kpi.get('kpi_name','')} ({kpi.get('unit','')})",fontsize=9,color="#566573")
    ax.legend(fontsize=9,frameon=False); ax.grid(axis="y",color="#eeeeee",lw=0.6)
    for sp in ["top","right"]: ax.spines[sp].set_visible(False)
    fig.tight_layout(pad=0.5); return fig

def _ramp_chart(weekly_scores, cur_week):
    fig,ax=plt.subplots(figsize=(11,3.5)); fig.patch.set_facecolor("#ffffff"); ax.set_facecolor("#fafafa")
    tw=list(range(1,13)); tt=[TARGET_RAMP[w] for w in tw]
    ax.fill_between(tw,0,tt,alpha=0.06,color="#566573")
    ax.plot(tw,tt,"s--",color="#BDC3C7",lw=1.5,ms=4,label="Target ramp",zorder=2)
    if weekly_scores:
        ws=sorted(weekly_scores.keys()); wv=[weekly_scores[w] for w in ws]
        ax.fill_between(ws,0,wv,alpha=0.12,color="#0C5595")
        ax.plot(ws,wv,"o-",color="#0C5595",lw=2.5,ms=7,label="Actual",zorder=5)
        for w,v in zip(ws,wv):
            ax.annotate(f"{int(v)}",(w,v),textcoords="offset points",xytext=(0,8),fontsize=8,ha="center",color="#0C5595",fontweight="bold")
    ax.axvline(cur_week,color="#DE201B",lw=1.5,ls=":",alpha=0.7)
    ax.set_xlim(0.5,12.5); ax.set_ylim(0,110)
    ax.set_xticks(range(1,13)); ax.set_xticklabels([f"W{i}" for i in range(1,13)],fontsize=9)
    ax.set_ylabel("Score / 100",fontsize=9,color="#566573")
    ax.legend(fontsize=9,frameon=False); ax.grid(axis="y",color="#eeeeee",lw=0.6)
    for sp in ["top","right"]: ax.spines[sp].set_visible(False)
    fig.tight_layout(pad=0.5); return fig

def _dim_bar_chart(q_status):
    DIMS={"Involvement":[1,2,34,35,36],"Method":[8,9,10,11,12,13,14,15,16],
          "Action Plan":[17,18,19,20,21,22,23],"Results":[3,4,5,6,7,24,25],"Stabilisation":[26,27,28,29,30,31,32,33]}
    labels=[]; achieved=[]; possible=[]
    for dim,qs in DIMS.items():
        labels.append(dim)
        achieved.append(sum(q_status[q]["score"] for q in qs if q in q_status))
        possible.append(sum(QUESTIONS[q]["score"] for q in qs))
    pcts=[a/p*100 if p else 0 for a,p in zip(achieved,possible)]
    fig,ax=plt.subplots(figsize=(11,4)); fig.patch.set_facecolor("#ffffff"); ax.set_facecolor("#fafafa")
    colors=["#1E8449" if p>=90 else "#D68910" if p>=60 else "#DE201B" for p in pcts]
    bars=ax.bar(labels,pcts,color=colors,alpha=0.85,zorder=3)
    for bar,a,p_val in zip(bars,achieved,possible):
        ax.text(bar.get_x()+bar.get_width()/2,bar.get_height()+1.5,
                f"{int(bar.get_height())}%\n({a}/{p_val})",ha="center",va="bottom",fontsize=9,color="#1A1A2E",fontweight="bold")
    ax.set_ylim(0,115); ax.set_ylabel("% Achieved",fontsize=9,color="#566573")
    ax.axhline(90,color="#1E8449",lw=1,ls="--",alpha=0.5); ax.axhline(60,color="#D68910",lw=1,ls="--",alpha=0.5)
    for sp in ["top","right"]: ax.spines[sp].set_visible(False)
    ax.grid(axis="y",color="#eeeeee",lw=0.6); fig.tight_layout(pad=0.5); return fig

def _action_summary_chart(actions):
    if not actions: return None
    from collections import Counter
    status_counts=Counter(a.get("status","Open") for a in actions)
    labels=list(status_counts.keys()); vals=list(status_counts.values())
    colors_map={"Open":"#0C5595","In Progress":"#D68910","Completed":"#1E8449","Overdue":"#DE201B"}
    colors=[colors_map.get(l,"#566573") for l in labels]
    fig,ax=plt.subplots(figsize=(8,4)); fig.patch.set_facecolor("#ffffff"); ax.set_facecolor("#fafafa")
    bars=ax.bar(labels,vals,color=colors,alpha=0.85,zorder=3,width=0.5)
    for bar,v in zip(bars,vals):
        ax.text(bar.get_x()+bar.get_width()/2,bar.get_height()+0.1,str(v),ha="center",va="bottom",fontsize=12,fontweight="bold",color="#1A1A2E")
    ax.set_ylabel("Count",fontsize=9,color="#566573")
    for sp in ["top","right"]: ax.spines[sp].set_visible(False)
    ax.grid(axis="y",color="#eeeeee",lw=0.6); fig.tight_layout(pad=0.5); return fig


def _slide_project_overview(project, team, kpi):
    """Project Overview slide — styled card layout matching HTML template."""
    import textwrap
    fig = plt.figure(figsize=(16,9)); fig.patch.set_facecolor("#ffffff")
    ax = fig.add_axes([0,0,1,1]); ax.set_xlim(0,16); ax.set_ylim(0,9); ax.axis("off")
    # title bar
    ax.text(0.5,8.55,"Project Overview",fontsize=22,fontweight="bold",color="#0D68A3",va="center")
    ax.add_patch(mpatches.FancyBboxPatch((0.5,8.28),2.2,0.07,boxstyle="square,pad=0",facecolor="#DE201B",linewidth=0))
    ax.add_patch(mpatches.FancyBboxPatch((2.75,8.28),13.0,0.07,boxstyle="square,pad=0",facecolor="#0D68A3",linewidth=0))
    # LEFT card
    ax.add_patch(mpatches.FancyBboxPatch((0.4,0.3),9.5,7.7,boxstyle="round,pad=0.1",facecolor="#F8FAFC",edgecolor="#E2E8F0",linewidth=1.2))
    # Problem
    ax.add_patch(mpatches.FancyBboxPatch((0.7,6.45),0.52,0.52,boxstyle="round,pad=0.05",facecolor="#EBF8FF",linewidth=0))
    ax.text(0.96,6.71,"!",fontsize=15,fontweight="bold",color="#0D68A3",ha="center",va="center")
    ax.text(1.42,7.02,"PROBLEM STATEMENT",fontsize=8,fontweight="700",color="#64748B")
    prob=project.get("problem_statement","—")
    for li,line in enumerate(textwrap.wrap(prob,width=52)[:3]):
        ax.text(1.42,6.65-li*0.38,line,fontsize=11,color="#1E293B",fontweight="500")
    # Area
    ax.add_patch(mpatches.FancyBboxPatch((0.7,5.25),0.52,0.52,boxstyle="round,pad=0.05",facecolor="#EBF8FF",linewidth=0))
    ax.text(0.96,5.51,"@",fontsize=13,fontweight="bold",color="#0D68A3",ha="center",va="center")
    ax.text(1.42,5.82,"AREA",fontsize=8,fontweight="700",color="#64748B")
    ax.text(1.42,5.47,project.get("target_area","—"),fontsize=13,color="#1E293B",fontweight="600")
    # Timeline
    ax.add_patch(mpatches.FancyBboxPatch((5.2,5.25),0.52,0.52,boxstyle="round,pad=0.05",facecolor="#EBF8FF",linewidth=0))
    ax.text(5.46,5.51,"T",fontsize=13,fontweight="bold",color="#0D68A3",ha="center",va="center")
    ax.text(5.9,5.82,"TIMELINE",fontsize=8,fontweight="700",color="#64748B")
    ax.text(5.9,5.47,f"Launch: {str(project.get('launch_date','—'))[:10]}",fontsize=11,color="#1E293B",fontweight="600")
    ax.text(5.9,5.1,f"Target: {str(project.get('expected_completion_date','—'))[:10]}",fontsize=11,color="#1E293B")
    # KPI
    ax.add_patch(mpatches.FancyBboxPatch((0.7,3.9),0.52,0.52,boxstyle="round,pad=0.05",facecolor="#EBF8FF",linewidth=0))
    ax.text(0.96,4.16,"K",fontsize=13,fontweight="bold",color="#0D68A3",ha="center",va="center")
    ax.text(1.42,4.48,"COMPANY KPI",fontsize=8,fontweight="700",color="#64748B")
    ax.text(1.42,4.12,project.get("company_kpi_link","—"),fontsize=13,color="#0D68A3",fontweight="700")
    ax.add_patch(mpatches.FancyBboxPatch((0.7,3.35),9.0,0.22,boxstyle="round,pad=0.04",facecolor="#E2E8F0",linewidth=0))
    ax.add_patch(mpatches.FancyBboxPatch((0.7,3.35),9.0,0.22,boxstyle="round,pad=0.04",facecolor="#0D68A3",linewidth=0,alpha=0.55))
    ax.text(5.2,3.0,"ACTIVE",fontsize=9,fontweight="700",color="#0D68A3",ha="center")
    # RIGHT card team
    ax.add_patch(mpatches.FancyBboxPatch((10.2,0.3),5.5,7.7,boxstyle="round,pad=0.1",facecolor="#FFFFFF",edgecolor="#E2E8F0",linewidth=1.2))
    ax.add_patch(mpatches.FancyBboxPatch((10.2,7.72),5.5,0.28,boxstyle="square,pad=0",facecolor="#DE201B",linewidth=0))
    ax.text(12.95,8.05,"Project Team",fontsize=16,fontweight="700",color="#1E293B",ha="center",va="center")
    member_colors=["#0D68A3","#DE201B","#1E8449","#D68910","#566573","#8E44AD"]
    for mi,m in enumerate(team[:6]):
        y_m=7.1-mi*1.1; col=member_colors[mi%len(member_colors)]
        ax.add_patch(plt.Circle((10.85,y_m+0.22),0.3,facecolor="#F1F5F9",edgecolor=col if mi==0 else "#BDC3C7",linewidth=2,zorder=3))
        ax.text(10.85,y_m+0.22,m["member_name"][0].upper(),fontsize=11,fontweight="bold",color=col if mi==0 else "#64748B",ha="center",va="center",zorder=4)
        ax.text(11.3,y_m+0.4,m["member_name"],fontsize=12,fontweight="700" if mi==0 else "500",color="#1E293B")
        ax.text(11.3,y_m+0.08,f"{m.get('role','—')}  ·  {m.get('department','')}",fontsize=9,color=col if mi==0 else "#64748B")
    if not team:
        ax.text(12.95,4.5,"No team members yet",fontsize=11,color="#94A3B8",ha="center",va="center")
    fig.tight_layout(pad=0); return fig


def _slide_gap_register(q_status):
    """Gap Register slide — styled list matching HTML template, adapts to any number of gaps."""
    gaps=[(qn,qs) for qn,qs in q_status.items() if qs["due"] and not qs["met"]]
    gaps_sorted=sorted(gaps,key=lambda x:-x[1]["max"])
    n_gaps=len(gaps_sorted)
    if n_gaps==0:
        fig,ax=plt.subplots(figsize=(16,4)); fig.patch.set_facecolor("#ffffff"); ax.axis("off")
        ax.text(8,2,"All due questions met!",fontsize=22,fontweight="bold",color="#1E8449",ha="center",va="center")
        return fig
    row_h=0.70; fig_h=max(5,n_gaps*row_h+2.6)
    fig=plt.figure(figsize=(16,fig_h)); fig.patch.set_facecolor("#ffffff")
    ax=fig.add_axes([0,0,1,1]); ax.set_xlim(0,16); ax.set_ylim(0,fig_h); ax.axis("off")
    ax.text(0.5,fig_h-0.45,"Gap Register",fontsize=22,fontweight="bold",color="#0D68A3",va="center")
    ax.add_patch(mpatches.FancyBboxPatch((0.5,fig_h-0.75),2.2,0.06,boxstyle="square,pad=0",facecolor="#DE201B",linewidth=0))
    ax.add_patch(mpatches.FancyBboxPatch((2.75,fig_h-0.75),13.0,0.06,boxstyle="square,pad=0",facecolor="#0D68A3",linewidth=0))
    card_top=fig_h-1.0; card_h=n_gaps*row_h+1.0
    ax.add_patch(mpatches.FancyBboxPatch((0.4,card_top-card_h),15.2,card_h,boxstyle="round,pad=0.1",facecolor="#FFFFFF",edgecolor="#E2E8F0",linewidth=1.2))
    ax.add_patch(mpatches.FancyBboxPatch((0.4,card_top-0.09),15.2,0.09,boxstyle="square,pad=0",facecolor="#DE201B",linewidth=0))
    ax.text(1.1,card_top-0.38,"GAP REGISTER — Questions Not Yet Met:",fontsize=11,fontweight="700",color="#1E293B")
    ax.plot([0.55,15.5],[card_top-0.58,card_top-0.58],color="#E2E8F0",lw=1.0)
    for ri,(qn,qs) in enumerate(gaps_sorted):
        y=card_top-0.9-ri*row_h
        if ri%2==0:
            ax.add_patch(mpatches.FancyBboxPatch((0.45,y-0.22),15.1,row_h-0.06,boxstyle="square,pad=0",facecolor="#F8FAFC",linewidth=0,zorder=1))
        ax.add_patch(mpatches.FancyBboxPatch((0.65,y-0.16),0.82,0.50,boxstyle="round,pad=0.05",facecolor="#EBF8FF",edgecolor="#90CDF4",linewidth=0.8,zorder=2))
        ax.text(1.06,y+0.09,f"Q{qn}",fontsize=11,fontweight="800",color="#0D68A3",ha="center",va="center",zorder=3)
        ax.add_patch(mpatches.FancyBboxPatch((1.62,y-0.12),0.72,0.42,boxstyle="round,pad=0.05",facecolor="#EBF8FF",edgecolor="#90CDF4",linewidth=0.8,zorder=2))
        ax.text(1.98,y+0.09,f"{qs['max']}pts",fontsize=9,fontweight="700",color="#3182CE",ha="center",va="center",zorder=3)
        d_col="#DE201B" if qs["week_due"]<=2 else "#DD6B20"
        bg_c="#FEF2F2" if qs["week_due"]<=2 else "#FFFAF0"
        bd_c="#FEB2B2" if qs["week_due"]<=2 else "#FBD38D"
        ax.add_patch(mpatches.FancyBboxPatch((2.48,y-0.12),0.92,0.42,boxstyle="round,pad=0.05",facecolor=bg_c,edgecolor=bd_c,linewidth=0.8,zorder=2))
        ax.text(2.94,y+0.09,f"Due W{qs['week_due']}",fontsize=9,fontweight="700",color=d_col,ha="center",va="center",zorder=3)
        ax.text(3.65,y+0.09,qs["text"],fontsize=11,color="#334155",fontweight="500",va="center",zorder=3)
        if ri<n_gaps-1: ax.plot([0.55,15.5],[y-0.26,y-0.26],color="#F1F5F9",lw=0.8)
    fig.tight_layout(pad=0); return fig


def _slide_stabilisation(stab, q_status):
    """Stabilisation Status slide — 8 info cards matching HTML template."""
    stab=stab or {}
    opls_count=len(_parse_json(stab.get("opls"),[]))
    cards=[
        ("CIL Standards Defined","Done" if stab.get("cil_standards_defined") else "Pending","#0D68A3","#EBF8FF",bool(stab.get("cil_standards_defined"))),
        ("CIL Audit Score",f"{stab.get('cil_audit_score','—')}%","#DE201B","#FEF2F2",float(stab.get("cil_audit_score") or 0)>=90),
        ("5S Rating",f"{stab.get('five_s_rating','—')} / 5","#DE201B","#FEF2F2",int(stab.get("five_s_rating") or 0)>=3),
        ("Monitoring In Place","Yes" if stab.get("monitoring_in_place") else "No","#0D68A3","#EBF8FF",bool(stab.get("monitoring_in_place"))),
        ("Monitoring Active","Active" if stab.get("monitoring_active") else "No","#0D68A3","#EBF8FF",bool(stab.get("monitoring_active"))),
        ("Improvements Visible","Visible" if stab.get("improvements_visible") else "No","#0D68A3","#EBF8FF",bool(stab.get("improvements_visible"))),
        ("OPLs Created",str(opls_count),"#DE201B","#FEF2F2",opls_count>0),
        ("Procedures In Place","Yes" if stab.get("procedures_created") else "No","#0D68A3","#EBF8FF",bool(stab.get("procedures_created"))),
    ]
    fig=plt.figure(figsize=(16,9)); fig.patch.set_facecolor("#ffffff")
    ax=fig.add_axes([0,0,1,1]); ax.set_xlim(0,16); ax.set_ylim(0,9); ax.axis("off")
    ax.text(0.5,8.55,"Stabilisation Status",fontsize=22,fontweight="bold",color="#0D68A3",va="center")
    ax.add_patch(mpatches.FancyBboxPatch((0.5,8.28),2.2,0.07,boxstyle="square,pad=0",facecolor="#DE201B",linewidth=0))
    ax.add_patch(mpatches.FancyBboxPatch((2.75,8.28),13.0,0.07,boxstyle="square,pad=0",facecolor="#0D68A3",linewidth=0))
    card_w=3.5; card_h=3.2; gap_x=0.4
    start_x=0.5; row_ys=[4.6,1.1]
    for ci,(title,value,top_col,icon_bg,is_good) in enumerate(cards):
        row=ci//4; col=ci%4
        cx=start_x+col*(card_w+gap_x); cy=row_ys[row]
        ax.add_patch(mpatches.FancyBboxPatch((cx,cy),card_w,card_h,boxstyle="round,pad=0.08",facecolor="#FFFFFF",edgecolor="#E2E8F0",linewidth=1.0,zorder=1))
        ax.add_patch(mpatches.FancyBboxPatch((cx,cy+card_h-0.14),card_w,0.14,boxstyle="square,pad=0",facecolor=top_col,linewidth=0,zorder=2))
        ax.add_patch(plt.Circle((cx+0.55,cy+card_h-0.65),0.28,facecolor=icon_bg,edgecolor="white",linewidth=1.2,zorder=3))
        ax.text(cx+0.55,cy+card_h-0.65,"●",fontsize=9,color=top_col,ha="center",va="center",zorder=4)
        ax.text(cx+0.22,cy+1.45,title.upper(),fontsize=7,fontweight="700",color="#64748B",zorder=3)
        val_col="#4CAF50" if is_good else ("#DE201B" if value in ("No","Pending") else "#94A3B8")
        ax.text(cx+0.22,cy+0.65,value,fontsize=16,fontweight="800",color=val_col,zorder=3)
    fig.tight_layout(pad=0); return fig


def _slide_action_details(actions):
    """Action Plan Details slide — styled table matching template."""
    if not actions:
        fig,ax=plt.subplots(figsize=(16,4)); fig.patch.set_facecolor("#ffffff"); ax.axis("off")
        ax.text(8,2,"No actions recorded yet.",fontsize=16,color="#94A3B8",ha="center",va="center")
        return fig
    row_h=0.65; fig_h=max(5,min(len(actions),18)*row_h+2.5)
    fig=plt.figure(figsize=(16,fig_h)); fig.patch.set_facecolor("#ffffff")
    ax=fig.add_axes([0,0,1,1]); ax.set_xlim(0,16); ax.set_ylim(0,fig_h); ax.axis("off")
    ax.text(0.5,fig_h-0.45,"Action Plan Details",fontsize=22,fontweight="bold",color="#0D68A3",va="center")
    ax.add_patch(mpatches.FancyBboxPatch((0.5,fig_h-0.75),2.2,0.06,boxstyle="square,pad=0",facecolor="#DE201B",linewidth=0))
    ax.add_patch(mpatches.FancyBboxPatch((2.75,fig_h-0.75),13.0,0.06,boxstyle="square,pad=0",facecolor="#0D68A3",linewidth=0))
    card_top=fig_h-1.0; card_h=min(len(actions),18)*row_h+0.9
    ax.add_patch(mpatches.FancyBboxPatch((0.4,card_top-card_h),15.2,card_h,boxstyle="round,pad=0.1",facecolor="#FFFFFF",edgecolor="#E2E8F0",linewidth=1.2))
    ax.add_patch(mpatches.FancyBboxPatch((0.4,card_top-0.09),15.2,0.09,boxstyle="square,pad=0",facecolor="#0D68A3",linewidth=0))
    for htext,hpos in zip(["#","Description","Owner","Due Date","Status"],[0.65,1.8,8.5,11.2,13.2]):
        ax.text(hpos,card_top-0.42,htext,fontsize=9,fontweight="700",color="#64748B")
    ax.plot([0.55,15.5],[card_top-0.6,card_top-0.6],color="#E2E8F0",lw=1.0)
    status_colors={"Open":"#0C5595","In Progress":"#D68910","Completed":"#1E8449","Overdue":"#DE201B"}
    for ri,a in enumerate(actions[:18]):
        y=card_top-1.0-ri*row_h
        if ri%2==0:
            ax.add_patch(mpatches.FancyBboxPatch((0.45,y-0.18),15.1,row_h-0.06,boxstyle="square,pad=0",facecolor="#F8FAFC",linewidth=0,zorder=1))
        is_od=(a.get("target_date") and _safe_date(a["target_date"]) and _safe_date(a["target_date"])<date.today() and a.get("status")!="Completed")
        status="Overdue" if is_od else a.get("status","Open")
        s_col=status_colors.get(status,"#566573")
        ax.text(0.65,y+0.08,str(ri+1),fontsize=10,fontweight="700",color="#0D68A3",va="center",zorder=2)
        ax.text(1.8,y+0.08,a.get("description","")[:55],fontsize=10,color="#334155",va="center",zorder=2)
        ax.text(8.5,y+0.08,a.get("owner","—")[:18],fontsize=10,color="#334155",va="center",zorder=2)
        ax.text(11.2,y+0.08,str(a.get("target_date",""))[:10],fontsize=10,color="#334155",va="center",zorder=2)
        ax.add_patch(mpatches.FancyBboxPatch((13.1,y-0.1),1.8,0.4,boxstyle="round,pad=0.05",facecolor=s_col,linewidth=0,alpha=0.15,zorder=2))
        ax.text(14.0,y+0.1,status,fontsize=9,fontweight="700",color=s_col,ha="center",va="center",zorder=3)
        if ri<min(len(actions),18)-1: ax.plot([0.55,15.5],[y-0.22,y-0.22],color="#F1F5F9",lw=0.6)
    fig.tight_layout(pad=0); return fig


# ─────────────────────────────────────────────
# ANALYSIS CHARTS (5-Why, Fishbone, VSM)
# ─────────────────────────────────────────────
def _five_why_chart(problem, whys):
    items=[("PROBLEM",problem,"#DE201B")]+\
          [(f"WHY {i+1}",w,"#0C5595") for i,w in enumerate(whys) if w.strip()]+\
          [("ROOT CAUSE",whys[-1] if whys and whys[-1].strip() else "","#1E8449")]
    n=len(items); fig_h=n*1.2+0.5
    fig,ax=plt.subplots(figsize=(13,fig_h)); fig.patch.set_facecolor("#ffffff")
    ax.set_xlim(0,10); ax.set_ylim(0,fig_h); ax.axis("off")
    box_h=0.80; y=fig_h-0.4
    for label,text,color in items:
        y-=box_h
        fancy=plt.matplotlib.patches.FancyBboxPatch((0.3,y),9.4,box_h-0.05,boxstyle="round,pad=0.08",
              facecolor=color,edgecolor="white",linewidth=1.5,zorder=2)
        ax.add_patch(fancy)
        ax.text(0.7,y+box_h/2,label,va="center",ha="left",fontsize=8.5,fontweight="bold",color=(1,1,1,0.80),zorder=3)
        wrapped=text if len(text)<=90 else text[:87]+"..."
        ax.text(2.5,y+box_h/2,wrapped,va="center",ha="left",fontsize=9,color="white",zorder=3)
        if y>0.4:
            ax.annotate("",xy=(5,y),xytext=(5,y-0.18),arrowprops=dict(arrowstyle="->",color="#BDC3C7",lw=1.5),zorder=1)
    fig.tight_layout(pad=0.3); return fig

def _fishbone_chart(problem, categories):
    import textwrap
    W,H=22,10; SY=H/2; SX0=1.5; SX1=19.0; HEAD_X=20.5
    fig,ax=plt.subplots(figsize=(20,10)); fig.patch.set_facecolor("#FAFAFA"); ax.set_facecolor("#FAFAFA")
    ax.set_xlim(0,W); ax.set_ylim(0,H); ax.axis("off")
    ax.annotate("",xy=(SX1+0.2,SY),xytext=(SX0,SY),arrowprops=dict(arrowstyle="-|>",color="#1A1A2E",lw=3,mutation_scale=22))
    head=mpatches.FancyBboxPatch((SX1,SY-1.1),W-SX1-0.3,2.2,boxstyle="round,pad=0.12",
         facecolor="#DE201B",edgecolor="white",lw=2,zorder=5)
    ax.add_patch(head)
    prob_lines=textwrap.wrap(problem[:80],width=16)
    for li,line in enumerate(prob_lines[:3]):
        ax.text(HEAD_X,SY+0.35-li*0.55,line,ha="center",va="center",fontsize=9,color="white",fontweight="bold",zorder=6)
    cats=list(categories.items()); n=len(cats)
    top_cats=cats[:math.ceil(n/2)]; bot_cats=cats[math.ceil(n/2):]
    def _x_positions(nb):
        if nb==0: return []
        step=(SX1-SX0-1.0)/(nb+1)
        return [SX0+0.5+step*(i+1) for i in range(nb)]
    def _draw_branch(cat_name,causes,x_attach,side):
        tip_x=max(SX0+0.2,x_attach-3.0); tip_y=SY+side*(H*0.28)
        ax.plot([tip_x,x_attach],[tip_y,SY],color="#0C5595",lw=2.2,zorder=2,solid_capstyle="round")
        lbl_bg=mpatches.FancyBboxPatch((tip_x-1.1,tip_y+side*0.15-0.28),2.2,0.56,boxstyle="round,pad=0.08",
               facecolor="#0C5595",edgecolor="white",lw=1.5,zorder=4)
        ax.add_patch(lbl_bg)
        ax.text(tip_x,tip_y+side*0.15,cat_name,ha="center",va="center",fontsize=9.5,color="white",fontweight="bold",zorder=5)
        if causes:
            n_causes=min(len(causes),6)
            for ci in range(n_causes):
                t=(ci+1)/(n_causes+1)
                cx=tip_x+t*(x_attach-tip_x); cy=tip_y+t*(SY-tip_y); rib_len=0.7
                ax.plot([cx,cx],[cy,cy+side*rib_len],color="#BDC3C7",lw=1.2,zorder=1)
                ax.text(cx,cy+side*(rib_len+0.12),causes[ci][:28],ha="center",
                        va="bottom" if side==1 else "top",fontsize=8,color="#1A1A2E",zorder=3,
                        bbox=dict(boxstyle="round,pad=0.15",facecolor="white",edgecolor="#BDC3C7",lw=0.8,alpha=0.9))
    for (cat,causes),xp in zip(top_cats,_x_positions(len(top_cats))): _draw_branch(cat,causes,xp,+1)
    for (cat,causes),xp in zip(bot_cats,_x_positions(len(bot_cats))): _draw_branch(cat,causes,xp,-1)
    ax.text(W/2,H-0.35,"Fishbone (Ishikawa) Diagram",ha="center",va="top",fontsize=13,fontweight="bold",color="#1A1A2E")
    fig.tight_layout(pad=0.3); return fig

def _vsm_chart(process_steps, title="Value Stream Map"):
    import textwrap
    n=len(process_steps)
    if n==0: return None
    W=max(24,n*4.5+4); H=14
    fig,ax=plt.subplots(figsize=(W*0.9,8)); fig.patch.set_facecolor("#FAFAFA"); ax.set_facecolor("#FAFAFA")
    ax.set_xlim(0,W); ax.set_ylim(0,H); ax.axis("off")
    ax.text(W/2,H-0.4,title,ha="center",va="top",fontsize=13,fontweight="bold",color="#1A1A2E")
    for bx,label,col in [(0.2,"SUPPLIER","#566573"),(W-3.0,"CUSTOMER","#1E8449")]:
        bp=mpatches.FancyBboxPatch((bx,8.5),2.6,1.6,boxstyle="round,pad=0.1",facecolor=col,edgecolor="white",lw=1.5,zorder=3)
        ax.add_patch(bp); ax.text(bx+1.3,9.3,label,ha="center",va="center",fontsize=9,fontweight="bold",color="white",zorder=4)
    step_w=(W-5.0)/n; centres=[2.8+step_w*i+step_w/2 for i in range(n)]
    total_va=0; total_nva=0; total_lt=0
    for i,(step,cx) in enumerate(zip(process_steps,centres)):
        bx=cx-1.6; inv=step.get("inventory_before",0)
        if inv>0:
            tri_x=bx-0.6
            tri=plt.Polygon([[tri_x,7.3],[tri_x+0.7,7.3],[tri_x+0.35,6.7]],facecolor="#D68910",edgecolor="white",lw=1,zorder=3)
            ax.add_patch(tri); ax.text(tri_x+0.35,6.5,str(inv),ha="center",va="top",fontsize=7.5,color="#D68910",fontweight="bold")
        if i<n-1:
            ax.annotate("",xy=(cx+1.6+0.1,7.6),xytext=(cx+1.6+0.5,7.6),
                        arrowprops=dict(arrowstyle="-|>",color="#0C5595",lw=1.5,mutation_scale=12))
            ax.text(cx+1.6+0.3,7.85,"Push",ha="center",va="bottom",fontsize=6.5,color="#0C5595")
        pb=mpatches.FancyBboxPatch((bx,6.0),3.2,1.8,boxstyle="round,pad=0.12",facecolor="#EAF4FB",edgecolor="#0C5595",lw=1.8,zorder=3)
        ax.add_patch(pb)
        name_lines=textwrap.wrap(step.get("name","Step"),width=16)
        for li,ln in enumerate(name_lines[:2]):
            ax.text(cx,7.6-li*0.45,ln,ha="center",va="center",fontsize=9,fontweight="bold",color="#0C5595",zorder=4)
        ops=step.get("operators",1)
        ax.text(bx+0.25,6.25,"👤"*min(ops,4),ha="left",va="bottom",fontsize=8,zorder=4)
        db=mpatches.FancyBboxPatch((bx,3.8),3.2,2.0,boxstyle="round,pad=0.1",facecolor="white",edgecolor="#BDC3C7",lw=1.2,zorder=3)
        ax.add_patch(db)
        ct=step.get("cycle_time",0); co=step.get("changeover_time",0); upt=step.get("uptime",100)
        va=step.get("va_time",0); nva=step.get("nva_time",0)
        for ri,(lbl,val) in enumerate([("C/T",f"{ct}s"),("C/O",f"{co}min"),("Uptime",f"{upt}%"),("VA",f"{va}s"),("NVA",f"{nva}s")]):
            y_r=5.55-ri*0.36
            ax.text(bx+0.18,y_r,lbl+":",ha="left",va="center",fontsize=7.5,color="#566573",zorder=4)
            ax.text(bx+1.35,y_r,val,ha="left",va="center",fontsize=7.5,fontweight="bold",color="#1A1A2E",zorder=4)
        va_w=va/max(ct+nva,1)*3.0 if (ct+nva)>0 else 1.5; nva_w=nva/max(ct+nva,1)*3.0 if (ct+nva)>0 else 1.5
        ax.barh(3.2,nva_w,left=bx,height=0.4,color="#DE201B",alpha=0.75,zorder=3)
        ax.barh(3.2,va_w,left=bx+nva_w,height=0.4,color="#1E8449",alpha=0.85,zorder=3)
        ax.text(cx,2.95,f"VA:{va}s / NVA:{nva}s",ha="center",va="top",fontsize=7,color="#566573",zorder=4)
        total_va+=va; total_nva+=nva; total_lt+=ct
    eff=total_va/max(total_va+total_nva,1)*100
    eff_col="#1E8449" if eff>=60 else "#D68910" if eff>=30 else "#DE201B"
    sx=0.5
    for lbl,val in [("Total Lead Time",f"{total_lt}s"),("Total VA",f"{total_va}s"),("Total NVA",f"{total_nva}s"),("Flow Efficiency",f"{eff:.1f}%")]:
        sb=mpatches.FancyBboxPatch((sx,0.3),4.8,1.4,boxstyle="round,pad=0.1",
           facecolor=eff_col if "Efficiency" in lbl else "#0C5595",edgecolor="white",lw=1.5,zorder=3)
        ax.add_patch(sb)
        ax.text(sx+2.4,1.2,lbl,ha="center",va="center",fontsize=8,color=(1,1,1,0.75),zorder=4)
        ax.text(sx+2.4,0.72,val,ha="center",va="center",fontsize=12,fontweight="bold",color="white",zorder=4)
        sx+=5.3
    fig.tight_layout(pad=0.3); return fig


# ─────────────────────────────────────────────
# ANALYSIS SAVE/LOAD HELPERS
# ─────────────────────────────────────────────
def _save_analysis(supabase, pid, week_number, analysis_type, data, created_by):
    try:
        supabase.table("fi_project_analysis").insert({
            "project_id": pid,
            "week_number": week_number,
            "analysis_type": analysis_type,
            "data": json.dumps(data),
            "created_by": created_by,
        }).execute()
        return True
    except Exception as e:
        st.error(f"Save failed: {e}"); return False

def _load_analyses(supabase, pid, analysis_type):
    try:
        rows = supabase.table("fi_project_analysis").select("*")\
            .eq("project_id", pid).eq("analysis_type", analysis_type)\
            .order("created_at", desc=True).execute().data or []
        return rows
    except: return []

def _delete_analysis(supabase, row_id):
    try:
        supabase.table("fi_project_analysis").delete().eq("id", row_id).execute()
        return True
    except: return False

def _render_saved_analyses(supabase, pid, analysis_type, label, load_callback, cw):
    rows = _load_analyses(supabase, pid, analysis_type)
    if not rows:
        st.caption(f"No saved {label} analyses yet.")
        return
    st.markdown(f"**Saved {label} Analyses**")
    for row in rows:
        rc1, rc2, rc3 = st.columns([3, 1, 1])
        rc1.markdown(f"Week {row.get('week_number','?')} — {row.get('created_by','?')} — {str(row.get('created_at',''))[:10]}")
        if rc2.button("Load", key=f"load_{analysis_type}_{row['id']}"):
            load_callback(json.loads(row["data"]) if isinstance(row["data"], str) else row["data"])
        if rc3.button("Delete", key=f"del_{analysis_type}_{row['id']}"):
            if _delete_analysis(supabase, row["id"]): st.rerun()


# ─────────────────────────────────────────────
# MAIN RENDER
# ─────────────────────────────────────────────
def render_fi_projects_tab(supabase, role, pillar, name):
    can_edit   = (role == "plant_manager") or (role == "pillar_leader" and pillar == "FI")
    is_auditor = role in ["plant_manager", "pillar_leader"]

    # ── project selector ──────────────────────────────────────────────────────
    try:
        all_projects = supabase.table("fi_projects").select("*").order("created_at", desc=True).execute().data or []
    except Exception as e:
        st.error(f"Could not load projects: {e}"); return

    top_l, top_mid, top_r = st.columns([4, 1, 1])
    if not all_projects:
        st.info("No FI projects yet."); selected_project = None
    else:
        proj_map = {p["id"]: f"{p['project_name']}  (W{_cw(p.get('launch_date'))})" for p in all_projects}
        sel_id = top_l.selectbox("Project", list(proj_map.keys()), format_func=lambda x: proj_map[x], key="fi_sel", label_visibility="collapsed")
        selected_project = next((p for p in all_projects if p["id"] == sel_id), None)

    if can_edit and top_r.button("＋ New", key="fi_new_btn"):
        st.session_state["fi_creating"] = True

    # ── delete project (plant_manager only) ──────────────────────────────────
    if role == "plant_manager" and selected_project:
        if top_mid.button("🗑 Delete", key="fi_del_proj_btn"):
            st.session_state["fi_confirm_delete"] = sel_id
        if st.session_state.get("fi_confirm_delete") == sel_id:
            st.warning(f"Permanently delete '{selected_project.get('project_name','')}' and ALL its data? This cannot be undone.")
            conf1, conf2, _ = st.columns([1,1,4])
            if conf1.button("Yes, delete", type="primary", key="fi_del_confirm"):
                try:
                    for tbl in ["fi_weekly_updates","fi_project_steps","fi_project_team",
                                "fi_project_kpi","fi_project_cost","fi_actions",
                                "fi_audit_records","fi_stabilisation","fi_company_kpis",
                                "fi_project_analysis"]:
                        try: supabase.table(tbl).delete().eq("project_id", sel_id).execute()
                        except: pass
                    supabase.table("fi_projects").delete().eq("id", sel_id).execute()
                    st.session_state.pop("fi_confirm_delete", None)
                    st.success("Project deleted."); st.rerun()
                except Exception as e:
                    st.error(f"Delete failed: {e}")
            if conf2.button("Cancel", key="fi_del_cancel"):
                st.session_state.pop("fi_confirm_delete", None); st.rerun()

    # ── create project modal ──────────────────────────────────────────────────
    if st.session_state.get("fi_creating") and can_edit:
        with st.expander("Create New Project", expanded=True):
            c1, c2 = st.columns(2)
            _ns = c1.selectbox("Section", list(PLANT_SECTIONS.keys()), key="fi_ns")
            _nm = PLANT_SECTIONS.get(_ns, [])
            _area_new = f"{_ns} — {c2.selectbox('Machine', ['— All —']+_nm, key='fi_nm')}" if _nm else _ns
            with st.form("fi_create_form"):
                n_name    = st.text_input("Project Name *")
                n_problem = st.text_area("Problem Statement *", height=70)
                nl1, nl2  = st.columns(2)
                n_launch  = nl1.date_input("Launch Date", value=date.today())
                n_end     = nl2.date_input("Expected Completion", value=date.today()+timedelta(weeks=12))
                n_kpi_link = st.text_input("Company KPI this project supports")
                if st.form_submit_button("Create", type="primary"):
                    if n_name and n_problem:
                        supabase.table("fi_projects").insert({
                            "project_name":n_name,"problem_statement":n_problem,
                            "target_area":_area_new,"launch_date":str(n_launch),
                            "expected_completion_date":str(n_end),
                            "company_kpi_link":n_kpi_link,"created_by":name,
                        }).execute()
                        st.session_state["fi_creating"] = False; st.rerun()
                    else:
                        st.error("Name and Problem Statement required.")
        if st.button("Cancel", key="fi_cancel_create"):
            st.session_state["fi_creating"] = False; st.rerun()

    if not selected_project: return

    pid = selected_project["id"]
    cw  = _cw(selected_project.get("launch_date"))

    # ── load all data ─────────────────────────────────────────────────────────
    try:
        team       = supabase.table("fi_project_team").select("*").eq("project_id",pid).execute().data or []
        kpi_rows   = supabase.table("fi_project_kpi").select("*").eq("project_id",pid).execute().data or []
        kpi        = kpi_rows[0] if kpi_rows else {}
        steps      = supabase.table("fi_project_steps").select("*").eq("project_id",pid).order("sort_order").execute().data or []
        cost_rows  = supabase.table("fi_project_cost").select("*").eq("project_id",pid).execute().data or []
        cost       = cost_rows[0] if cost_rows else None
        wu_rows    = supabase.table("fi_weekly_updates").select("*").eq("project_id",pid).order("week_number").execute().data or []
        actions    = supabase.table("fi_actions").select("*").eq("project_id",pid).order("created_at").execute().data or []
        stab_rows  = supabase.table("fi_stabilisation").select("*").eq("project_id",pid).execute().data or []
        stab       = stab_rows[0] if stab_rows else {}
        audit_records = supabase.table("fi_audit_records").select("*").eq("project_id",pid).order("week_number").execute().data or []
    except Exception as e:
        st.error(f"Error loading data: {e}"); return

    wu_by_week = {w["week_number"]: w for w in wu_rows}
    q_status, weekly_scores, total_score = _compute_scores(
        selected_project, team, kpi, steps, wu_rows, actions, stab, audit_records)
    tgt_this_week = TARGET_RAMP.get(cw, 100)

    SECTIONS = ["Overview","Master Plan","Weekly Log","Analysis Tools","Actions","Stabilisation"]
    if is_auditor: SECTIONS.append("Audit")
    sec = st.radio("", SECTIONS, horizontal=True, key="fi_section", label_visibility="collapsed")
    st.divider()


    # ══════════════════════════════════════════════════
    # OVERVIEW
    # ══════════════════════════════════════════════════
    if sec == "Overview":
        h1, h2, h3 = st.columns([3,1,1])
        h1.markdown(f"### {selected_project['project_name']}")
        h1.caption(f"{selected_project.get('target_area','')}   ·   Launch: {str(selected_project.get('launch_date',''))[:10]}")
        h2.metric("Current Week", f"W{cw} / 12")
        _sc = "#1E8449" if total_score>=tgt_this_week else "#D68910" if total_score>=tgt_this_week*0.7 else "#DE201B"
        h3.markdown(f"**Score** &nbsp; <span style='font-size:26px;font-weight:700;color:{_sc}'>{int(total_score)}</span> / 100 &nbsp; *(target {tgt_this_week})*", unsafe_allow_html=True)

        if kpi:
            sub_c = [s for s in _parse_json(kpi.get("sub_components"),[]) if isinstance(s,dict)]
            kpi_vals_s = sorted([w for w in wu_rows if w.get("kpi_value") is not None], key=lambda x:x["week_number"])
            kpi_cur = float(kpi_vals_s[-1]["kpi_value"]) if kpi_vals_s else float(kpi.get("baseline_value") or 0)
            b=float(kpi.get("baseline_value") or 0); t=float(kpi.get("target_value") or 0)
            gap=abs(t-b); prog=abs(kpi_cur-b)/gap*100 if gap>0 else 0
            cols=st.columns(4)
            with cols[0]: _card("Baseline",f"{b} {kpi.get('unit','')}","Start","#17375E")
            with cols[1]: _card("Current",f"{kpi_cur:.1f} {kpi.get('unit','')}",f"Week {cw}","#0C5595")
            with cols[2]: _card("Target",f"{t} {kpi.get('unit','')}","Goal","#145A32")
            with cols[3]: _card("Progress",f"{prog:.0f}%","Toward target","#1E8449" if prog>=80 else "#D68910" if prog>=40 else "#DE201B")
            st.progress(min(int(prog),100))
            fig_kpi=_kpi_chart(kpi,wu_rows); st.pyplot(fig_kpi); plt.close(fig_kpi)

        st.markdown("##### Score vs Target Ramp")
        fig_ramp=_ramp_chart(weekly_scores,cw); st.pyplot(fig_ramp); plt.close(fig_ramp)

        if team:
            st.markdown("##### Team")
            st.dataframe(pd.DataFrame([{"Name":m["member_name"],"Role":m.get("role",""),"Department":m.get("department","")} for m in team]),use_container_width=True,hide_index=True)

        if can_edit:
            with st.expander("Edit Project Details"):
                _cur_area=selected_project.get("target_area","") or ""
                _cur_sec=_cur_area.split(" — ")[0] if " — " in _cur_area else list(PLANT_SECTIONS.keys())[0]
                ea1,ea2=st.columns(2)
                e_sec=ea1.selectbox("Section",list(PLANT_SECTIONS.keys()),index=list(PLANT_SECTIONS.keys()).index(_cur_sec) if _cur_sec in PLANT_SECTIONS else 0,key="fi_e_sec")
                _em=PLANT_SECTIONS.get(e_sec,[])
                _cur_mach=_cur_area.split(" — ")[1] if " — " in _cur_area else ""
                e_mach=ea2.selectbox("Machine",["— All —"]+_em,index=(_em.index(_cur_mach)+1) if _cur_mach in _em else 0,key="fi_e_mach") if _em else None
                e_area=f"{e_sec} — {e_mach}" if e_mach and e_mach!="— All —" else e_sec
                with st.form("fi_edit_project"):
                    ep1,ep2=st.columns(2)
                    e_name=ep1.text_input("Project Name",value=selected_project.get("project_name",""))
                    e_kpi_lnk=ep2.text_input("Company KPI Link",value=selected_project.get("company_kpi_link",""))
                    e_problem=st.text_area("Problem Statement",value=selected_project.get("problem_statement",""),height=70)
                    ed1,ed2=st.columns(2)
                    e_launch=ed1.date_input("Launch Date",value=date.fromisoformat(str(selected_project.get("launch_date",date.today()))[:10]))
                    e_end=ed2.date_input("Expected End",value=date.fromisoformat(str(selected_project.get("expected_completion_date",date.today()+timedelta(weeks=12)))[:10]))
                    if st.form_submit_button("Save"):
                        supabase.table("fi_projects").update({"project_name":e_name,"target_area":e_area,"problem_statement":e_problem,"company_kpi_link":e_kpi_lnk,"launch_date":str(e_launch),"expected_completion_date":str(e_end)}).eq("id",pid).execute()
                        st.success("Saved"); st.rerun()

            with st.expander("Manage Team"):
                with st.form("fi_add_team"):
                    tm1,tm2,tm3=st.columns(3)
                    t_name=tm1.text_input("Name *"); t_role=tm2.selectbox("Role",TEAM_ROLES); t_dept=tm3.text_input("Department")
                    if st.form_submit_button("Add"):
                        if t_name:
                            supabase.table("fi_project_team").insert({"project_id":pid,"member_name":t_name,"role":t_role,"department":t_dept,"contribution_target":""}).execute(); st.rerun()
                if team:
                    del_m=st.selectbox("Remove member",["—"]+[m["member_name"] for m in team],key="fi_del_m")
                    if st.button("Remove",key="fi_rem_m") and del_m!="—":
                        mid=next((m["id"] for m in team if m["member_name"]==del_m),None)
                        if mid: supabase.table("fi_project_team").delete().eq("id",mid).execute(); st.rerun()

            with st.expander("KPI & KAI Setup"):
                _kpi_cats=list(KPI_TREE.keys()); _cur_cat=kpi.get("kpi_category","OEE Improvement")
                _cat_idx=_kpi_cats.index(_cur_cat) if _cur_cat in _kpi_cats else 0
                kpi_cat=st.selectbox("KPI Category",_kpi_cats,index=_cat_idx,key="fi_kpi_cat")
                _def_unit=KPI_TREE[kpi_cat]["unit"]; _kai_opts=KPI_TREE[kpi_cat]["kais"]
                _ex_subs=[s for s in _parse_json(kpi.get("sub_components"),[]) if isinstance(s,dict)]
                _ex_names=[s.get("name","") for s in _ex_subs]; _def_kais=[k for k in _ex_names if k in _kai_opts]
                sel_kais=st.multiselect("KAIs to track",_kai_opts,default=_def_kais,key="fi_kais")
                with st.form("fi_kpi_form"):
                    k1,k2,k3,k4=st.columns(4)
                    k1.caption("Unit"); k2.caption("Baseline"); k3.caption("Target"); k4.caption("Target Date")
                    _cur_unit=kpi.get("unit",_def_unit); _u_idx=UNITS.index(_cur_unit) if _cur_unit in UNITS else 0
                    k_unit=k1.selectbox("Unit",UNITS,index=_u_idx,label_visibility="collapsed")
                    k_base=k2.number_input("Baseline",value=float(kpi.get("baseline_value",0) or 0),label_visibility="collapsed")
                    k_tgt=k3.number_input("Target",value=float(kpi.get("target_value",0) or 0),label_visibility="collapsed")
                    k_date=k4.date_input("Target Date",label_visibility="collapsed",
                        value=date.fromisoformat(str(kpi.get("target_date",date.today()+timedelta(weeks=12)))[:10]) if kpi.get("target_date") else date.today()+timedelta(weeks=12))
                    kai_rows=[]
                    if sel_kais:
                        st.divider()
                        _sub_map={s.get("name",""):s for s in _ex_subs}; _tnames=[""]+[m["member_name"] for m in team]
                        rh=st.columns([3,1.5,1.2,1.2,2])
                        for h,t_h in zip(rh,["KAI","Unit","Baseline","Target","Owner"]): h.caption(t_h)
                        for si,kai in enumerate(sel_kais):
                            ex=_sub_map.get(kai,{}); r=st.columns([3,1.5,1.2,1.2,2]); r[0].markdown(f"**{kai}**")
                            _ku=ex.get("unit",_def_unit); _ko=ex.get("owner","")
                            kai_rows.append({"name":kai,"unit":r[1].selectbox("U",UNITS,index=UNITS.index(_ku) if _ku in UNITS else 0,key=f"ku_{si}",label_visibility="collapsed"),
                                "baseline":r[2].number_input("B",value=float(ex.get("baseline",0)),key=f"kb_{si}",label_visibility="collapsed"),
                                "target":r[3].number_input("T",value=float(ex.get("target",0)),key=f"kt_{si}",label_visibility="collapsed"),
                                "owner":r[4].selectbox("O",_tnames,index=_tnames.index(_ko) if _ko in _tnames else 0,key=f"ko_{si}",label_visibility="collapsed")})
                    if st.form_submit_button("Save KPI & KAIs",type="primary"):
                        payload={"project_id":pid,"kpi_name":kpi_cat,"unit":k_unit,"kpi_category":kpi_cat,"baseline_value":k_base,"target_value":k_tgt,"target_date":str(k_date),"sub_components":json.dumps(kai_rows)}
                        if kpi.get("id"): supabase.table("fi_project_kpi").update(payload).eq("id",kpi["id"]).execute()
                        else: supabase.table("fi_project_kpi").insert(payload).execute()
                        st.success("Saved"); st.rerun()


    # ══════════════════════════════════════════════════
    # MASTER PLAN
    # ══════════════════════════════════════════════════
    elif sec == "Master Plan":
        st.markdown(f"### Master Plan &nbsp; <small style='color:#566573'>Week {cw} of 12</small>", unsafe_allow_html=True)
        if steps:
            fig_g=_gantt(steps,wu_rows,cw)
            if fig_g: st.pyplot(fig_g,use_container_width=True); plt.close(fig_g)
        if can_edit:
            with st.expander("Add / Remove Steps"):
                with st.form("fi_add_step"):
                    s1,s2=st.columns(2); s_name=s1.text_input("Step Name *"); s_desc=s2.text_input("Description")
                    sw1,sw2,sw3=st.columns(3)
                    s_start=sw1.number_input("Start Week",min_value=1,max_value=12,value=1)
                    s_end=sw2.number_input("End Week",min_value=1,max_value=12,value=2)
                    s_owner=sw3.selectbox("Owner",[""]+[m["member_name"] for m in team])
                    if st.form_submit_button("Add Step"):
                        if s_name:
                            supabase.table("fi_project_steps").insert({"project_id":pid,"step_name":s_name,"description":s_desc,"planned_start_week":int(s_start),"planned_end_week":int(s_end),"owner":s_owner,"sort_order":len(steps)}).execute(); st.rerun()
                if steps:
                    del_s=st.selectbox("Remove step",["—"]+[s["step_name"] for s in steps],key="fi_del_s")
                    if st.button("Remove Step",key="fi_rem_s") and del_s!="—":
                        sid=next((s["id"] for s in steps if s["step_name"]==del_s),None)
                        if sid: supabase.table("fi_project_steps").delete().eq("id",sid).execute(); st.rerun()

    # ══════════════════════════════════════════════════
    # WEEKLY LOG
    # ══════════════════════════════════════════════════
    elif sec == "Weekly Log":
        st.markdown("### Weekly Log")
        st.caption("Each week fill in your KPI reading, step progress, and meeting.")
        week_labels={w:f"Week {w} {'(current)' if w==cw else ''}" for w in range(1,13)}
        sel_week=st.selectbox("Select Week",options=list(range(1,13)),index=cw-1,format_func=lambda x:week_labels[x],key="fi_wk_sel")
        wu=wu_by_week.get(sel_week,{}); wu_id=wu.get("id")
        due_qs=[(qn,QUESTIONS[qn]) for qn in QUESTIONS if QUESTIONS[qn]["week"]==sel_week]
        new_qs=[(qn,QUESTIONS[qn]) for qn in QUESTIONS if QUESTIONS[qn]["week"]<=sel_week]
        met_count=sum(1 for qn,_ in new_qs if q_status[qn]["met"])
        cum_score=sum(q_status[qn]["score"] for qn,_ in new_qs)
        cum_max=sum(q["score"] for _,q in new_qs); tgt_w=TARGET_RAMP.get(sel_week,100)
        rc1,rc2,rc3=st.columns(3)
        rc1.metric(f"Cumulative Score (W{sel_week})",f"{cum_score} / {cum_max}")
        rc2.metric("Target Score",tgt_w,delta=f"{cum_score-tgt_w:+.0f}")
        rc3.metric("Requirements Met",f"{met_count} / {len(new_qs)}")
        if due_qs:
            with st.expander(f"New requirements unlocking in Week {sel_week}",expanded=False):
                for qn,qdata in due_qs:
                    icon="✅" if q_status[qn]["met"] else "🔴"
                    st.markdown(f"{icon} **Q{qn}** ({qdata['score']} pts) — {qdata['text']}")

        def _save_wu(updates):
            updates["project_id"]=pid; updates["week_number"]=sel_week
            if wu_id: supabase.table("fi_weekly_updates").update(updates).eq("id",wu_id).execute()
            else: supabase.table("fi_weekly_updates").insert(updates).execute()
            st.rerun()

        with st.expander(f"KPI Reading — Week {sel_week}",expanded=True):
            if not kpi: st.info("Set up your KPI in Overview first.")
            else:
                with st.form(f"fi_kpi_w{sel_week}"):
                    wk1,wk2=st.columns(2)
                    kpi_val_in=wk1.number_input(f"{kpi.get('kpi_name','')} value ({kpi.get('unit','')})",value=float(wu.get("kpi_value",kpi.get("baseline_value",0)) or 0))
                    coll_by=wk2.multiselect("Collected by",[m["member_name"] for m in team],default=[x for x in (wu.get("kpi_collected_by","") or "").split(",") if x.strip() in [m["member_name"] for m in team]])
                    sub_c=[s for s in _parse_json(kpi.get("sub_components"),[]) if isinstance(s,dict)]
                    _wu_kai_data={}
                    if wu.get("kpi_notes"):
                        try: _wu_kai_data=json.loads(wu["kpi_notes"]) if wu["kpi_notes"].startswith("{") else {}
                        except: pass
                    kai_weekly={}
                    if sub_c:
                        st.divider(); st.caption("KAI readings:")
                        chunks=[sub_c[i:i+4] for i in range(0,len(sub_c),4)]; _kai_global=0
                        for chunk in chunks:
                            kcols=st.columns(len(chunk))
                            for ci,kai in enumerate(chunk):
                                kn=kai.get("name",""); ex_v=float(_wu_kai_data.get(kn,{}).get("value",0))
                                with kcols[ci]:
                                    st.caption(f"**{kn}**  _{kai.get('baseline','')}→{kai.get('target','')} {kai.get('unit','')}_")
                                    kai_weekly[kn]={"value":st.number_input(kn,value=ex_v,key=f"kai_{sel_week}_{_kai_global}",label_visibility="collapsed")}
                                _kai_global+=1
                    if st.form_submit_button("Save KPI",type="primary"):
                        _save_wu({"kpi_value":kpi_val_in,"kpi_collected_by":",".join(coll_by),"kpi_notes":json.dumps(kai_weekly),"updated_by":name})

        with st.expander(f"Step Progress — Week {sel_week}",expanded=True):
            due_steps=[s for s in steps if s.get("planned_start_week",1)<=sel_week]
            if not due_steps: st.info("No steps due yet this week.")
            else:
                sp_list=_parse_json(wu.get("step_progress"),[]); sp_map2={sp["step_id"]:sp for sp in sp_list if isinstance(sp,dict)}
                PCT=["0%","25%","50%","75%","100%"]
                with st.form(f"fi_sp_w{sel_week}"):
                    new_sp=[]; chunks2=[due_steps[i:i+3] for i in range(0,len(due_steps),3)]
                    for chunk in chunks2:
                        scols=st.columns(len(chunk))
                        for ci,step in enumerate(chunk):
                            sid=str(step["id"]); ex=sp_map2.get(sid,{})
                            with scols[ci]:
                                st.markdown(f"**{step['step_name']}**")
                                st.caption(f"W{step.get('planned_start_week','')}→W{step.get('planned_end_week','')} · {step.get('owner','') or '—'}")
                                cur_pct=f"{int(ex.get('pct_complete',0))}%"
                                pct_sel=st.selectbox("Progress",PCT,index=PCT.index(cur_pct) if cur_pct in PCT else 0,key=f"spp_{sel_week}_{sid}",label_visibility="collapsed")
                                note=st.text_input("Note",value=ex.get("notes",""),key=f"spn_{sel_week}_{sid}",label_visibility="collapsed",placeholder="Notes…")
                                p_int=int(pct_sel.replace("%",""))
                                new_sp.append({"step_id":sid,"pct_complete":p_int,"status":"Completed" if p_int==100 else "In Progress" if p_int>0 else "Not Started","notes":note})
                    if st.form_submit_button("Save Progress",type="primary"):
                        _save_wu({"step_progress":json.dumps(new_sp),"updated_by":name})

        with st.expander(f"Meeting Log — Week {sel_week}"):
            with st.form(f"fi_mtg_w{sel_week}"):
                mt1,mt2=st.columns(2)
                mtg_held=mt1.checkbox("Meeting held?",value=bool(wu.get("meeting_held")))
                attendees=mt2.multiselect("Attendees",[m["member_name"] for m in team],default=[x.strip() for x in (wu.get("meeting_attendees","") or "").split(",") if x.strip() in [m["member_name"] for m in team]])
                mtg_notes=st.text_area("Notes",value=wu.get("meeting_notes",""),height=60)
                if team: st.caption(f"Attendance: {len(attendees)/len(team)*100:.0f}%  ({len(attendees)}/{len(team)})")
                if st.form_submit_button("Save"):
                    _save_wu({"meeting_held":mtg_held,"meeting_attendees":",".join(attendees),"meeting_notes":mtg_notes,"updated_by":name})

        if sel_week>=4:
            with st.expander(f"Basic Conditions — Week {sel_week}"):
                with st.form(f"fi_bc_w{sel_week}"):
                    bc_done=st.checkbox("Critical areas restored to basic conditions?",value=bool(wu.get("basic_conditions_restored")))
                    if bc_done:
                        bc_desc=st.text_area("Describe",value=wu.get("basic_conditions_description",""),height=60)
                        bc1,bc2=st.columns(2)
                        bc_before=bc1.file_uploader("Before photo",type=["jpg","jpeg","png"],key=f"bcb_{sel_week}")
                        bc_after=bc2.file_uploader("After photo",type=["jpg","jpeg","png"],key=f"bca_{sel_week}")
                    if st.form_submit_button("Save"):
                        upd={"basic_conditions_restored":bc_done,"updated_by":name}
                        if bc_done:
                            upd["basic_conditions_description"]=bc_desc
                            if bc_before: upd["basic_conditions_before_b64"]=_b64(bc_before)
                            if bc_after: upd["basic_conditions_after_b64"]=_b64(bc_after)
                        _save_wu(upd)

        if sel_week>=3:
            with st.expander(f"Root Cause Analysis — Week {sel_week}"):
                with st.form(f"fi_rca_w{sel_week}"):
                    rca_done=st.checkbox("RCA performed?",value=bool(wu.get("rca_performed")))
                    if rca_done:
                        r1,r2=st.columns(2)
                        rca_method=r1.selectbox("Method",RCA_METHODS,index=RCA_METHODS.index(wu.get("rca_method","5-Why")) if wu.get("rca_method") in RCA_METHODS else 0)
                        rca_findings=r2.text_area("Findings",value=wu.get("rca_findings",""),height=80)
                        causes_ver=st.checkbox("Causes verified with data?",value=bool(wu.get("causes_verified")))
                        causes_meth=st.text_input("Verification method",value=wu.get("causes_verification_method","")) if causes_ver else ""
                        shifts_opts=["All","Day only","Night only","Day & Afternoon","Other"]
                        _sc_cur=wu.get("shifts_covered","")
                        shifts_cov=st.selectbox("Shifts data collected from",shifts_opts,index=shifts_opts.index(_sc_cur) if _sc_cur in shifts_opts else 0)
                    if sel_week>=7:
                        st.divider(); st.caption("Reoccurrence Analysis")
                        reoc_b=st.checkbox("Similar problem occurred before?",value=bool(wu.get("reoccurrence_before")))
                        reoc_d=st.text_area("Description & what was done",value=wu.get("reoccurrence_description",""),height=60)
                        reoc_p=st.checkbox("Reoccurrence prevention in place?",value=bool(wu.get("reoccurrence_prevention")))
                        s_pa=st.checkbox("Single problem analysis applied?",value=bool(wu.get("single_problem_analysis")))
                        s_n=st.text_input("Follow-up notes",value=wu.get("single_problem_notes",""))
                    if st.form_submit_button("Save"):
                        upd={"rca_performed":rca_done,"updated_by":name}
                        if rca_done: upd.update({"rca_method":rca_method,"rca_findings":rca_findings,"causes_verified":causes_ver,"causes_verification_method":causes_meth,"shifts_covered":shifts_cov})
                        if sel_week>=7: upd.update({"reoccurrence_before":reoc_b,"reoccurrence_description":reoc_d,"reoccurrence_prevention":reoc_p,"single_problem_analysis":s_pa,"single_problem_notes":s_n})
                        _save_wu(upd)


    # ══════════════════════════════════════════════════
    # ANALYSIS TOOLS  (restyled + save/load)
    # ══════════════════════════════════════════════════
    elif sec == "Analysis Tools":
        st.markdown("### Analysis Tools")
        tool=st.radio("Tool",["5-Why","Fishbone","VSM Builder"],horizontal=True,key="fi_tool")

        # ── 5-WHY ─────────────────────────────────────
        if tool == "5-Why":
            st.markdown("""
            <div style="background:#0C5595;border-radius:10px;padding:14px 20px;margin-bottom:16px">
              <span style="color:white;font-size:16px;font-weight:700">5-Why Analysis</span>
              <span style="color:rgba(255,255,255,.7);font-size:12px;margin-left:12px">Ask "Why?" five times to reach the root cause</span>
            </div>""", unsafe_allow_html=True)

            # Load previously saved
            saved_5why = _load_analyses(supabase, pid, "5why")
            if saved_5why:
                with st.expander(f"📂 Saved 5-Why Analyses ({len(saved_5why)})", expanded=False):
                    for row in saved_5why:
                        d = json.loads(row["data"]) if isinstance(row["data"],str) else row["data"]
                        rc1,rc2,rc3 = st.columns([3,1,1])
                        rc1.markdown(f"**W{row.get('week_number','?')}** — {row.get('created_by','?')} — {str(row.get('created_at',''))[:10]}")
                        rc1.caption(d.get("problem","")[:60])
                        if rc2.button("Load",key=f"load_5why_{row['id']}"):
                            st.session_state["fy_loaded"]=d; st.rerun()
                        if rc3.button("🗑",key=f"del_5why_{row['id']}"):
                            _delete_analysis(supabase,row["id"]); st.rerun()

            _loaded=st.session_state.get("fy_loaded",{})
            prob=st.text_area("Problem Statement",height=70,value=_loaded.get("problem",selected_project.get("problem_statement","")),key="fy_problem")
            whys=[]
            _saved_whys=_loaded.get("whys",["","","","",""])
            for i in range(5):
                col_l,col_r=st.columns([1,10])
                col_l.markdown(f"""<div style="background:#DE201B;border-radius:6px;padding:6px 10px;margin-top:8px;color:white;font-weight:700;text-align:center">W{i+1}</div>""",unsafe_allow_html=True)
                whys.append(col_r.text_input(f"Why {i+1}",value=_saved_whys[i] if i<len(_saved_whys) else "",key=f"fy_{i}",label_visibility="collapsed",placeholder=f"Why {i+1}..."))

            bca,bcb,bcc=st.columns([1,1,2])
            week_sel_5w=bca.number_input("Save for Week",min_value=1,max_value=12,value=cw,key="fy_week")

            if bcb.button("💾 Save Analysis",key="fy_save",type="primary"):
                active=[w for w in whys if w.strip()]
                if prob and active:
                    if _save_analysis(supabase,pid,int(week_sel_5w),"5why",{"problem":prob,"whys":whys},name):
                        st.success("Saved!"); st.session_state.pop("fy_loaded",None); st.rerun()

            if bcc.button("📊 Generate Chart",key="fy_gen"):
                active=[w for w in whys if w.strip()]
                if prob and active:
                    fig_fy=_five_why_chart(prob,active)
                    st.pyplot(fig_fy,use_container_width=True)
                    buf_fy=_fig_to_bytes(fig_fy)
                    st.download_button("Download PNG",data=buf_fy,file_name="5why.png",mime="image/png")
                else:
                    st.warning("Fill in the problem and at least one Why.")

        # ── FISHBONE ──────────────────────────────────
        elif tool == "Fishbone":
            st.markdown("""
            <div style="background:#0C5595;border-radius:10px;padding:14px 20px;margin-bottom:16px">
              <span style="color:white;font-size:16px;font-weight:700">Fishbone (Ishikawa) Diagram</span>
              <span style="color:rgba(255,255,255,.7);font-size:12px;margin-left:12px">Enter causes per category</span>
            </div>""", unsafe_allow_html=True)

            saved_fb = _load_analyses(supabase, pid, "fishbone")
            if saved_fb:
                with st.expander(f"📂 Saved Fishbone Analyses ({len(saved_fb)})", expanded=False):
                    for row in saved_fb:
                        d=json.loads(row["data"]) if isinstance(row["data"],str) else row["data"]
                        rc1,rc2,rc3=st.columns([3,1,1])
                        rc1.markdown(f"**W{row.get('week_number','?')}** — {row.get('created_by','?')} — {str(row.get('created_at',''))[:10]}")
                        rc1.caption(d.get("problem","")[:60])
                        if rc2.button("Load",key=f"load_fb_{row['id']}"):
                            st.session_state["fb_loaded"]=d; st.rerun()
                        if rc3.button("🗑",key=f"del_fb_{row['id']}"):
                            _delete_analysis(supabase,row["id"]); st.rerun()

            _fb_loaded=st.session_state.get("fb_loaded",{})
            prob_fb=st.text_input("Effect / Problem",value=_fb_loaded.get("problem",selected_project.get("problem_statement","")[:80]),key="fb_prob")
            default_cats=["Man","Machine","Method","Material","Measurement","Environment"]
            saved_cats=_fb_loaded.get("categories",{})
            cats_in={}
            n_cols=3; cat_cols=st.columns(n_cols)
            for ci,cat in enumerate(default_cats):
                with cat_cols[ci%n_cols]:
                    existing="\n".join(saved_cats.get(cat,[]))
                    st.markdown(f"""<div style="background:#17375E;color:white;font-weight:700;padding:6px 10px;border-radius:6px 6px 0 0;font-size:13px">{cat}</div>""",unsafe_allow_html=True)
                    raw=st.text_area("",height=90,key=f"fb_{cat}",value=existing,placeholder="One cause per line…",label_visibility="collapsed")
                    cats_in[cat]=[l.strip() for l in raw.split("\n") if l.strip()]

            fba,fbb,fbc=st.columns([1,1,2])
            week_sel_fb=fba.number_input("Save for Week",min_value=1,max_value=12,value=cw,key="fb_week")

            if fbb.button("💾 Save Analysis",key="fb_save",type="primary"):
                filled={k:v for k,v in cats_in.items() if v}
                if prob_fb and filled:
                    if _save_analysis(supabase,pid,int(week_sel_fb),"fishbone",{"problem":prob_fb,"categories":cats_in},name):
                        st.success("Saved!"); st.session_state.pop("fb_loaded",None); st.rerun()

            if fbc.button("🦴 Generate Fishbone",key="fb_gen",type="primary"):
                filled={k:v for k,v in cats_in.items() if v}
                if prob_fb and filled:
                    fig_fb=_fishbone_chart(prob_fb,filled)
                    st.pyplot(fig_fb,use_container_width=True)
                    buf_fb=_fig_to_bytes(fig_fb)
                    st.download_button("Download PNG",data=buf_fb,file_name="fishbone.png",mime="image/png")
                else:
                    st.warning("Fill in the problem and at least one category.")

        # ── VSM ───────────────────────────────────────
        elif tool == "VSM Builder":
            st.markdown("""
            <div style="background:#0C5595;border-radius:10px;padding:14px 20px;margin-bottom:16px">
              <span style="color:white;font-size:16px;font-weight:700">Value Stream Map</span>
              <span style="color:rgba(255,255,255,.7);font-size:12px;margin-left:12px">Build a current-state VSM</span>
            </div>""", unsafe_allow_html=True)

            saved_vsm = _load_analyses(supabase, pid, "vsm")
            if saved_vsm:
                with st.expander(f"📂 Saved VSM Analyses ({len(saved_vsm)})", expanded=False):
                    for row in saved_vsm:
                        d=json.loads(row["data"]) if isinstance(row["data"],str) else row["data"]
                        rc1,rc2,rc3=st.columns([3,1,1])
                        rc1.markdown(f"**W{row.get('week_number','?')}** — {row.get('created_by','?')} — {str(row.get('created_at',''))[:10]}")
                        rc1.caption(d.get("title","")[:60])
                        if rc2.button("Load",key=f"load_vsm_{row['id']}"):
                            st.session_state["vsm_loaded"]=d; st.session_state["vsm_steps"]=d.get("steps",[]); st.rerun()
                        if rc3.button("🗑",key=f"del_vsm_{row['id']}"):
                            _delete_analysis(supabase,row["id"]); st.rerun()

            _vsm_loaded=st.session_state.get("vsm_loaded",{})
            vsm_title=st.text_input("Map Title",value=_vsm_loaded.get("title",f"VSM — {selected_project.get('target_area','')}"),key="vsm_title")

            with st.expander("📐 Recommended Sample Sizes"):
                st.markdown("""| Study | Samples | Notes |\n|---|---|---|\n| **Cycle Time** | 30 min | Increase if high variation |\n| **Changeover** | 10 obs | Cover all shifts |\n| **Work Sampling** | 384+ | 95% confidence ±5% |""")

            if "vsm_steps" not in st.session_state: st.session_state["vsm_steps"]=[]
            with st.form("vsm_add_step"):
                v1,v2,v3=st.columns(3)
                v_name=v1.text_input("Step Name *",placeholder="e.g. Die Cut")
                v_ct=v2.number_input("Cycle Time (sec)",min_value=0,value=0)
                v_co=v3.number_input("Changeover (min)",min_value=0,value=0)
                v4,v5,v6,v7=st.columns(4)
                v_ops=v4.number_input("Operators",min_value=1,max_value=20,value=1)
                v_upt=v5.number_input("Uptime %",min_value=0,max_value=100,value=95)
                v_va=v6.number_input("VA Time (sec)",min_value=0,value=0)
                v_nva=v7.number_input("NVA/Wait (sec)",min_value=0,value=0)
                v_inv=st.number_input("Inventory before this step (units)",min_value=0,value=0)
                if st.form_submit_button("Add Step"):
                    if v_name:
                        st.session_state["vsm_steps"].append({"name":v_name,"cycle_time":v_ct,"changeover_time":v_co,"operators":v_ops,"uptime":v_upt,"va_time":v_va,"nva_time":v_nva,"inventory_before":v_inv})
                        st.rerun()

            if st.session_state.get("vsm_steps"):
                st.dataframe(pd.DataFrame(st.session_state["vsm_steps"]),use_container_width=True)
                vc1,vc2,vc3,vc4=st.columns(4)
                if vc1.button("Clear All",key="vsm_clear"): st.session_state["vsm_steps"]=[]; st.rerun()
                if vc2.button("Remove Last",key="vsm_rem"): st.session_state["vsm_steps"].pop(); st.rerun()
                week_sel_vsm=vc3.number_input("Week",min_value=1,max_value=12,value=cw,key="vsm_week")
                if vc4.button("💾 Save",key="vsm_save",type="primary"):
                    if _save_analysis(supabase,pid,int(week_sel_vsm),"vsm",{"title":vsm_title,"steps":st.session_state["vsm_steps"]},name):
                        st.success("Saved!"); st.session_state.pop("vsm_loaded",None); st.rerun()
                if st.button("🗺 Generate VSM",type="primary",key="vsm_gen"):
                    fig_vsm=_vsm_chart(st.session_state["vsm_steps"],title=vsm_title)
                    st.pyplot(fig_vsm,use_container_width=True)
                    buf_vsm=_fig_to_bytes(fig_vsm)
                    st.download_button("Download VSM PNG",data=buf_vsm,file_name="vsm.png",mime="image/png")
            else:
                st.info("Add your first process step above.")


    # ══════════════════════════════════════════════════
    # ACTIONS
    # ══════════════════════════════════════════════════
    elif sec == "Actions":
        st.markdown("### Action Plan")
        completed=[a for a in actions if a.get("status")=="Completed"]
        in_prog=[a for a in actions if a.get("status")=="In Progress"]
        open_a=[a for a in actions if a.get("status")=="Open"]
        overdue=[a for a in actions if a.get("status")=="Overdue" or
                 (a.get("target_date") and _safe_date(a["target_date"]) and _safe_date(a["target_date"])<date.today() and a.get("status")!="Completed")]
        ac1,ac2,ac3,ac4=st.columns(4)
        ac1.metric("Open",len(open_a)); ac2.metric("In Progress",len(in_prog))
        ac3.metric("Completed",len(completed)); ac4.metric("Overdue",len(overdue),delta=f"-{len(overdue)}" if overdue else None,delta_color="inverse")
        k1,k2,k3=st.columns(3)
        STATUS_COLS={"Open":(k1,"🔵"),"In Progress":(k2,"🟡"),"Completed":(k3,"🟢")}
        for status,(col,icon) in STATUS_COLS.items():
            with col:
                st.markdown(f"**{icon} {status}**")
                for a in [x for x in actions if x.get("status")==status]:
                    is_od=(a.get("target_date") and _safe_date(a["target_date"]) and _safe_date(a["target_date"])<date.today() and status!="Completed")
                    border={"Open":"#0C5595","In Progress":"#D68910","Completed":"#1E8449"}.get(status,"#BDC3C7")
                    if is_od: border="#DE201B"
                    st.markdown(f"""<div style="border-left:4px solid {border};background:#f9f9f9;border-radius:6px;padding:10px 12px;margin-bottom:8px;font-size:13px">
                      <b>{a.get('description','')[:60]}</b><br>
                      <span style="color:#566573;font-size:11px">👤 {a.get('owner','—')} &nbsp; 📅 {str(a.get('target_date',''))[:10]}{' &nbsp; ⚠️ OVERDUE' if is_od else ''}</span>
                    </div>""",unsafe_allow_html=True)
        st.divider()
        if actions and can_edit:
            with st.expander("Update Action Status"):
                act_names=[f"{a['description'][:45]} [{a.get('status','')}]" for a in actions]
                act_idx=st.selectbox("Select action",range(len(actions)),format_func=lambda x:act_names[x],key="fi_act_sel")
                act_obj=actions[act_idx]
                with st.form("fi_upd_act_form"):
                    us1,us2=st.columns(2)
                    new_st=us1.selectbox("New Status",ACTION_STATUSES,index=ACTION_STATUSES.index(act_obj.get("status","Open")))
                    new_ev=us2.file_uploader("Evidence",type=["pdf","png","jpg","xlsx"])
                    if st.form_submit_button("Update Action",type="primary"):
                        upd_a={"status":new_st}
                        if new_st=="Completed": upd_a["completed_date"]=str(date.today())
                        if new_ev: upd_a["evidence_b64"]=_b64(new_ev); upd_a["evidence_filename"]=new_ev.name
                        supabase.table("fi_actions").update(upd_a).eq("id",act_obj["id"]).execute(); st.rerun()
        if can_edit:
            with st.expander("Add New Action"):
                with st.form("fi_add_action"):
                    na1,na2=st.columns(2)
                    a_desc=na1.text_area("Description *",height=70); a_rc=na2.text_area("Root Cause Addressed",height=70)
                    na3,na4=st.columns(2)
                    a_own=na3.selectbox("Owner",[""]+[m["member_name"] for m in team])
                    a_date=na4.date_input("Target Date",value=date.today()+timedelta(weeks=2))
                    a_ev=st.file_uploader("Evidence (optional)",type=["pdf","png","jpg"],key="fi_new_ev")
                    if st.form_submit_button("Add Action"):
                        if a_desc:
                            supabase.table("fi_actions").insert({"project_id":pid,"description":a_desc,"root_cause_addressed":a_rc,"owner":a_own,"target_date":str(a_date),"status":"Open","created_week":cw,"evidence_b64":_b64(a_ev) if a_ev else None,"evidence_filename":a_ev.name if a_ev else None}).execute(); st.rerun()

    # ══════════════════════════════════════════════════
    # STABILISATION
    # ══════════════════════════════════════════════════
    elif sec == "Stabilisation":
        st.markdown("### Stabilisation")
        stab_v=stab or {}
        def _stab_save(d):
            if stab_v.get("id"): supabase.table("fi_stabilisation").update(d).eq("id",stab_v["id"]).execute()
            else: supabase.table("fi_stabilisation").insert({"project_id":pid,**d}).execute()
            st.rerun()

        with st.expander(f"{'✅' if q_status[31]['met'] else '⬜'} CIL Standards {'· UNLOCKS W5' if cw<5 else ''}"):
            with st.form("fi_cil"):
                cil_def=st.checkbox("CIL standards defined?",value=bool(stab_v.get("cil_standards_defined")),disabled=cw<5)
                cil_score=st.number_input("Latest CIL Audit Score (%)",min_value=0.0,max_value=100.0,value=float(stab_v.get("cil_audit_score",0) or 0))
                cil_file=st.file_uploader("Upload CIL audit record",type=["pdf","xlsx","png"],key="fi_cil_file")
                if cil_score>=90: st.success("Meets >=90% target")
                elif cil_score>0: st.warning(f"{cil_score:.0f}% — target >=90%")
                if st.form_submit_button("Save"):
                    d={"cil_standards_defined":cil_def,"cil_audit_score":cil_score}
                    if cil_file: d["cil_file_b64"]=_b64(cil_file)
                    _stab_save(d)

        with st.expander(f"{'✅' if q_status[32]['met'] else '⬜'} Workplace & 5S {'· UNLOCKS W5' if cw<5 else ''}"):
            with st.form("fi_5s"):
                five_s=st.slider("5S Rating",1,5,int(stab_v.get("five_s_rating",1) or 1))
                five_ph=st.file_uploader("5S photos",type=["jpg","jpeg","png"],key="fi_5s_ph")
                five_n=st.text_area("5S notes",value=stab_v.get("five_s_notes",""),height=60)
                imp_vis=st.checkbox("Improvements visible on machine/area?",value=bool(stab_v.get("improvements_visible")))
                imp_ph=st.file_uploader("Improvement photos",type=["jpg","jpeg","png"],key="fi_imp_ph")
                if st.form_submit_button("Save"):
                    d={"five_s_rating":five_s,"five_s_notes":five_n,"improvements_visible":imp_vis}
                    if five_ph: d["five_s_photos_b64"]=_b64(five_ph)
                    if imp_ph: d["improvements_photos_b64"]=_b64(imp_ph)
                    _stab_save(d)

        with st.expander(f"{'✅' if q_status[27]['met'] else '⬜'} Monitoring Systems {'· UNLOCKS W7' if cw<7 else ''}"):
            with st.form("fi_mon"):
                mon_in=st.checkbox("Monitoring systems in place?",value=bool(stab_v.get("monitoring_in_place")),disabled=cw<7)
                mon_types=st.multiselect("Types",MONITORING_TYPES,default=[x.strip() for x in (stab_v.get("monitoring_types","") or "").split(",") if x.strip() in MONITORING_TYPES])
                mon_act=st.checkbox("Actively used?",value=bool(stab_v.get("monitoring_active")))
                mon_date=st.date_input("Last update",value=date.fromisoformat(str(stab_v["monitoring_last_update"])) if stab_v.get("monitoring_last_update") else date.today())
                mon_ev=st.file_uploader("Evidence",type=["pdf","png","jpg"],key="fi_mon_ev")
                if st.form_submit_button("Save"):
                    d={"monitoring_in_place":mon_in,"monitoring_types":",".join(mon_types),"monitoring_active":mon_act,"monitoring_last_update":str(mon_date)}
                    if mon_ev: d["monitoring_evidence_b64"]=_b64(mon_ev)
                    _stab_save(d)

        with st.expander(f"{'✅' if q_status[29]['met'] else '⬜'} OPLs & Training {'· UNLOCKS W10' if cw<10 else ''}"):
            with st.form("fi_opls"):
                ex_opls=_parse_json(stab_v.get("opls"),[])
                n_opls=st.number_input("Number of OPLs/SOPs",min_value=0,max_value=20,value=max(0,len(ex_opls)))
                new_opls=[]
                for oi in range(int(n_opls)):
                    eo=ex_opls[oi] if oi<len(ex_opls) else {}
                    oo1,oo2,oo3=st.columns(3)
                    new_opls.append({"title":oo1.text_input(f"OPL {oi+1} Title",value=eo.get("title",""),key=f"ot_{oi}"),"covers":oo2.text_input("Covers",value=eo.get("covers",""),key=f"oc_{oi}"),"created_by":oo3.selectbox("By",[""]+[m["member_name"] for m in team],key=f"ocb_{oi}")})
                st.divider(); st.caption("Training Matrix")
                ex_tm=_parse_json(stab_v.get("training_matrix"),[]); new_tm=[]
                for mi,mem in enumerate(team):
                    mn=mem["member_name"]; em=next((x for x in ex_tm if x.get("member")==mn),{})
                    tc1,tc2,tc3=st.columns(3)
                    new_tm.append({"member":mn,"opls_trained":tc1.multiselect(mn,[o["title"] for o in new_opls if o["title"]],default=[x for x in (em.get("opls_trained") or []) if x in [o["title"] for o in new_opls]],key=f"tm_{mi}"),"training_date":str(tc2.date_input("Date",value=date.today(),key=f"td_{mi}")),"plan_confirmed":tc3.checkbox("Planned?",value=bool(em.get("plan_confirmed")),key=f"tp_{mi}")})
                if st.form_submit_button("Save"): _stab_save({"opls":json.dumps(new_opls),"training_matrix":json.dumps(new_tm)})

        with st.expander(f"{'✅' if q_status[26]['met'] else '⬜'} Procedures {'· UNLOCKS W10' if cw<10 else ''}"):
            with st.form("fi_proc"):
                proc_done=st.checkbox("Procedures created to hold gains?",value=bool(stab_v.get("procedures_created")),disabled=cw<10)
                ex_procs=_parse_json(stab_v.get("procedures"),[]); n_p=st.number_input("Number of procedures",min_value=0,max_value=10,value=max(0,len(ex_procs))); new_procs=[]
                for pi in range(int(n_p)):
                    ep=ex_procs[pi] if pi<len(ex_procs) else {}; pp1,pp2=st.columns(2)
                    new_procs.append({"name":pp1.text_input(f"Procedure {pi+1}",value=ep.get("name",""),key=f"pn_{pi}"),"description":pp2.text_input("Description",value=ep.get("description",""),key=f"pd_{pi}")})
                if st.form_submit_button("Save"): _stab_save({"procedures_created":proc_done,"procedures":json.dumps(new_procs)})


    # ══════════════════════════════════════════════════
    # AUDIT  (with Generate Report button)
    # ══════════════════════════════════════════════════
    elif sec == "Audit" and is_auditor:
        st.markdown("### Audit View")
        tgt=TARGET_RAMP.get(cw,100); gap_n=total_score-tgt
        col_s="green" if total_score>=tgt else "orange" if total_score>=tgt*0.7 else "red"
        ah1,ah2,ah3,ah4=st.columns(4)
        ah1.markdown(f"**Score** &nbsp;<span style='font-size:28px;font-weight:700;color:{'#1E8449' if col_s=='green' else '#D68910' if col_s=='orange' else '#DE201B'}'>{int(total_score)}</span>/100",unsafe_allow_html=True)
        ah2.metric(f"Target W{cw}",tgt)
        ah3.metric("Gap",f"{'+' if gap_n>=0 else ''}{gap_n:.0f} pts",delta_color="normal" if gap_n>=0 else "inverse")
        ah4.metric("Questions Met",f"{sum(1 for v in q_status.values() if v['met'])}/{len(q_status)}")

        fig_r2=_ramp_chart(weekly_scores,cw); st.pyplot(fig_r2); plt.close(fig_r2)

        st.markdown("#### Dimensions")
        DIMS={"Involvement":[1,2,34,35,36],"Method":[8,9,10,11,12,13,14,15,16],
              "Action Plan":[17,18,19,20,21,22,23],"Results":[3,4,5,6,7,24,25],"Stabilisation":[26,27,28,29,30,31,32,33]}
        dc=st.columns(len(DIMS))
        for ci,(dim,qs) in enumerate(DIMS.items()):
            achieved=sum(q_status[q]["score"] for q in qs if q in q_status)
            possible=sum(QUESTIONS[q]["score"] for q in qs); pct=achieved/possible*100 if possible else 0
            icon="🟢" if pct>=90 else "🟡" if pct>=60 else "🔴"
            dc[ci].metric(f"{icon} {dim}",f"{achieved}/{possible}",f"{pct:.0f}%")

        st.markdown("#### Question Status")
        rows=[]
        for qn,qs in q_status.items():
            status_s="✅ Met" if qs["met"] else ("🔴 Not Met" if qs["due"] else "⏳ Not Due")
            rows.append({"Q":qn,"Question":qs["text"],"Dim":qs["dim"],"Weight":qs["max"],"Due W":qs["week_due"],"Status":status_s,"Score":f"{qs['score']}/{qs['max']}"})
        st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)

        gaps3=[(qn,qs) for qn,qs in q_status.items() if qs["due"] and not qs["met"]]
        if gaps3:
            st.markdown("#### Gap Register")
            for qn,qs in sorted(gaps3,key=lambda x:-x[1]["max"]):
                st.markdown(f"🔴 **Q{qn}** &nbsp; {qs['text']} &nbsp; *({qs['max']} pts — due W{qs['week_due']})*")

        st.divider()
        st.markdown("#### Team Understanding Check")
        with st.form("fi_audit_form"):
            au1,au2=st.columns(2)
            t_member=au1.text_input("Team member tested"); u_pass=au2.checkbox("Explained correctly?")
            a_notes=st.text_area("Auditor notes",height=80)
            if st.form_submit_button("Save Audit Record"):
                qs_save={str(qn):qs["met"] for qn,qs in q_status.items()}
                if u_pass: qs_save["34"]=True; qs_save["35"]=True
                ar_payload={"project_id":pid,"week_number":cw,"question_scores":json.dumps(qs_save),"total_score":total_score,
                            "team_understanding_tested":bool(t_member),"member_tested":t_member,"understanding_pass":u_pass,
                            "auditor_notes":a_notes,"audited_by":name}
                ex_ar=next((ar for ar in audit_records if ar["week_number"]==cw),None)
                if ex_ar: supabase.table("fi_audit_records").update(ar_payload).eq("id",ex_ar["id"]).execute()
                else: supabase.table("fi_audit_records").insert(ar_payload).execute()
                st.success("Saved"); st.rerun()

        # ── GENERATE REPORT ───────────────────────────
        st.divider()
        st.markdown("#### 📄 Generate Audit Report")
        rpt_col1, rpt_col2 = st.columns([1,3])
        rpt_week = rpt_col1.number_input("Report Week", min_value=1, max_value=12, value=cw, key="rpt_week")

        if rpt_col2.button("📊 Generate PDF Report", type="primary", key="rpt_gen"):
            with st.spinner("Building report..."):
                try:
                    # Load logo
                    logo_reader = None
                    try:
                        import base64 as _b64mod, tempfile as _tf
                        # Logo is stored in 4_QM.py — attempt to load from session or skip
                        if "napco_logo_reader" in st.session_state:
                            logo_reader = st.session_state["napco_logo_reader"]
                    except: pass

                    slides = []
                    _proj_name = selected_project.get("project_name","FI Project")
                    _area      = selected_project.get("target_area","")
                    _today_str = date.today().strftime("%d – %B – %Y")
                    _subtitle  = f"Easternpak · {_area} · Week {rpt_week}"

                    # 1. Cover
                    slides.append({"cover":True,"title":_proj_name,"subtitle":_subtitle,"date":_today_str})

                    # 2. Project Overview
                    slides.append({"section":True,"title":"Project Overview"})

                    # Project Overview as rendered figure
                    fig_ov = _slide_project_overview(selected_project, team, kpi)
                    slides.append({"title":"Project Overview","fig":fig_ov})

                    # 3. KPI Trend
                    if kpi and wu_rows:
                        fig_kpi2 = _kpi_chart(kpi, wu_rows)
                        slides.append({"title":f"KPI Trend — {kpi.get('kpi_name','')}","fig":fig_kpi2})

                    # 4. Score vs Ramp
                    fig_ramp2 = _ramp_chart(weekly_scores, rpt_week)
                    slides.append({"title":f"Score vs Target Ramp — Week {rpt_week}","fig":fig_ramp2})

                    # 5. Master Plan
                    slides.append({"section":True,"title":"Master Plan"})
                    if steps:
                        fig_g2 = _gantt(steps, wu_rows, rpt_week)
                        if fig_g2: slides.append({"title":"Master Plan — Gantt Chart","fig":fig_g2})

                    # 6. Score & Progress
                    slides.append({"section":True,"title":"Score & Progress"})
                    fig_dim = _dim_bar_chart(q_status)
                    slides.append({"title":"Dimension Breakdown","fig":fig_dim})

                    # Gap Register as rendered figure
                    fig_gap = _slide_gap_register(q_status)
                    slides.append({"title":"Gap Register","fig":fig_gap})

                    # 7. Analysis
                    slides.append({"section":True,"title":"Analysis"})

                    # 5-Why saved analyses
                    saved_5why_rpt = _load_analyses(supabase, pid, "5why")
                    for row in saved_5why_rpt:
                        d = json.loads(row["data"]) if isinstance(row["data"],str) else row["data"]
                        whys_r = [w for w in d.get("whys",[]) if w.strip()]
                        if d.get("problem") and whys_r:
                            fig_fy_r = _five_why_chart(d["problem"], whys_r)
                            slides.append({"title":f"5-Why Analysis — W{row.get('week_number','?')}","fig":fig_fy_r})

                    # Fishbone saved analyses
                    saved_fb_rpt = _load_analyses(supabase, pid, "fishbone")
                    for row in saved_fb_rpt:
                        d = json.loads(row["data"]) if isinstance(row["data"],str) else row["data"]
                        filled = {k:v for k,v in d.get("categories",{}).items() if v}
                        if d.get("problem") and filled:
                            fig_fb_r = _fishbone_chart(d["problem"], filled)
                            slides.append({"title":f"Fishbone Analysis — W{row.get('week_number','?')}","fig":fig_fb_r})

                    # VSM saved analyses
                    saved_vsm_rpt = _load_analyses(supabase, pid, "vsm")
                    for row in saved_vsm_rpt:
                        d = json.loads(row["data"]) if isinstance(row["data"],str) else row["data"]
                        if d.get("steps"):
                            fig_vsm_r = _vsm_chart(d["steps"], title=d.get("title","VSM"))
                            if fig_vsm_r:
                                slides.append({"title":f"Value Stream Map — W{row.get('week_number','?')}","fig":fig_vsm_r})

                    if not (saved_5why_rpt or saved_fb_rpt or saved_vsm_rpt):
                        slides.append({"title":"Analysis","text_lines":["No analysis tools saved yet.","Use the Analysis Tools section to create and save 5-Why, Fishbone, or VSM analyses."]})

                    # 8. Action Plan
                    slides.append({"section":True,"title":"Action Plan"})
                    if actions:
                        fig_act = _action_summary_chart(actions)
                        if fig_act: slides.append({"title":"Action Plan Summary","fig":fig_act})
                        fig_act_det = _slide_action_details(actions)
                        slides.append({"title":"Action Plan Details","fig":fig_act_det})
                    else:
                        fig_act_det = _slide_action_details([])
                        slides.append({"title":"Action Plan","fig":fig_act_det})

                    # 9. Stabilisation
                    slides.append({"section":True,"title":"Stabilisation"})
                    fig_stab = _slide_stabilisation(stab, q_status)
                    slides.append({"title":"Stabilisation Status","fig":fig_stab})

                    # Build PDF
                    pdf_buf = _build_fi_pdf(slides, logo_reader=logo_reader)
                    _fname = f"FI_Report_{_proj_name.replace(' ','_')}_W{rpt_week}_{date.today().strftime('%Y%m%d')}.pdf"
                    st.success(f"✅ Report ready — {len(slides)} slides")
                    st.download_button(
                        label="📥 Download PDF Report",
                        data=pdf_buf,
                        file_name=_fname,
                        mime="application/pdf",
                        key="dl_fi_report"
                    )
                except Exception as e:
                    st.error(f"Report generation failed: {e}")
                    import traceback; st.code(traceback.format_exc())
