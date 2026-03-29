# pages/5_FI.py - Focused Improvement Pillar
import os, sys, io
import xlrd
import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import matplotlib as mpl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ["MPLCONFIGDIR"] = os.path.join(os.getcwd(), ".mplconfig")

from utils.supabase_client import get_supabase

st.set_page_config(page_title="FI Pillar", page_icon="💡", layout="wide")

if "user" not in st.session_state:
    st.switch_page("app.py")

supabase = get_supabase()
role     = st.session_state.get("role", "member")
pillar   = st.session_state.get("pillar", "ALL")
name     = st.session_state.get("name", "User")
can_edit = (role == "plant_manager") or (role == "pillar_leader" and pillar == "FI")

col1, col2, col3 = st.columns([5, 1, 1])
with col1:
    st.markdown("# 💡 Focused Improvement")
    st.markdown(f"Logged in as **{name}** · `{role}` · {'✏️ Full Access' if can_edit else '👁️ View Only'}")
with col3:
    if st.button("🏠 Home", use_container_width=True):
        st.switch_page("pages/1_Home.py")

st.divider()

tab1, = st.tabs(["📊 OEE Report & Analysis"])

mpl.rcParams["font.family"] = "DejaVu Sans"
mpl.rcParams["font.size"]   = 11

def _oee_bar_chart(machines_df, oee_col, section_colors):
    fig, ax = plt.subplots(figsize=(13.33, 5), dpi=150)
    x = np.arange(len(machines_df))
    machine_col = "Machine" if "Machine" in machines_df.columns else "Machine Name"
    colors = [section_colors.get(str(s), "#AAAAAA") for s in machines_df["Section"].fillna("")]
    bars = ax.bar(x, machines_df[oee_col].fillna(0), color=colors, width=0.6, alpha=0.9)
    for bar, val in zip(bars, machines_df[oee_col].fillna(0)):
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.5,
                f"{val:.1f}%", ha="center", va="bottom", fontsize=9, color="#000000")
    ax.axhline(85, color="#DE201B", linewidth=1.5, linestyle="--", label="Target 85%")
    ax.set_xticks(x)
    ax.set_xticklabels(machines_df[machine_col].tolist(), rotation=35, ha="right", fontsize=9)
    ax.set_ylabel("OEE (%)"); ax.set_ylim(0, 115)
    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.grid(False)
    from matplotlib.patches import Patch
    legend_patches = [Patch(facecolor=v, label=k) for k,v in section_colors.items()
                      if k in machines_df["Section"].values]
    legend_patches.append(plt.Line2D([0],[0], color="#DE201B", linewidth=1.5,
                                     linestyle="--", label="Target 85%"))
    ax.legend(handles=legend_patches, loc="upper center", bbox_to_anchor=(0.5,-0.22),
              ncol=5, frameon=False, fontsize=9)
    plt.tight_layout(rect=[0,0.10,1,1])
    buf = io.BytesIO(); fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    buf.seek(0); plt.close(fig); return buf

def _arq_chart(machines_df, avl_col, rate_col, qual_col):
    machine_col = "Machine" if "Machine" in machines_df.columns else "Machine Name"
    fig, ax = plt.subplots(figsize=(13.33, 5), dpi=150)
    x = np.arange(len(machines_df)); w = 0.25
    b1 = ax.bar(x-w, machines_df[avl_col].fillna(0),  w, color="#006394", label="Availability")
    b2 = ax.bar(x,   machines_df[rate_col].fillna(0), w, color="#C1A02E", label="Rate")
    b3 = ax.bar(x+w, machines_df[qual_col].fillna(0), w, color="#D8C37D", label="Quality")
    for b in [b1,b2,b3]:
        for bar in b:
            h = bar.get_height()
            if h > 0:
                ax.text(bar.get_x()+bar.get_width()/2, h+0.5, f"{h:.1f}%",
                        ha="center", va="bottom", fontsize=8, color="#000000")
    ax.set_xticks(x)
    ax.set_xticklabels(machines_df[machine_col].tolist(), rotation=35, ha="right", fontsize=9)
    ax.set_ylabel("%"); ax.set_ylim(0, 115)
    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.grid(False)
    ax.legend(loc="upper center", bbox_to_anchor=(0.5,-0.22), ncol=3, frameon=False, fontsize=10)
    plt.tight_layout(rect=[0,0.10,1,1])
    buf = io.BytesIO(); fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    buf.seek(0); plt.close(fig); return buf

def _build_excel(df, title, header_color="006394"):
    wb = Workbook(); ws = wb.active; ws.title = title
    header_fill  = PatternFill("solid", fgColor=header_color)
    header_font  = Font(bold=True, color="FFFFFF", size=10)
    section_fill = PatternFill("solid", fgColor="EAF4FB")
    section_font = Font(bold=True, size=10)
    total_fill   = PatternFill("solid", fgColor="FFF9E6")
    total_font   = Font(bold=True, size=9)
    normal_font  = Font(size=9)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align   = Alignment(horizontal="left",   vertical="center")
    thin = Border(left=Side(style="thin"), right=Side(style="thin"),
                  top=Side(style="thin"),  bottom=Side(style="thin"))
    skip_label_cols = {"Capacity Utilization %","PM & Cleaning (label)","Micro-Stop (label)"}
    headers = [c for c in df.columns if c not in skip_label_cols]
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.fill = header_fill; cell.font = header_font
        cell.alignment = center_align; cell.border = thin
    ws.row_dimensions[1].height = 30
    data_row = 2; current_section = None
    for _, row in df.iterrows():
        sec = str(row.get("Section",""))
        if sec != current_section and sec not in ("TOTAL","Unknown",""):
            current_section = sec
            cell = ws.cell(row=data_row, column=1, value=f"── {sec} ──")
            cell.fill = section_fill; cell.font = section_font; cell.alignment = left_align
            ws.merge_cells(start_row=data_row, start_column=1,
                           end_row=data_row, end_column=len(headers))
            ws.row_dimensions[data_row].height = 18; data_row += 1
        is_total = row.get("Row Type") == "Total"
        for ci, h in enumerate(headers, 1):
            val  = row.get(h)
            cell = ws.cell(row=data_row, column=ci, value=val)
            cell.border = thin
            cell.alignment = center_align if ci > 2 else left_align
            cell.fill = total_fill if is_total else PatternFill()
            cell.font = total_font if is_total else normal_font
        ws.row_dimensions[data_row].height = 16; data_row += 1
    for ci, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(ci)].width = max(12, len(str(h))+2)
    ws.column_dimensions["A"].width = 10; ws.column_dimensions["B"].width = 24
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf

# ── Corrugator parser ──
CORR_SECTIONS = {"BHS", "FOSBER", "SINGLE FACER"}
CORR_MACHINE_COLS = [
    "Shift Time","Run Duration (Hrs)","Stop Duration Hrs","Idle Duration Hrs",
    "# of Stops","# of Change Over","# of Quality Change",
    "Linear Meters","Standard Output M","Square Meters",
    "Availability (%)","Rate (%)","Quality (%)","OEE (%)","Avg Max Speed"
]
CORR_TOTAL_COLS = [
    "Shift Time","Run Duration (Hrs)","Stop Duration Hrs","Idle Duration Hrs",
    "# of Stops","# of Change Over","# of Quality Change",
    "Linear Meters","Standard Output M","Square Meters",
    "Availability (%)","Rate (%)","Quality (%)","OEE (%)"
]

def parse_corrugator(file_bytes):
    wb = xlrd.open_workbook(file_contents=file_bytes)
    sh = wb.sheet_by_index(0)
    current_section = None; all_rows = []
    for row_idx in range(sh.nrows):
        row   = [sh.cell_value(row_idx, c) for c in range(sh.ncols)]
        col_a = str(row[0]).strip()
        if col_a in CORR_SECTIONS:
            current_section = col_a; continue
        if not any(str(v).strip() for v in row): continue
        if isinstance(row[0], str) and col_a and col_a not in CORR_SECTIONS:
            rec = {"Section": current_section, "Machine": col_a, "Row Type": "Machine"}
            for i, cn in enumerate(CORR_MACHINE_COLS):
                rec[cn] = row[i+1] if (i+1) < len(row) else None
            for idx, key in [(15,"Available Time (Hrs)"),(17,"Average Length/Run"),(19,"Average SQM/Run"),
                             (21,"PM & Cleaning"),(22,"MT"),(23,"Speed Value"),(25,"Average Width/Run"),
                             (27,"Average SQM/Hrs"),(28,"Operator"),(30,"Average Sheet Length"),(32,"Average Sheet Width")]:
                rec[key] = row[idx] if idx < len(row) else None
            all_rows.append(rec)
        elif isinstance(row[0], (int, float)) and col_a:
            rec = {"Section": current_section, "Machine": "TOTAL", "Row Type": "Total"}
            for i, cn in enumerate(CORR_TOTAL_COLS):
                rec[cn] = row[i] if i < len(row) else None
            for idx, key in [(15,"Available Time (Hrs)"),(17,"PM & Cleaning"),(19,"Average Length/Run"),
                             (21,"Average SQM/Run"),(22,"MT"),(24,"Average Width/Run"),(26,"Average SQM/Hrs"),
                             (27,"Speed Value"),(29,"# of ProgramID"),(31,"Average Sheet Length"),(33,"Average Sheet Width"),
                             (35,"# of Change Overs"),(37,"Capacity Utilization (%)"),(39,"# of Quality Changes"),
                             (41,"% of Quality Changes"),(43,"Planning Efficiency")]:
                rec[key] = row[idx] if idx < len(row) else None
            all_rows.append(rec)
    df = pd.DataFrame(all_rows)
    skip = {"Section","Machine","Row Type","Operator"}
    for col in df.columns:
        if col not in skip:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df

# ── Converting parser ──
CONV_MAP = {
    "BOBST 160-II":"Die Cutters","BOBST 203":"Die Cutters",
    "BOBST MASTERCUT 1":"Die Cutters","BOBST MASTERCUT 2":"Die Cutters",
    "BOBST 924":"FFG","LMC FFG":"FFG","MARTIN 616":"FFG","SATURN":"FFG",
    "IPACK":"Printer","VEGA 2":"Folder Gluers","BAHMÜLLER TURBOX":"Folder Gluers",
    "BAHMULLER STITCHER":"Stitcher","JUMBO":"Jumbo",
    "SINGLE FACER":"Single Facer","SHRINK-WRAPPER":"Shrink Wrapper",
}
CONV_MAP_NORM = {k.upper().strip(): v for k,v in CONV_MAP.items()}

def parse_converting(file_bytes, filename=""):
    if filename.lower().endswith(".xlsx"):
        import openpyxl
        wb2 = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        ws2 = wb2.active
        rows_raw = [[c if c is not None else "" for c in row]
                    for row in ws2.iter_rows(values_only=True)]
    else:
        wb2 = xlrd.open_workbook(file_contents=file_bytes)
        ws2 = wb2.sheet_by_index(0)
        rows_raw = [[ws2.cell_value(r,c) for c in range(ws2.ncols)] for r in range(ws2.nrows)]

    all_rows = []; total_weight = None; full_qty_weight = None
    for row in rows_raw[2:]:
        col_a = str(row[0]).strip() if row else ""
        if not col_a: continue
        if "total weight" in col_a.lower():
            total_weight    = row[1] if len(row)>1 else None
            full_qty_weight = row[2] if len(row)>2 else None
            break
        if col_a.lower().startswith("page"): continue
        section = CONV_MAP_NORM.get(col_a.upper(), "Unknown")
        data_cells = [v for v in row[1:] if str(v).strip() not in ("","0","0.0")]
        if len(data_cells) < 3: continue
        col_map = {
            "Hits / Hour":1,"Standard Hits (Rhr)":2,
            "Capacity Utilization %":3,"Capacity Utilization Value":4,
            "Shift Time (Hrs)":5,"Available Time (Hrs)":6,"Run Duration (Hrs)":7,
            "STD Setup (min)":8,"Setup Duration (min)":9,
            "Stop Duration (Hrs)":10,"Idle Duration (Hrs)":11,
            "# of Setup":12,"# of Orders":13,"SQM":14,"Actual Hits":15,
            "Hits / Avl Hour":16,"Hits Rhr+Setup":17,
            "Availability %":18,"Rate %":19,"Quality %":20,"OEE %":21,
            "FG Produced":22,"PM & Cleaning (label)":23,"PM & Cleaning Time":24,
            "Micro-Stop (label)":25,"Micro Stop (min)":26,
            "Unknown Value":27,"Full Qty Weight":28,
        }
        rec = {"Section": section, "Machine Name": col_a, "Row Type": "Machine"}
        for cn, idx in col_map.items():
            rec[cn] = row[idx] if idx < len(row) else None
        all_rows.append(rec)

    if total_weight is not None:
        all_rows.append({
            "Section":"TOTAL","Machine Name":"TOTAL","Row Type":"Total",
            "FG Produced": pd.to_numeric(total_weight, errors="coerce"),
            "Full Qty Weight": pd.to_numeric(full_qty_weight, errors="coerce"),
        })
    df = pd.DataFrame(all_rows)
    skip = {"Section","Machine Name","Row Type","Capacity Utilization %",
            "PM & Cleaning (label)","Micro-Stop (label)"}
    for col in df.columns:
        if col not in skip:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df

# ════════════════════════════════════════
# TAB 1 — OEE REPORT & ANALYSIS
# ════════════════════════════════════════
with tab1:
    st.markdown("### 📊 OEE Report & Analysis")
    st.caption("Upload one or both OEE report files, then click Generate Analysis.")

    _uc1, _uc2 = st.columns(2)
    oee_file  = _uc1.file_uploader("📂 Corrugator OEE Report (.xls/.xlsx)",
                                    type=["xls","xlsx"], key="fi_oee_file")
    conv_file = _uc2.file_uploader("📂 Converting OEE Report (.xls/.xlsx)",
                                    type=["xls","xlsx"], key="fi_conv_file")

    if not oee_file and not conv_file:
        st.info("⬆️ Upload at least one OEE report file above.")

    if st.button("🔄 Generate OEE Analysis", type="primary", key="fi_oee_run"):
        st.session_state["fi_run_oee"] = True

    if st.session_state.get("fi_run_oee") and (oee_file or conv_file):

        # ── CORRUGATOR ──
        if oee_file:
            st.divider()
            st.markdown("## 🏭 Corrugator OEE")
            try:
                with st.spinner("Parsing Corrugator file..."):
                    df_corr = parse_corrugator(oee_file.getvalue())
                _corr_m = df_corr[df_corr["Row Type"]=="Machine"]
                st.success(f"✅ Corrugator: {len(_corr_m)} machines extracted")

                _cc1, _cc2 = st.columns([2,5])
                _sel_cs  = _cc1.selectbox("Filter Section",
                               ["All"]+sorted(df_corr["Section"].dropna().unique().tolist()),
                               key="fi_corr_sec")
                _sel_ct  = _cc2.multiselect("Row Type", ["Machine","Total"],
                               default=["Machine","Total"], key="fi_corr_type")
                _dfcv = df_corr.copy()
                if _sel_cs != "All": _dfcv = _dfcv[_dfcv["Section"]==_sel_cs]
                if _sel_ct:         _dfcv = _dfcv[_dfcv["Row Type"].isin(_sel_ct)]

                st.markdown("#### OEE Summary")
                _ct = df_corr[df_corr["Row Type"]=="Total"]
                if not _ct.empty:
                    _tcols = st.columns(len(_ct))
                    for i,(_, tr) in enumerate(_ct.iterrows()):
                        _tcols[i].markdown(f"**{tr.get('Section','')}**")
                        _tcols[i].metric("OEE",          f"{tr.get('OEE (%)',0) or 0:.1f}%")
                        _tcols[i].metric("Availability", f"{tr.get('Availability (%)',0) or 0:.1f}%")
                        _tcols[i].metric("Rate",         f"{tr.get('Rate (%)',0) or 0:.1f}%")
                        _tcols[i].metric("Quality",      f"{tr.get('Quality (%)',0) or 0:.1f}%")

                st.markdown("#### Extracted Data")
                st.dataframe(_dfcv, use_container_width=True, hide_index=True)

                CORR_COLORS = {"BHS":"#006394","FOSBER":"#C1A02E","SINGLE FACER":"#D8C37D"}
                if not _corr_m.empty and "OEE (%)" in _corr_m.columns:
                    st.markdown("#### OEE by Machine")
                    st.image(_oee_bar_chart(_corr_m, "OEE (%)", CORR_COLORS))
                    st.markdown("#### Availability / Rate / Quality")
                    st.image(_arq_chart(_corr_m, "Availability (%)","Rate (%)","Quality (%)"))

                st.markdown("#### Export")
                st.download_button("📥 Download Corrugator OEE (.xlsx)",
                    data=_build_excel(df_corr,"Corrugator OEE","006394"),
                    file_name=f"Corrugator_OEE_{oee_file.name.replace('.xls','').replace('.xlsx','')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="fi_corr_dl")
            except Exception as e:
                st.error(f"Corrugator Error: {e}")

        # ── CONVERTING ──
        if conv_file:
            st.divider()
            st.markdown("## 🏭 Converting OEE")
            try:
                with st.spinner("Parsing Converting file..."):
                    df_conv = parse_converting(conv_file.getvalue(), conv_file.name)
                _conv_m = df_conv[df_conv["Row Type"]=="Machine"]
                st.success(f"✅ Converting: {len(_conv_m)} machines extracted")

                _vc1, = st.columns([1])
                _sel_vs = _vc1.selectbox("Filter Section",
                    ["All"]+sorted([s for s in df_conv["Section"].dropna().unique()
                                    if s not in ("TOTAL","Unknown")]),
                    key="fi_conv_sec")
                _dfvv = _conv_m.copy()
                if _sel_vs != "All": _dfvv = _dfvv[_dfvv["Section"]==_sel_vs]

                st.markdown("#### OEE Summary by Section")
                _sec_ord = ["Die Cutters","FFG","Printer","Folder Gluers",
                            "Stitcher","Jumbo","Single Facer","Shrink Wrapper"]
                _scols = st.columns(min(4, len(_sec_ord)))
                _si = 0
                for _sec in _sec_ord:
                    _sdf = _conv_m[_conv_m["Section"]==_sec]
                    if _sdf.empty: continue
                    _sc = _scols[_si % len(_scols)]
                    _sc.markdown(f"**{_sec}**")
                    for metric, col in [("OEE","OEE %"),("Availability","Availability %"),
                                        ("Rate","Rate %"),("Quality","Quality %")]:
                        _v = _sdf[col].mean() if col in _sdf.columns else None
                        _sc.metric(metric, f"{_v:.1f}%" if pd.notna(_v) else "—")
                    _si += 1

                _tr = df_conv[df_conv["Row Type"]=="Total"]
                if not _tr.empty:
                    _tw1, _tw2 = st.columns(2)
                    _tw1.metric("Total FG Produced", f"{_tr['FG Produced'].iloc[0]:,.0f}")
                    _tw2.metric("Full Qty Weight",   f"{_tr['Full Qty Weight'].iloc[0]:,.0f}")

                st.markdown("#### Extracted Data")
                _disp = [c for c in _dfvv.columns if c not in
                         {"Capacity Utilization %","PM & Cleaning (label)","Micro-Stop (label)"}]
                st.dataframe(_dfvv[_disp], use_container_width=True, hide_index=True)

                CONV_COLORS = {
                    "Die Cutters":"#9B59B6","FFG":"#006394","Printer":"#C1A02E",
                    "Folder Gluers":"#D8C37D","Stitcher":"#E74C3C","Jumbo":"#27AE60",
                    "Single Facer":"#1ABC9C","Shrink Wrapper":"#E67E22","Unknown":"#AAAAAA",
                }
                if not _conv_m.empty and "OEE %" in _conv_m.columns:
                    st.markdown("#### OEE by Machine")
                    st.image(_oee_bar_chart(_conv_m, "OEE %", CONV_COLORS))
                    st.markdown("#### Availability / Rate / Quality")
                    st.image(_arq_chart(_conv_m, "Availability %","Rate %","Quality %"))

                st.markdown("#### Export")
                st.download_button("📥 Download Converting OEE (.xlsx)",
                    data=_build_excel(df_conv,"Converting OEE","9B59B6"),
                    file_name=f"Converting_OEE_{conv_file.name.replace('.xls','').replace('.xlsx','')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="fi_conv_dl")
            except Exception as e:
                st.error(f"Converting Error: {e}")
