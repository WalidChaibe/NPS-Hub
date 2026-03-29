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

tab1, tab2 = st.tabs(["📊 Corrugator OEE Report & Analysis", "📊 Converting OEE Report & Analysis"])

# ════════════════════════════════════════
# TAB 1 — CORRUGATOR OEE
# ════════════════════════════════════════
with tab1:
    st.markdown("### 📊 Corrugator OEE Report & Analysis")

    mpl.rcParams["font.family"] = "DejaVu Sans"
    mpl.rcParams["font.size"]   = 11

    # ── Constants ──
    SECTION_KEYWORDS = {"BHS", "FOSBER", "SINGLE FACER"}

    MACHINE_FIXED_COLS = [
        "Shift Time", "Run Duration (Hrs)", "Stop Duration Hrs", "Idle Duration Hrs",
        "# of Stops", "# of Change Over", "# of Quality Change",
        "Linear Meters", "Standard Output M", "Square Meters",
        "Availability (%)", "Rate (%)", "Quality (%)", "OEE (%)", "Avg Max Speed"
    ]
    TOTAL_FIXED_COLS = [
        "Shift Time", "Run Duration (Hrs)", "Stop Duration Hrs", "Idle Duration Hrs",
        "# of Stops", "# of Change Over", "# of Quality Change",
        "Linear Meters", "Standard Output M", "Square Meters",
        "Availability (%)", "Rate (%)", "Quality (%)", "OEE (%)"
    ]

    def parse_oee_file(file_bytes) -> pd.DataFrame:
        wb  = xlrd.open_workbook(file_contents=file_bytes)
        sh  = wb.sheet_by_index(0)
        current_section = None
        all_rows = []

        for row_idx in range(sh.nrows):
            row   = [sh.cell_value(row_idx, col) for col in range(sh.ncols)]
            col_a = str(row[0]).strip()

            if col_a in SECTION_KEYWORDS:
                current_section = col_a
                continue

            if not any(str(v).strip() for v in row):
                continue

            # Machine row
            if isinstance(row[0], str) and col_a and col_a not in SECTION_KEYWORDS:
                record = {"Section": current_section, "Machine": col_a, "Row Type": "Machine"}
                for i, col_name in enumerate(MACHINE_FIXED_COLS):
                    record[col_name] = row[i + 1] if (i + 1) < len(row) else None
                record["Available Time (Hrs)"] = row[15] if 15 < len(row) else None
                record["Average Length/Run"]   = row[17] if 17 < len(row) else None
                record["Average SQM/Run"]      = row[19] if 19 < len(row) else None
                record["PM & Cleaning"]        = row[21] if 21 < len(row) else None
                record["MT"]                   = row[22] if 22 < len(row) else None
                record["Speed Value"]          = row[23] if 23 < len(row) else None
                record["Average Width/Run"]    = row[25] if 25 < len(row) else None
                record["Average SQM/Hrs"]      = row[27] if 27 < len(row) else None
                record["Operator"]             = row[28] if 28 < len(row) else None
                record["Average Sheet Length"] = row[30] if 30 < len(row) else None
                record["Average Sheet Width"]  = row[32] if 32 < len(row) else None
                all_rows.append(record)

            # Total row
            elif isinstance(row[0], (int, float)) and col_a:
                record = {"Section": current_section, "Machine": "TOTAL", "Row Type": "Total"}
                for i, col_name in enumerate(TOTAL_FIXED_COLS):
                    record[col_name] = row[i] if i < len(row) else None
                record["Available Time (Hrs)"]     = row[15] if 15 < len(row) else None
                record["PM & Cleaning"]            = row[17] if 17 < len(row) else None
                record["Average Length/Run"]       = row[19] if 19 < len(row) else None
                record["Average SQM/Run"]          = row[21] if 21 < len(row) else None
                record["MT"]                       = row[22] if 22 < len(row) else None
                record["Average Width/Run"]        = row[24] if 24 < len(row) else None
                record["Average SQM/Hrs"]          = row[26] if 26 < len(row) else None
                record["Speed Value"]              = row[27] if 27 < len(row) else None
                record["# of ProgramID"]           = row[29] if 29 < len(row) else None
                record["Average Sheet Length"]     = row[31] if 31 < len(row) else None
                record["Average Sheet Width"]      = row[33] if 33 < len(row) else None
                record["# of Change Overs"]        = row[35] if 35 < len(row) else None
                record["Capacity Utilization (%)"] = row[37] if 37 < len(row) else None
                record["# of Quality Changes"]     = row[39] if 39 < len(row) else None
                record["% of Quality Changes"]     = row[41] if 41 < len(row) else None
                record["Planning Efficiency"]      = row[43] if 43 < len(row) else None
                all_rows.append(record)

        df = pd.DataFrame(all_rows)
        skip_cols = {"Section", "Machine", "Row Type", "Operator"}
        for col in df.columns:
            if col not in skip_cols:
                df[col] = pd.to_numeric(df[col], errors="coerce")
        return df

    def build_excel_output(df: pd.DataFrame) -> io.BytesIO:
        wb  = Workbook()
        ws  = wb.active
        ws.title = "OEE Report"

        # Styles
        header_fill    = PatternFill("solid", fgColor="006394")
        header_font    = Font(bold=True, color="FFFFFF", size=10)
        section_fill   = PatternFill("solid", fgColor="D8C37D")
        section_font   = Font(bold=True, size=10)
        total_fill     = PatternFill("solid", fgColor="EAF4FB")
        total_font     = Font(bold=True, size=9)
        normal_font    = Font(size=9)
        center_align   = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left_align     = Alignment(horizontal="left",   vertical="center")
        thin_border    = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"),  bottom=Side(style="thin")
        )

        # Headers
        headers = list(df.columns)
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.fill      = header_fill
            cell.font      = header_font
            cell.alignment = center_align
            cell.border    = thin_border

        ws.row_dimensions[1].height = 30

        # Data rows
        current_section = None
        data_row = 2
        for _, row in df.iterrows():
            # Section divider row
            if row.get("Section") != current_section:
                current_section = row.get("Section")
                if current_section:
                    cell = ws.cell(row=data_row, column=1, value=f"── {current_section} ──")
                    cell.fill      = section_fill
                    cell.font      = section_font
                    cell.alignment = left_align
                    ws.merge_cells(start_row=data_row, start_column=1,
                                   end_row=data_row, end_column=len(headers))
                    ws.row_dimensions[data_row].height = 18
                    data_row += 1

            is_total = row.get("Row Type") == "Total"
            for col_idx, h in enumerate(headers, 1):
                val  = row.get(h)
                cell = ws.cell(row=data_row, column=col_idx, value=val)
                cell.border    = thin_border
                cell.alignment = center_align if col_idx > 2 else left_align
                if is_total:
                    cell.fill = total_fill
                    cell.font = total_font
                else:
                    cell.font = normal_font

                # Format percentages
                if isinstance(val, float) and h.endswith("(%)"):
                    cell.number_format = "0.00%"
                    ws.cell(row=data_row, column=col_idx).value = val / 100 if val and val > 1 else val

            ws.row_dimensions[data_row].height = 16
            data_row += 1

        # Column widths
        for col_idx, h in enumerate(headers, 1):
            ws.column_dimensions[get_column_letter(col_idx)].width = max(12, len(str(h)) + 2)
        ws.column_dimensions["A"].width = 10
        ws.column_dimensions["B"].width = 22
        ws.column_dimensions["C"].width = 10

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    # ── File Upload ──
    oee_file = st.file_uploader(
        "📂 Upload Corrugator OEE Report (.xls / .xlsx)",
        type=["xls", "xlsx"], key="fi_oee_file"
    )

    if not oee_file:
        st.info("⬆️ Upload the OEE report file to begin analysis.")
    else:
        try:
            with st.spinner("Parsing OEE file..."):
                _file_bytes = oee_file.getvalue()
                df_oee = parse_oee_file(_file_bytes)

            st.success(f"✅ Extracted {len(df_oee)} rows, {len(df_oee.columns)} columns")

            # ── Section filter ──
            _sections = ["All"] + sorted(df_oee["Section"].dropna().unique().tolist())
            _col1, _col2 = st.columns([2, 5])
            _sel_section = _col1.selectbox("Filter by Section", _sections, key="fi_section_filter")
            _row_types   = _col2.multiselect("Row Type", ["Machine", "Total"],
                                              default=["Machine", "Total"], key="fi_row_type")

            _df_view = df_oee.copy()
            if _sel_section != "All":
                _df_view = _df_view[_df_view["Section"] == _sel_section]
            if _row_types:
                _df_view = _df_view[_df_view["Row Type"].isin(_row_types)]

            # ── OEE Summary metrics ──
            st.divider()
            st.markdown("#### OEE Summary")
            _totals = df_oee[df_oee["Row Type"] == "Total"]
            if not _totals.empty:
                _mc = st.columns(len(_totals) + 1)
                for i, (_, trow) in enumerate(_totals.iterrows()):
                    _sec = trow.get("Section","")
                    _oee = trow.get("OEE (%)", 0) or 0
                    _avl = trow.get("Availability (%)", 0) or 0
                    _rte = trow.get("Rate (%)", 0) or 0
                    _qlt = trow.get("Quality (%)", 0) or 0
                    _mc[i].markdown(f"**{_sec}**")
                    _mc[i].metric("OEE",          f"{_oee:.1f}%")
                    _mc[i].metric("Availability", f"{_avl:.1f}%")
                    _mc[i].metric("Rate",         f"{_rte:.1f}%")
                    _mc[i].metric("Quality",      f"{_qlt:.1f}%")

            # ── Raw data table ──
            st.divider()
            st.markdown("#### Extracted Data")
            st.dataframe(_df_view, use_container_width=True, hide_index=True)

            # ── OEE Bar Chart by Machine ──
            _machines = df_oee[df_oee["Row Type"] == "Machine"].copy()
            if not _machines.empty and "OEE (%)" in _machines.columns:
                st.divider()
                st.markdown("#### OEE by Machine")
                _fig_oee, _ax_oee = plt.subplots(figsize=(13.33, 5), dpi=150)
                _x = np.arange(len(_machines))
                _colors = ["#006394" if s == "BHS" else "#C1A02E" if s == "FOSBER"
                           else "#D8C37D" for s in _machines["Section"].fillna("")]
                _bars = _ax_oee.bar(_x, _machines["OEE (%)"].fillna(0), color=_colors,
                                    width=0.6, alpha=0.9)
                for _bar, _val in zip(_bars, _machines["OEE (%)"].fillna(0)):
                    _ax_oee.text(_bar.get_x()+_bar.get_width()/2,
                                 _bar.get_height()+0.5,
                                 f"{_val:.1f}%", ha="center", va="bottom",
                                 fontsize=9, color="#000000")
                _ax_oee.axhline(85, color="#DE201B", linewidth=1.5,
                                linestyle="--", label="Target 85%")
                _ax_oee.set_xticks(_x)
                _ax_oee.set_xticklabels(_machines["Machine"].tolist(),
                                        rotation=30, ha="right", fontsize=9)
                _ax_oee.set_ylabel("OEE (%)")
                _ax_oee.set_ylim(0, 110)
                _ax_oee.spines["top"].set_visible(False)
                _ax_oee.spines["right"].set_visible(False)
                _ax_oee.grid(False)
                _ax_oee.legend(frameon=False)

                # Color legend for sections
                from matplotlib.patches import Patch
                _legend_els = [
                    Patch(facecolor="#006394", label="BHS"),
                    Patch(facecolor="#C1A02E", label="FOSBER"),
                    Patch(facecolor="#D8C37D", label="SINGLE FACER"),
                ]
                _ax_oee.legend(handles=_legend_els + [
                    plt.Line2D([0],[0], color="#DE201B", linewidth=1.5,
                               linestyle="--", label="Target 85%")],
                    loc="upper right", frameon=False, fontsize=9)
                plt.tight_layout()
                _buf_oee = io.BytesIO()
                _fig_oee.savefig(_buf_oee, format="png", dpi=150, bbox_inches="tight")
                _buf_oee.seek(0)
                st.image(_buf_oee)
                plt.close(_fig_oee)

            # ── Availability / Rate / Quality breakdown ──
            if not _machines.empty:
                st.divider()
                st.markdown("#### Availability / Rate / Quality by Machine")
                _fig_arq, _ax_arq = plt.subplots(figsize=(13.33, 5), dpi=150)
                _xm = np.arange(len(_machines))
                _w  = 0.25
                _b1 = _ax_arq.bar(_xm - _w, _machines["Availability (%)"].fillna(0),
                                   _w, color="#006394", label="Availability")
                _b2 = _ax_arq.bar(_xm,       _machines["Rate (%)"].fillna(0),
                                   _w, color="#C1A02E", label="Rate")
                _b3 = _ax_arq.bar(_xm + _w, _machines["Quality (%)"].fillna(0),
                                   _w, color="#D8C37D", label="Quality")
                for _b in [_b1, _b2, _b3]:
                    for _bar in _b:
                        _h = _bar.get_height()
                        if _h > 0:
                            _ax_arq.text(_bar.get_x()+_bar.get_width()/2,
                                         _h+0.5, f"{_h:.1f}%",
                                         ha="center", va="bottom", fontsize=8, color="#000000")
                _ax_arq.set_xticks(_xm)
                _ax_arq.set_xticklabels(_machines["Machine"].tolist(),
                                        rotation=30, ha="right", fontsize=9)
                _ax_arq.set_ylabel("%")
                _ax_arq.set_ylim(0, 115)
                _ax_arq.spines["top"].set_visible(False)
                _ax_arq.spines["right"].set_visible(False)
                _ax_arq.grid(False)
                _ax_arq.legend(loc="upper center", bbox_to_anchor=(0.5, -0.18),
                               ncol=3, frameon=False, fontsize=10)
                plt.tight_layout(rect=[0, 0.08, 1, 1])
                _buf_arq = io.BytesIO()
                _fig_arq.savefig(_buf_arq, format="png", dpi=150, bbox_inches="tight")
                _buf_arq.seek(0)
                st.image(_buf_arq)
                plt.close(_fig_arq)

            # ── Excel Download ──
            st.divider()
            st.markdown("#### Export")
            _excel_buf = build_excel_output(df_oee)
            st.download_button(
                label="📥 Download Extracted OEE Report (.xlsx)",
                data=_excel_buf,
                file_name=f"OEE_Report_{oee_file.name.replace('.xls','').replace('.xlsx','')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="fi_oee_download"
            )

        except Exception as e:
            st.error(f"Error parsing file: {e}")

# ════════════════════════════════════════
# TAB 2 — CONVERTING OEE
# ════════════════════════════════════════
with tab2:
    st.markdown("### 📊 Converting OEE Report & Analysis")

    # ── Machine → Section mapping ──
    CONVERTING_SECTION_MAP = {
        "BOBST 160-II":        "Die Cutters",
        "BOBST 203":           "Die Cutters",
        "BOBST MASTERCUT 1":   "Die Cutters",
        "BOBST MASTERCUT 2":   "Die Cutters",
        "BOBST 924":           "FFG",
        "LMC FFG":             "FFG",
        "MARTIN 616":          "FFG",
        "SATURN":              "FFG",
        "IPACK":               "Printer",
        "VEGA 2":              "Folder Gluers",
        "BAHMÜLLER TURBOX":    "Folder Gluers",
        "BAHMULLER STITCHER":  "Stitcher",
        "JUMBO":               "Jumbo",
        "SINGLE FACER":        "Single Facer",
        "SHRINK-WRAPPER":      "Shrink Wrapper",
    }
    # Normalised lookup (upper strip)
    CONVERTING_SECTION_MAP_NORM = {k.upper().strip(): v for k, v in CONVERTING_SECTION_MAP.items()}

    CONVERTING_COLS = [
        "Machine Name",
        "Hits / Hour",
        "Standard Hits (Rhr)",
        "Capacity Utilization %",
        "Capacity Utilization Value",
        "Shift Time (Hrs)",
        "Available Time (Hrs)",
        "Run Duration (Hrs)",
        "STD Setup (min)",
        "Setup Duration (min)",
        "Stop Duration (Hrs)",
        "Idle Duration (Hrs)",
        "# of Setup",
        "# of Orders",
        "SQM",
        "Actual Hits",
        "Hits / Avl Hour",
        "Hits Rhr+Setup",
        "Availability %",
        "Rate %",
        "Quality %",
        "OEE %",
        "FG Produced",
        "PM & Cleaning (label)",
        "PM & Cleaning Time",
        "Micro-Stop (label)",
        "Micro Stop (min)",
        "Unknown Value",
        "Full Qty Weight",
    ]

    def parse_converting_oee(file_bytes, filename=""):
        # Support both .xls and .xlsx
        if filename.lower().endswith(".xlsx"):
            import openpyxl
            _wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
            _ws = _wb.active
            rows_raw = []
            for row in _ws.iter_rows(values_only=True):
                rows_raw.append([c if c is not None else "" for c in row])
        else:
            import xlrd as _xlrd
            _wb = _xlrd.open_workbook(file_contents=file_bytes)
            _ws = _wb.sheet_by_index(0)
            rows_raw = []
            for r in range(_ws.nrows):
                rows_raw.append([_ws.cell_value(r, c) for c in range(_ws.ncols)])

        all_rows = []
        # Skip first 2 header rows
        for row in rows_raw[2:]:
            col_a = str(row[0]).strip() if row else ""

            # Stop conditions
            if not col_a:
                continue
            if "total weight" in col_a.lower():
                # Capture total weight row
                _total_weight     = row[1] if len(row) > 1 else None
                _full_qty_weight  = row[2] if len(row) > 2 else None
                all_rows.append({
                    "Section":      "TOTAL",
                    "Machine Name": "TOTAL",
                    "Row Type":     "Total",
                    "FG Produced":  pd.to_numeric(_total_weight, errors="coerce"),
                    "Full Qty Weight": pd.to_numeric(_full_qty_weight, errors="coerce"),
                })
                break
            if col_a.lower().startswith("page"):
                continue

            # Check if machine name is known
            _section = CONVERTING_SECTION_MAP_NORM.get(col_a.upper())

            # Check if row has meaningful data (at least 5 non-empty cells after col A)
            _data_cells = [v for v in row[1:] if str(v).strip() not in ("", "0", "0.0")]
            if not _data_cells or len(_data_cells) < 3:
                continue  # Empty/name-only row — skip

            if _section is None:
                # Unknown machine but has data — still include with section "Unknown"
                _section = "Unknown"

            record = {"Section": _section, "Machine Name": col_a, "Row Type": "Machine"}

            # Map columns by position
            col_map = {
                "Hits / Hour":              1,
                "Standard Hits (Rhr)":      2,
                "Capacity Utilization %":   3,  # label cell
                "Capacity Utilization Value": 4,
                "Shift Time (Hrs)":         5,
                "Available Time (Hrs)":     6,
                "Run Duration (Hrs)":       7,
                "STD Setup (min)":          8,
                "Setup Duration (min)":     9,
                "Stop Duration (Hrs)":      10,
                "Idle Duration (Hrs)":      11,
                "# of Setup":              12,
                "# of Orders":             13,
                "SQM":                     14,
                "Actual Hits":             15,
                "Hits / Avl Hour":         16,
                "Hits Rhr+Setup":          17,
                "Availability %":          18,
                "Rate %":                  19,
                "Quality %":               20,
                "OEE %":                   21,
                "FG Produced":             22,
                "PM & Cleaning (label)":   23,
                "PM & Cleaning Time":      24,
                "Micro-Stop (label)":      25,
                "Micro Stop (min)":        26,
                "Unknown Value":           27,
                "Full Qty Weight":         28,
            }
            for col_name, idx in col_map.items():
                record[col_name] = row[idx] if idx < len(row) else None

            all_rows.append(record)

        df = pd.DataFrame(all_rows)
        skip_cols = {"Section", "Machine Name", "Row Type",
                     "Capacity Utilization %", "PM & Cleaning (label)", "Micro-Stop (label)"}
        for col in df.columns:
            if col not in skip_cols:
                df[col] = pd.to_numeric(df[col], errors="coerce")
        return df

    def build_converting_excel(df: pd.DataFrame) -> io.BytesIO:
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        wb = Workbook(); ws = wb.active; ws.title = "Converting OEE"

        header_fill  = PatternFill("solid", fgColor="9B59B6")
        header_font  = Font(bold=True, color="FFFFFF", size=10)
        section_fill = PatternFill("solid", fgColor="E8DAEF")
        section_font = Font(bold=True, size=10)
        total_fill   = PatternFill("solid", fgColor="EAF4FB")
        total_font   = Font(bold=True, size=9)
        normal_font  = Font(size=9)
        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left_align   = Alignment(horizontal="left",   vertical="center")
        thin = Border(left=Side(style="thin"), right=Side(style="thin"),
                      top=Side(style="thin"),  bottom=Side(style="thin"))

        headers = [c for c in df.columns if c not in
                   {"Capacity Utilization %", "PM & Cleaning (label)", "Micro-Stop (label)"}]
        for ci, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=ci, value=h)
            cell.fill = header_fill; cell.font = header_font
            cell.alignment = center_align; cell.border = thin
        ws.row_dimensions[1].height = 30

        data_row = 2; current_section = None
        for _, row in df.iterrows():
            sec = row.get("Section","")
            if sec != current_section and sec not in ("TOTAL","Unknown"):
                current_section = sec
                cell = ws.cell(row=data_row, column=1, value=f"── {sec} ──")
                cell.fill = section_fill; cell.font = section_font
                cell.alignment = left_align
                ws.merge_cells(start_row=data_row, start_column=1,
                               end_row=data_row, end_column=len(headers))
                ws.row_dimensions[data_row].height = 18
                data_row += 1

            is_total = row.get("Row Type") == "Total"
            for ci, h in enumerate(headers, 1):
                val  = row.get(h)
                cell = ws.cell(row=data_row, column=ci, value=val)
                cell.border = thin
                cell.alignment = center_align if ci > 2 else left_align
                cell.fill = total_fill if is_total else PatternFill()
                cell.font = total_font if is_total else normal_font
            ws.row_dimensions[data_row].height = 16
            data_row += 1

        for ci, h in enumerate(headers, 1):
            ws.column_dimensions[get_column_letter(ci)].width = max(12, len(str(h)) + 2)
        ws.column_dimensions["A"].width = 10
        ws.column_dimensions["B"].width = 24

        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        return buf

    # ── File Upload ──
    conv_file = st.file_uploader(
        "📂 Upload Converting OEE Report (.xls / .xlsx)",
        type=["xls","xlsx"], key="fi_conv_file"
    )

    if not conv_file:
        st.info("⬆️ Upload the Converting OEE report file to begin analysis.")
    else:
        try:
            with st.spinner("Parsing Converting OEE file..."):
                _conv_bytes = conv_file.getvalue()
                df_conv = parse_converting_oee(_conv_bytes, conv_file.name)

            _machine_rows = df_conv[df_conv["Row Type"]=="Machine"]
            st.success(f"✅ Extracted {len(_machine_rows)} machine rows, {len(df_conv.columns)} columns")

            # ── Section filter ──
            _sections_conv = ["All"] + sorted(
                [s for s in df_conv["Section"].dropna().unique() if s not in ("TOTAL","Unknown")]
            )
            _c1, _c2 = st.columns([2,5])
            _sel_sec_conv = _c1.selectbox("Filter by Section", _sections_conv, key="fi_conv_section")

            _df_conv_view = _machine_rows.copy()
            if _sel_sec_conv != "All":
                _df_conv_view = _df_conv_view[_df_conv_view["Section"]==_sel_sec_conv]

            # ── Summary metrics ──
            st.divider()
            st.markdown("#### OEE Summary by Section")
            _section_order = ["Die Cutters","FFG","Printer","Folder Gluers",
                              "Stitcher","Jumbo","Single Facer","Shrink Wrapper"]
            _sum_cols = st.columns(min(4, len(_section_order)))
            _col_idx  = 0
            for _sec in _section_order:
                _sec_df = _machine_rows[_machine_rows["Section"]==_sec]
                if _sec_df.empty: continue
                _avg_oee  = _sec_df["OEE %"].mean()
                _avg_avl  = _sec_df["Availability %"].mean()
                _avg_rate = _sec_df["Rate %"].mean()
                _avg_qlt  = _sec_df["Quality %"].mean()
                _col = _sum_cols[_col_idx % len(_sum_cols)]
                _col.markdown(f"**{_sec}**")
                _col.metric("OEE",          f"{_avg_oee:.1f}%" if pd.notna(_avg_oee) else "—")
                _col.metric("Availability", f"{_avg_avl:.1f}%" if pd.notna(_avg_avl) else "—")
                _col.metric("Rate",         f"{_avg_rate:.1f}%" if pd.notna(_avg_rate) else "—")
                _col.metric("Quality",      f"{_avg_qlt:.1f}%" if pd.notna(_avg_qlt) else "—")
                _col_idx += 1

            # ── Total weight metrics ──
            _total_row = df_conv[df_conv["Row Type"]=="Total"]
            if not _total_row.empty:
                st.divider()
                _tw1, _tw2 = st.columns(2)
                _tw1.metric("Total FG Produced (Weight)", f"{_total_row['FG Produced'].iloc[0]:,.0f}")
                _tw2.metric("Full Qty Weight (All Machines)", f"{_total_row['Full Qty Weight'].iloc[0]:,.0f}")

            # ── Raw table ──
            st.divider()
            st.markdown("#### Extracted Data")
            _display_cols = [c for c in _df_conv_view.columns
                             if c not in {"Capacity Utilization %","PM & Cleaning (label)","Micro-Stop (label)"}]
            st.dataframe(_df_conv_view[_display_cols], use_container_width=True, hide_index=True)

            # ── OEE Bar Chart ──
            if not _machine_rows.empty and "OEE %" in _machine_rows.columns:
                st.divider()
                st.markdown("#### OEE by Machine")
                _section_colors = {
                    "Die Cutters":   "#9B59B6",
                    "FFG":           "#006394",
                    "Printer":       "#C1A02E",
                    "Folder Gluers": "#D8C37D",
                    "Stitcher":      "#E74C3C",
                    "Jumbo":         "#27AE60",
                    "Single Facer":  "#1ABC9C",
                    "Shrink Wrapper":"#E67E22",
                    "Unknown":       "#AAAAAA",
                }
                _fig_c, _ax_c = plt.subplots(figsize=(13.33, 5), dpi=150)
                _xc = np.arange(len(_machine_rows))
                _colors_c = [_section_colors.get(s, "#AAAAAA") for s in _machine_rows["Section"].fillna("")]
                _bars_c = _ax_c.bar(_xc, _machine_rows["OEE %"].fillna(0),
                                    color=_colors_c, width=0.6, alpha=0.9)
                for _bar, _val in zip(_bars_c, _machine_rows["OEE %"].fillna(0)):
                    _ax_c.text(_bar.get_x()+_bar.get_width()/2,
                               _bar.get_height()+0.5, f"{_val:.1f}%",
                               ha="center", va="bottom", fontsize=9, color="#000000")
                _ax_c.axhline(85, color="#DE201B", linewidth=1.5,
                              linestyle="--", label="Target 85%")
                _ax_c.set_xticks(_xc)
                _ax_c.set_xticklabels(_machine_rows["Machine Name"].tolist(),
                                      rotation=35, ha="right", fontsize=9)
                _ax_c.set_ylabel("OEE (%)"); _ax_c.set_ylim(0, 115)
                _ax_c.spines["top"].set_visible(False)
                _ax_c.spines["right"].set_visible(False)
                _ax_c.grid(False)
                from matplotlib.patches import Patch
                _legend_patches = [Patch(facecolor=v, label=k)
                                   for k, v in _section_colors.items()
                                   if k in _machine_rows["Section"].values]
                _legend_patches.append(
                    plt.Line2D([0],[0], color="#DE201B", linewidth=1.5,
                               linestyle="--", label="Target 85%")
                )
                _ax_c.legend(handles=_legend_patches, loc="upper center",
                             bbox_to_anchor=(0.5,-0.22), ncol=5, frameon=False, fontsize=9)
                plt.tight_layout(rect=[0,0.10,1,1])
                _buf_c = io.BytesIO()
                _fig_c.savefig(_buf_c, format="png", dpi=150, bbox_inches="tight")
                _buf_c.seek(0); st.image(_buf_c); plt.close(_fig_c)

            # ── Availability / Rate / Quality ──
            if not _machine_rows.empty:
                st.divider()
                st.markdown("#### Availability / Rate / Quality by Machine")
                _fig_arq2, _ax_arq2 = plt.subplots(figsize=(13.33, 5), dpi=150)
                _xm2 = np.arange(len(_machine_rows)); _w2 = 0.25
                _b1 = _ax_arq2.bar(_xm2-_w2, _machine_rows["Availability %"].fillna(0),
                                   _w2, color="#006394", label="Availability")
                _b2 = _ax_arq2.bar(_xm2,     _machine_rows["Rate %"].fillna(0),
                                   _w2, color="#C1A02E", label="Rate")
                _b3 = _ax_arq2.bar(_xm2+_w2, _machine_rows["Quality %"].fillna(0),
                                   _w2, color="#D8C37D", label="Quality")
                for _b in [_b1,_b2,_b3]:
                    for _bar in _b:
                        _h = _bar.get_height()
                        if _h > 0:
                            _ax_arq2.text(_bar.get_x()+_bar.get_width()/2,
                                          _h+0.5, f"{_h:.1f}%",
                                          ha="center", va="bottom", fontsize=8, color="#000000")
                _ax_arq2.set_xticks(_xm2)
                _ax_arq2.set_xticklabels(_machine_rows["Machine Name"].tolist(),
                                         rotation=35, ha="right", fontsize=9)
                _ax_arq2.set_ylabel("%"); _ax_arq2.set_ylim(0,115)
                _ax_arq2.spines["top"].set_visible(False)
                _ax_arq2.spines["right"].set_visible(False)
                _ax_arq2.grid(False)
                _ax_arq2.legend(loc="upper center", bbox_to_anchor=(0.5,-0.22),
                                ncol=3, frameon=False, fontsize=10)
                plt.tight_layout(rect=[0,0.10,1,1])
                _buf_arq2 = io.BytesIO()
                _fig_arq2.savefig(_buf_arq2, format="png", dpi=150, bbox_inches="tight")
                _buf_arq2.seek(0); st.image(_buf_arq2); plt.close(_fig_arq2)

            # ── Excel Download ──
            st.divider()
            st.markdown("#### Export")
            _excel_conv = build_converting_excel(df_conv)
            st.download_button(
                label="📥 Download Extracted Converting OEE Report (.xlsx)",
                data=_excel_conv,
                file_name=f"Converting_OEE_{conv_file.name.replace('.xls','').replace('.xlsx','')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="fi_conv_download"
            )

        except Exception as e:
            st.error(f"Error parsing file: {e}")
