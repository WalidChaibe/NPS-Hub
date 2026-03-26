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

tab1, = st.tabs(["📊 Corrugator OEE Report & Analysis"])

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
