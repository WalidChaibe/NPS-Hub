# utils/qm_pipeline.py
import io
import pandas as pd

SHEET_NAME = 0
CRM_COL = "Ref NB"
FINAL_APPROVAL_COL    = "Final Approval Date"
CREATION_DATETIME_COL = "Creation Date Time"
REASON_COL = "Reason"
GEN_CAT_COL = "Gen Categories"
PHYS_STATUS_COL = "Physical Status"
REASON_TYPE_COL = "Reason Type"
COST_COL = "Cost Amount"

GEN_CAT_ALLOWED = {"Customer Complaint", "Customer Return", "Process Improvement"}
PHYS_STATUS_BLOCKLIST = {"Baled Waste", "Plastic/Wood Waste"}

CRM_DELETE_MAP = {
    "EPAK-CRM-10376": 14, "EPAK-CRM-10439": 10, "EPAK-CRM-10452": 3,
    "EPAK-CRM-10514": 10, "EPAK-CRM-10573": 4,  "EPAK-CRM-10561": 8,
    "EPAK-CRM-10697": 9,  "EPAK-CRM-10698": 13, "EPAK-CRM-10699": 9,
    "EPAK-CRM-10772": 9,  "EPAK-CRM-10774": 11, "EPAK-CRM-10831": 14,
    "EPAK-CRM-10813": 3,  "EPAK-CRM-10771": 3,  "EPAK-CRM-10902": 10,
    "EPAK-CRM-10960": 10, "EPAK-CRM-10936": 3,  "EPAK-CRM-11115": 1,
    "EPAK-CRM-11124": 5,  "EPAK-CRM-11141": 8,  "EPAK-CRM-11142": 10,
    "EPAK-CRM-11143": 8,  "EPAK-CRM-11147": 13,
}

QUALITY_REASONS = {
    "Poor Ink Coverage / Pinholes","Score Cracking","Missing/ Hard Score","Delamination",
    "Warped Sheets","Belt Mark","Chemical odors","Wrong score size","Misaligned Paper",
    "Variation paper Shade","Deviation from customer flute requirement","Crushed boards",
    "Deviation from printing design","Incorrect Printing Layout","Ink rubbing",
    "Score Cracking;Cutting","Less Finished Goods Quantity",
    "Wrong Dimension; Sheet Size, Scores","Excess Quantity (Over Production)",
    "Incorrect Palletizing","Weak Glue-Lap Bond","Wrong Printing",
    "Dimension Incorrect - Printing","Damaged Wooden Pallet","Ink - Poor Coverage",
    "GSM Downgrade","Glue Quesszed Out","Scratch Marks","Weak board",
    "Poor coating/ paper quality","Paper Peel Off","Blisters / Bubbles","Wash boarding",
    "Wrinkles","Rough cut","Wet boards","Cut Misregistration",
    "Deviation from Customer approved GSM","Glue Joint Variations",
    "Poor Glue adhesion / Missing Glue","Sticky material",
    "Improper Folding at Glue Lap","Incorrect Stitching",
    "Uneven/ Black Wax application","Deviation from structural design (Die cut)","Ink Smearing",
    "Poor Die Cutting/ Hanging Trim","Hard Folding","Slotting Variation",
    "Deviation from Customer packing mode","Accumulated Damage",
    "Damaged Material/ Pallet","Wet Carton",
    "Oil, Dust & Foreign Body Contamination","Printing Mechanical Damage/ Poor legibility",
    "Ink Color Variation","Printing Miss-Registration","Wash Boarding","Missing FT Data",
}

SERVICE_REASONS = {
    "Incorrect Finished Goods Pallet Tag","Excess Quantity Produced",
    "Deviation from delivery Schedule","Incorrect Delivery Location / address",
    "Wrong Item Delivered / Invoiced","Wrong Unit Price",
    "Incorrect Sales Contract Processed","Incorrect Sales Contract Pricing",
    "Trailer Waiting Hours Exceeded","Incorrect Loading Manifest",
    "Variance in Quantity Invoiced",
}

def _require_cols(df, cols, df_name="df"):
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise KeyError(f"[{df_name}] Missing: {missing}\nAvailable: {list(df.columns)}")

def _clean_text(x) -> str:
    s = "" if pd.isna(x) else str(x)
    s = s.replace("\n", " ").replace("\r", " ").strip()
    s = " ".join(s.split())
    s = s.replace(" ;", ";").replace("; ", ";").replace(";", "; ")
    return " ".join(s.split())

QUALITY_REASONS_CLEAN = {_clean_text(r) for r in QUALITY_REASONS}
SERVICE_REASONS_CLEAN = {_clean_text(r) for r in SERVICE_REASONS}

def _classify_reason(rsn) -> str:
    rsn = _clean_text(rsn)
    if rsn == "": return "UNCLASSIFIED"
    if rsn in QUALITY_REASONS_CLEAN: return "Quality"
    if rsn in SERVICE_REASONS_CLEAN: return "Service"
    return "UNCLASSIFIED"

def apply_crm_deletions(df, crm_delete_map, crm_col=CRM_COL, df_name="df"):
    rows_to_drop = []
    crm_series = df[crm_col].astype(str).str.strip()
    for crm, n in crm_delete_map.items():
        if n <= 0: continue
        idx = df[crm_series == crm].index
        if len(idx) == 0: continue
        rows_to_drop.extend(idx[-n:].tolist())
    return df.drop(index=rows_to_drop).reset_index(drop=True)

def add_date_and_flags_final_issued(df, date_col, df_name="df"):
    df["Base_Date"] = pd.to_datetime(df[date_col], errors="coerce")
    df["Year"]  = df["Base_Date"].dt.year
    df["Month"] = df["Base_Date"].dt.month
    reason      = df[REASON_COL].map(_clean_text)
    gen_cat     = df[GEN_CAT_COL].map(_clean_text)
    phys_status = df[PHYS_STATUS_COL].map(_clean_text)
    u_raw = df[REASON_TYPE_COL]
    u_num = pd.to_numeric(u_raw, errors="coerce")
    u_str = u_raw.map(_clean_text)
    u_is_nonzero = (u_num.notna() & u_num.ne(0)) | (
        u_num.isna() & u_str.ne("") & u_str.ne("0") & u_str.ne("0.0") & u_str.str.lower().ne("nan")
    )
    df["Is_Valid"] = (
        reason.ne("Invalid") & u_is_nonzero &
        (~phys_status.isin(PHYS_STATUS_BLOCKLIST)) & gen_cat.isin(GEN_CAT_ALLOWED)
    )
    df["Complaint_Category"] = [
        "Invalid" if not v else _classify_reason(r)
        for v, r in zip(df["Is_Valid"].tolist(), reason.tolist())
    ]
    return df

def add_date_and_flags_ncr(df, date_col, df_name="NCR"):
    df["Base_Date"] = pd.to_datetime(df[date_col], errors="coerce")
    df["Year"]  = df["Base_Date"].dt.year
    df["Month"] = df["Base_Date"].dt.month
    gen_cat = df[GEN_CAT_COL].map(_clean_text)
    reason  = df[REASON_COL].map(_clean_text)
    df["Is_Valid"] = gen_cat.eq("Work In Progress")
    df["Complaint_Category"] = [
        "Invalid" if not v else _classify_reason(r)
        for v, r in zip(df["Is_Valid"].tolist(), reason.tolist())
    ]
    return df

def show_unclassified_counts(df):
    reason_clean = df[REASON_COL].map(_clean_text)
    unclassified = df[(df["Complaint_Category"] == "UNCLASSIFIED") & (reason_clean != "Invalid")].copy()
    unclassified_nonblank = unclassified[reason_clean.loc[unclassified.index] != ""]
    if unclassified_nonblank.empty: return pd.Series([], dtype="int64")
    return unclassified_nonblank[REASON_COL].map(_clean_text).value_counts()

def read_excel_from_upload(uploaded_file, sheet_name=SHEET_NAME, drop_last_two=True) -> pd.DataFrame:
    data = uploaded_file.getvalue()
    df = pd.read_excel(io.BytesIO(data), sheet_name=sheet_name)
    if drop_last_two and len(df) >= 2:
        df = df.iloc[:-2].reset_index(drop=True)
    return df

def build_dataset_final_issued(df_loaded, date_col, dataset_name="DATASET"):
    raw            = df_loaded.copy()
    raw_flagged    = add_date_and_flags_final_issued(raw.copy(), date_col=date_col, df_name=f"{dataset_name}.raw")
    cleaned        = apply_crm_deletions(df_loaded.copy(), CRM_DELETE_MAP, crm_col=CRM_COL, df_name=dataset_name)
    cleaned_flagged= add_date_and_flags_final_issued(cleaned, date_col=date_col, df_name=f"{dataset_name}.cleaned")
    return {"raw": raw, "raw_flagged": raw_flagged, "cleaned_flagged": cleaned_flagged,
            "unclassified_counts": show_unclassified_counts(cleaned_flagged)}

def build_dataset_ncr(df_loaded, date_col, dataset_name="NCR"):
    raw         = df_loaded.copy()
    raw_flagged = add_date_and_flags_ncr(raw.copy(), date_col=date_col, df_name=dataset_name)
    return {"raw": raw, "raw_flagged": raw_flagged, "cleaned_flagged": raw_flagged.copy(),
            "unclassified_counts": show_unclassified_counts(raw_flagged)}
