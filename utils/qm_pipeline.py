# qm_pipeline.py
import io
import pandas as pd

SHEET_NAME            = 0
CRM_COL               = "Ref NB"
FINAL_APPROVAL_COL    = "Final Approval Date"
CREATION_DATETIME_COL = "Creation Date Time"
REASON_COL            = "Reason"
GEN_CAT_COL           = "Gen Categories"
PHYS_STATUS_COL       = "Physical Status"
REASON_TYPE_COL       = "Reason Type"
COST_COL              = "Cost Amount"

GEN_CAT_ALLOWED       = {"Customer Complaint", "Customer Return", "Process Improvement"}
PHYS_STATUS_BLOCKLIST = {"Baled Waste", "Plastic/Wood Waste"}

_CRM_DELETE_MAP_DEFAULT = {
    "EPAK-CRM-10376": 14, "EPAK-CRM-10439": 10, "EPAK-CRM-10452": 3,
    "EPAK-CRM-10514": 10, "EPAK-CRM-10573": 4,  "EPAK-CRM-10561": 8,
    "EPAK-CRM-10697": 9,  "EPAK-CRM-10698": 13, "EPAK-CRM-10699": 9,
    "EPAK-CRM-10772": 9,  "EPAK-CRM-10774": 11, "EPAK-CRM-10831": 14,
    "EPAK-CRM-10813": 3,  "EPAK-CRM-10771": 3,  "EPAK-CRM-10902": 10,
    "EPAK-CRM-10960": 10, "EPAK-CRM-10936": 3,  "EPAK-CRM-11115": 1,
    "EPAK-CRM-11124": 5,  "EPAK-CRM-11141": 8,  "EPAK-CRM-11142": 10,
    "EPAK-CRM-11143": 8,  "EPAK-CRM-11147": 13,
}

_QUALITY_REASONS_DEFAULT = {
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

_SERVICE_REASONS_DEFAULT = {
    "Incorrect Finished Goods Pallet Tag","Excess Quantity Produced",
    "Deviation from delivery Schedule","Incorrect Delivery Location / address",
    "Wrong Item Delivered / Invoiced","Wrong Unit Price",
    "Incorrect Sales Contract Processed","Incorrect Sales Contract Pricing",
    "Trailer Waiting Hours Exceeded","Incorrect Loading Manifest",
    "Variance in Quantity Invoiced",
}

def load_settings_from_supabase(supabase=None):
    if supabase is None:
        return _CRM_DELETE_MAP_DEFAULT, _QUALITY_REASONS_DEFAULT, _SERVICE_REASONS_DEFAULT, set()
    try:
        crm_rows = supabase.table("qm_crm_delete_map").select("crm_ref,kept_count").execute()
        crm_map  = {r["crm_ref"]: r["kept_count"] for r in (crm_rows.data or [])}
        if not crm_map: crm_map = _CRM_DELETE_MAP_DEFAULT
    except Exception:
        crm_map = _CRM_DELETE_MAP_DEFAULT
    try:
        q_rows = supabase.table("qm_quality_reasons").select("reason").execute()
        q_set  = {r["reason"] for r in (q_rows.data or [])}
        if not q_set: q_set = _QUALITY_REASONS_DEFAULT
    except Exception:
        q_set = _QUALITY_REASONS_DEFAULT
    try:
        s_rows = supabase.table("qm_service_reasons").select("reason").execute()
        s_set  = {r["reason"] for r in (s_rows.data or [])}
        if not s_set: s_set = _SERVICE_REASONS_DEFAULT
    except Exception:
        s_set = _SERVICE_REASONS_DEFAULT
    try:
        i_rows = supabase.table("qm_invalid_reasons").select("reason").execute()
        i_set  = {r["reason"] for r in (i_rows.data or [])}
    except Exception:
        i_set = set()
    return crm_map, q_set, s_set, i_set

def _clean_text(x) -> str:
    s = "" if pd.isna(x) else str(x)
    s = s.replace("\n", " ").replace("\r", " ").strip()
    s = " ".join(s.split())
    s = s.replace(" ;", ";").replace("; ", ";").replace(";", "; ")
    return " ".join(s.split())

def make_classifier(quality_set, service_set, invalid_set):
    q_clean = {_clean_text(r) for r in quality_set}
    s_clean = {_clean_text(r) for r in service_set}
    i_clean = {_clean_text(r) for r in invalid_set}
    def _classify(rsn) -> str:
        rsn = _clean_text(rsn)
        if rsn == "": return "UNCLASSIFIED"
        if rsn in i_clean: return "Invalid"
        if rsn in q_clean: return "Quality"
        if rsn in s_clean: return "Service"
        return "UNCLASSIFIED"
    return _classify

_classify_reason_default = make_classifier(_QUALITY_REASONS_DEFAULT, _SERVICE_REASONS_DEFAULT, set())

def apply_crm_deletions(df, crm_delete_map, crm_col=CRM_COL, df_name="df"):
    rows_to_drop = []
    crm_series = df[crm_col].astype(str).str.strip()
    for crm, n in crm_delete_map.items():
        if n <= 0: continue
        idx = df[crm_series == crm].index
        if len(idx) == 0: continue
        rows_to_drop.extend(idx[-n:].tolist())
    return df.drop(index=rows_to_drop).reset_index(drop=True)

def add_date_and_flags_final_issued(df, date_col, df_name="df", classifier=None):
    if classifier is None: classifier = _classify_reason_default
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
        "Invalid" if not v else classifier(r)
        for v, r in zip(df["Is_Valid"].tolist(), reason.tolist())
    ]
    return df

def add_date_and_flags_ncr(df, date_col, df_name="NCR", classifier=None):
    if classifier is None: classifier = _classify_reason_default
    df["Base_Date"] = pd.to_datetime(df[date_col], errors="coerce")
    df["Year"]  = df["Base_Date"].dt.year
    df["Month"] = df["Base_Date"].dt.month
    gen_cat = df[GEN_CAT_COL].map(_clean_text)
    reason  = df[REASON_COL].map(_clean_text)
    df["Is_Valid"] = gen_cat.eq("Work In Progress")
    df["Complaint_Category"] = [
        "Invalid" if not v else classifier(r)
        for v, r in zip(df["Is_Valid"].tolist(), reason.tolist())
    ]
    return df

def show_unclassified_counts(df):
    reason_clean = df[REASON_COL].map(_clean_text)
    unclassified = df[
        (df["Complaint_Category"] == "UNCLASSIFIED") & (reason_clean != "Invalid")
    ].copy()
    unclassified_nonblank = unclassified[reason_clean.loc[unclassified.index] != ""]
    if unclassified_nonblank.empty: return pd.Series([], dtype="int64")
    return unclassified_nonblank[REASON_COL].map(_clean_text).value_counts()

def read_excel_from_upload(uploaded_file, sheet_name=SHEET_NAME, drop_last_two=True) -> pd.DataFrame:
    data = uploaded_file.getvalue()
    df = pd.read_excel(io.BytesIO(data), sheet_name=sheet_name)
    if drop_last_two and len(df) >= 2:
        df = df.iloc[:-2].reset_index(drop=True)
    return df

def build_dataset_final_issued(df_loaded, date_col, dataset_name="DATASET",
                                crm_delete_map=None, classifier=None):
    if crm_delete_map is None: crm_delete_map = _CRM_DELETE_MAP_DEFAULT
    if classifier is None: classifier = _classify_reason_default
    raw             = df_loaded.copy()
    raw_flagged     = add_date_and_flags_final_issued(raw.copy(), date_col=date_col,
                                                       df_name=f"{dataset_name}.raw", classifier=classifier)
    cleaned         = apply_crm_deletions(df_loaded.copy(), crm_delete_map, crm_col=CRM_COL, df_name=dataset_name)
    cleaned_flagged = add_date_and_flags_final_issued(cleaned, date_col=date_col,
                                                       df_name=f"{dataset_name}.cleaned", classifier=classifier)
    return {
        "raw": raw, "raw_flagged": raw_flagged, "cleaned_flagged": cleaned_flagged,
        "unclassified_counts": show_unclassified_counts(cleaned_flagged),
    }

def build_dataset_ncr(df_loaded, date_col, dataset_name="NCR", classifier=None):
    if classifier is None: classifier = _classify_reason_default
    raw         = df_loaded.copy()
    raw_flagged = add_date_and_flags_ncr(raw.copy(), date_col=date_col,
                                          df_name=dataset_name, classifier=classifier)
    return {
        "raw": raw, "raw_flagged": raw_flagged, "cleaned_flagged": raw_flagged.copy(),
        "unclassified_counts": show_unclassified_counts(raw_flagged),
    }

def get_crm_row_counts(df_loaded, crm_refs, crm_col=CRM_COL):
    crm_series = df_loaded[crm_col].astype(str).str.strip()
    return {crm: int((crm_series == crm).sum()) for crm in crm_refs}
