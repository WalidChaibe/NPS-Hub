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

# ── Commercial Reasons ────────────────────────────────────────────────────
# Rows where Root Cause == "Pre-Agreement" AND Reason is one of these
# are classified as "Commercial" instead of going through the normal classifier.
_COMMERCIAL_REASONS = {
    "Sales Discount",
    "FOC",
    "Tools cost reimbursement",
}
_COMMERCIAL_ROOT_CAUSE_LOWER = "pre-agreement"  # matched case-insensitively

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

# ── Manual CRM Reason / Root Cause Overrides ─────────────────────────────
# These CRM refs were manually corrected outside the source export. Their
# Reason (and Root Cause, if that column is present) get force-set to the
# values below, regardless of what the uploaded file says.
#
# IMPORTANT ORDERING NOTE: this map is applied BEFORE _CRM_DELETE_MAP_DEFAULT
# is used to drop rows. apply_crm_deletions() only removes the LAST n rows
# matching a CRM ref — if a ref appears in both maps and n covers every row
# for that ref, the override must already be baked in, or the correction is
# lost with no trace. See _check_override_delete_overlap() below, which
# raises a visible warning if a future edit puts the same ref in both maps.
# Format: crm_ref -> (Reason, Root Cause)
_CRM_REASON_OVERRIDE_DEFAULT = {
    "EPAK-CRM-10245": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10271": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-10274": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-10275": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-10296": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10242": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10295": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-10289": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-10281": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10326": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-10324": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-10286": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10241": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10007": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-10323": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-10331": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-10276": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10300": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10287": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10273": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10355": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-10347": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10320": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10321": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10263": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10378": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-10388": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10379": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10357": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10308": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10409": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10412": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10408": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-10432": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-10360": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-10382": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-10391": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10397": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-10380": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10345": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10441": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10465": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10336": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10440": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10466": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10341": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10488": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10505": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10523": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10493": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10570": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10558": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10518": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10596": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10597": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10595": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10592": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10620": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-10538": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-10564": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-10536": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10543": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-10588": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10547": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-10672": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10640": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-10625": ("Incorrect Loading Manifest", "Missing sign loading manifest"),
    "EPAK-CRM-10676": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10681": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10484": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10693": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-10682": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10521": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10628": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-10627": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10530": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10696": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10712": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10571": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10589": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10641": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10647": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10516": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10535": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10664": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10708": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-10534": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10716": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10720": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10739": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10656": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10559": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10565": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10761": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-10797": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10792": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10805": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-10768": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-10786": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-10717": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-10731": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10555": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10548": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10796": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10767": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-10839": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10840": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10854": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10869": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10661": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10554": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10832": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10872": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10865": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10884": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-10897": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10933": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-10623": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-10783": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-10943": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-10944": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-10929": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-10976": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-10874": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-10866": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-11003": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11016": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11030": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-11020": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11029": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11026": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11033": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-11067": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11070": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-11074": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-11045": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-11077": ("Incorrect Loading Manifest", "Missing sign loading manifest"),
    "EPAK-CRM-11072": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-11073": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-11063": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-11086": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-11098": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-11090": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11093": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-11088": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-11099": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11065": ("Incorrect Loading Manifest", "Missing sign loading manifest"),
    "EPAK-CRM-11114": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-11120": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11125": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11154": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-11157": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-11161": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11173": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-11183": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11194": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11196": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-11199": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-11214": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11220": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-11207": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-11206": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-11216": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-11236": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11235": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11247": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-11232": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11259": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-11275": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11279": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11310": ("Incorrect Loading Manifest", "Missing sign loading manifest"),
    "EPAK-CRM-11294": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-11316": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11328": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11341": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-11342": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-11357": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11359": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11339": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11325": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-11326": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-11327": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-11296": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-11329": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11344": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11374": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-11366": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11379": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-11369": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
    "EPAK-CRM-11356": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-11376": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11330": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-11375": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11381": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-11389": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11409": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-11406": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11422": ("Sales Discount", "Pre-Agreement"),
    "EPAK-CRM-11420": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11421": ("FOC", "Pre-Agreement"),
    "EPAK-CRM-11419": ("Trailer Waiting Hours Exceeded", "No place at customer warehouse"),
    "EPAK-CRM-11400": ("Trailer Waiting Hours Exceeded", "Customer change delivery schedule without intimation"),
}

def _check_override_delete_overlap(delete_map, override_map):
    """
    If a CRM ref appears in both the delete map and the override map, the
    override MUST be applied before deletion or it can be silently lost
    (apply_crm_deletions can remove every row for that ref). This doesn't
    fix that — it can't, by definition, if overrides run first — but it
    surfaces the collision so it never passes unnoticed.
    """
    overlap = set(delete_map.keys()) & set(override_map.keys())
    if overlap:
        import warnings
        warnings.warn(
            f"CRM ref(s) present in BOTH the delete map and the reason-override map: "
            f"{sorted(overlap)}. Verify override_crm_reasons() runs before apply_crm_deletions() "
            f"for these, and that the delete count doesn't remove every row for the ref.",
            stacklevel=2,
        )
    return overlap

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
        crm_rows = supabase.table("qm_crm_delete_map").select("crm_ref,delete_count").execute()
        crm_map  = {r["crm_ref"]: r["delete_count"] for r in (crm_rows.data or [])}
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

def load_crm_reason_overrides_from_supabase(supabase=None):
    """Returns {crm_ref: (reason, root_cause)}. Falls back to the hardcoded default."""
    if supabase is None:
        return _CRM_REASON_OVERRIDE_DEFAULT
    try:
        rows = supabase.table("qm_crm_reason_override").select("crm_ref,reason,root_cause").execute()
        override_map = {r["crm_ref"]: (r["reason"], r.get("root_cause")) for r in (rows.data or [])}
        if not override_map:
            override_map = _CRM_REASON_OVERRIDE_DEFAULT
    except Exception:
        override_map = _CRM_REASON_OVERRIDE_DEFAULT
    return override_map

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

def _is_commercial(reason_clean: str, root_cause_val) -> bool:
    """Return True if this row should be classified as Commercial."""
    if reason_clean not in _COMMERCIAL_REASONS:
        return False
    # If root cause column exists, also require Pre-Agreement
    if root_cause_val is not None:
        rc = _clean_text(str(root_cause_val)).lower()
        return rc == _COMMERCIAL_ROOT_CAUSE_LOWER
    # No root cause column — reason alone is sufficient
    return True

def _find_root_cause_col(df):
    rc_col = next((c for c in df.columns if "root" in c.lower() and "cause" in c.lower()), None)
    if rc_col is None:
        rc_col = next((c for c in df.columns if "root" in c.lower()), None)
    return rc_col

def apply_crm_reason_overrides(df, override_map, crm_col=CRM_COL, df_name="df"):
    """
    Force-sets Reason (and Root Cause, if that column exists) for specific CRM refs
    that were manually corrected outside the source export.

    MUST run before apply_crm_deletions(). apply_crm_deletions() drops the LAST n
    rows matching a CRM ref — if n covers every row for that ref, those rows are
    gone before classification ever sees them. Applying the override first means
    the correction survives regardless of what deletion does to the surviving rows.
    """
    if not override_map:
        return df
    rc_col = _find_root_cause_col(df)
    crm_series = df[crm_col].astype(str).str.strip()
    for crm, (reason, root_cause) in override_map.items():
        mask = crm_series == crm
        if not mask.any():
            continue
        df.loc[mask, REASON_COL] = reason
        if rc_col is not None:
            df.loc[mask, rc_col] = root_cause
    return df

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

    # Auto-detect Root Cause column
    rc_col = next((c for c in df.columns if "root" in c.lower() and "cause" in c.lower()), None)
    if rc_col is None:
        rc_col = next((c for c in df.columns if "root" in c.lower()), None)

    # Vectorised commercial detection — no row-by-row loop
    reason_in_commercial = reason.isin(_COMMERCIAL_REASONS)
    if rc_col:
        rc_clean = df[rc_col].map(lambda x: _clean_text(str(x)).lower() if pd.notna(x) else "")
        is_commercial = reason_in_commercial & rc_clean.eq(_COMMERCIAL_ROOT_CAUSE_LOWER)
    else:
        is_commercial = reason_in_commercial

    # Build category column vectorised: Commercial > Invalid > classifier
    cat = reason.map(classifier)                     # default: run classifier on every row
    cat[~df["Is_Valid"]] = "Invalid"                 # override invalid rows
    cat[is_commercial]   = "Commercial"              # commercial overrides everything

    df["Complaint_Category"] = cat
    return df

def add_date_and_flags_ncr(df, date_col, df_name="NCR", classifier=None):
    if classifier is None: classifier = _classify_reason_default
    df["Base_Date"] = pd.to_datetime(df[date_col], errors="coerce")
    df["Year"]  = df["Base_Date"].dt.year
    df["Month"] = df["Base_Date"].dt.month
    gen_cat = df[GEN_CAT_COL].map(_clean_text)
    reason  = df[REASON_COL].map(_clean_text)
    df["Is_Valid"] = gen_cat.eq("Work In Progress")

    # Auto-detect Root Cause column
    rc_col = next((c for c in df.columns if "root" in c.lower() and "cause" in c.lower()), None)
    if rc_col is None:
        rc_col = next((c for c in df.columns if "root" in c.lower()), None)

    # Vectorised commercial detection
    reason_in_commercial = reason.isin(_COMMERCIAL_REASONS)
    if rc_col:
        rc_clean = df[rc_col].map(lambda x: _clean_text(str(x)).lower() if pd.notna(x) else "")
        is_commercial = reason_in_commercial & rc_clean.eq(_COMMERCIAL_ROOT_CAUSE_LOWER)
    else:
        is_commercial = reason_in_commercial

    cat = reason.map(classifier)
    cat[~df["Is_Valid"]] = "Invalid"
    cat[is_commercial]   = "Commercial"

    df["Complaint_Category"] = cat
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
                                crm_delete_map=None, classifier=None, crm_reason_override=None):
    if crm_delete_map is None: crm_delete_map = _CRM_DELETE_MAP_DEFAULT
    if classifier is None: classifier = _classify_reason_default
    if crm_reason_override is None: crm_reason_override = _CRM_REASON_OVERRIDE_DEFAULT

    # Check the maps ACTUALLY being used (not just the hardcoded defaults) —
    # these may come from Supabase and change independently over time.
    overlap = _check_override_delete_overlap(crm_delete_map, crm_reason_override)

    # Override runs FIRST, before any deletion touches Reason/Root Cause.
    df_loaded = apply_crm_reason_overrides(df_loaded.copy(), crm_reason_override,
                                            crm_col=CRM_COL, df_name=dataset_name)

    # raw_flagged: flag the override-corrected original (no crm deletions)
    raw_flagged     = add_date_and_flags_final_issued(df_loaded.copy(), date_col=date_col,
                                                       df_name=f"{dataset_name}.raw", classifier=classifier)
    # cleaned: apply CRM deletions (on already-overridden data), then flag
    cleaned         = apply_crm_deletions(df_loaded, crm_delete_map, crm_col=CRM_COL, df_name=dataset_name)
    cleaned_flagged = add_date_and_flags_final_issued(cleaned, date_col=date_col,
                                                       df_name=f"{dataset_name}.cleaned", classifier=classifier)
    return {
        "raw_flagged": raw_flagged, "cleaned_flagged": cleaned_flagged,
        "unclassified_counts": show_unclassified_counts(cleaned_flagged),
        "override_delete_overlap": overlap,
    }

def build_dataset_ncr(df_loaded, date_col, dataset_name="NCR", classifier=None):
    if classifier is None: classifier = _classify_reason_default
    raw_flagged = add_date_and_flags_ncr(df_loaded, date_col=date_col,
                                          df_name=dataset_name, classifier=classifier)
    return {
        "raw_flagged": raw_flagged, "cleaned_flagged": raw_flagged,
        "unclassified_counts": show_unclassified_counts(raw_flagged),
    }

def get_crm_row_counts(df_loaded, crm_refs, crm_col=CRM_COL):
    crm_series = df_loaded[crm_col].astype(str).str.strip()
    return {crm: int((crm_series == crm).sum()) for crm in crm_refs}