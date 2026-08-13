
#!/usr/bin/env python3
"""
CC_AMT_PIPELINE Sanity Check Web App (v2.2)
=============================================
SETUP:  pip install streamlit pandas openpyxl
RUN:    streamlit run sanity_check_app.py

NEW IN v2.2:
- Drag-down detection: Flags columns with values incrementing by +1
- Extra columns detection: Flags columns not in the template (ALL 20 sheets)
- CASE DETAILS validation: All columns mandatory except COMMENT
"""

import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import BytesIO
import streamlit.components.v1 as components

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="CC_AMT_PIPELINE Sanity Checker | LMAQ",
    page_icon="✅",
    layout="wide"
)

# ============================================================
# HIDE ALL STREAMLIT DEFAULT UI ELEMENTS
# ============================================================
hide_streamlit_style = """
<style>
/* Hide hamburger menu */
#MainMenu {visibility: hidden !important; display: none !important;}

/* Hide footer */
footer {visibility: hidden !important; display: none !important;}

/* Hide header/toolbar completely */
header {visibility: hidden !important; display: none !important;}
[data-testid="stHeader"] {display: none !important;}
[data-testid="stToolbar"] {display: none !important;}

/* Hide deploy/manage button - multiple selectors for different versions */
.stDeployButton {display: none !important;}
[data-testid="manage-app-button"] {display: none !important;}
.viewerBadge_container__r5tak {display: none !important;}
[data-testid="stStatusWidget"] {display: none !important;}
.stAppDeployButton {display: none !important;}

/* Nuclear option - hide anything fixed at bottom-right */
div[class*="StatusWidget"] {display: none !important;}
button[kind="manage"] {display: none !important;}
div[data-testid="manage-app-button"] {display: none !important;}

/* Background watermark - Amazon logo + LMAQ */
.watermark {
    position: fixed;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    z-index: -1;
    pointer-events: none;
    text-align: center;
    opacity: 0.05;
}
.watermark img {
    width: 350px;
    display: block;
    margin: 0 auto;
}
.watermark .lmaq-text {
    font-size: 90px;
    font-weight: 900;
    color: #FF9900;
    letter-spacing: 15px;
    font-family: Arial, sans-serif;
    margin-top: 10px;
}
</style>

<div class="watermark">
    <img src="https://upload.wikimedia.org/wikipedia/commons/a/a9/Amazon_logo.svg" alt="">
    <div class="lmaq-text">LMAQ</div>
</div>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# JavaScript to remove "Manage app" button that loads dynamically
components.html("""
<script>
    function hideManageButton() {
        const selectors = [
            '[data-testid="manage-app-button"]',
            '.stDeployButton',
            'button[kind="manage"]',
            '.stAppDeployButton'
        ];
        selectors.forEach(sel => {
            document.querySelectorAll(sel).forEach(el => {
                el.style.display = 'none';
                el.remove();
            });
        });
        try {
            const parentDoc = window.parent.document;
            selectors.forEach(sel => {
                parentDoc.querySelectorAll(sel).forEach(el => {
                    el.style.display = 'none';
                    el.remove();
                });
            });
            const header = parentDoc.querySelector('header');
            if (header) header.style.display = 'none';
            const toolbar = parentDoc.querySelector('[data-testid="stToolbar"]');
            if (toolbar) toolbar.style.display = 'none';
            parentDoc.querySelectorAll('[data-testid="stStatusWidget"]').forEach(el => {
                el.style.display = 'none';
            });
        } catch(e) {}
    }
    hideManageButton();
    setInterval(hideManageButton, 500);
</script>
""", height=0, width=0)

# ============================================================
# TEMPLATE DEFINITION (v2.2 - 20 sheets)
# ============================================================
REQUIRED_SHEETS = [
    "CASE DETAILS", "CAMPUS", "BUILDING", "UNIT", "AID",
    "CREATE_CAMPUS", "MERGE_NODE", "CREATE_BUILDING", "GROUPING",
    "MERGE_BUILDING_NODE", "REPARENT_NODE", "CREATE_DG",
    "MERGE_LG_NODE", "MERGE_DG_NODE", "REPARENT_DG_NODE",
    "REPARENT_LG_NODE", "DEPRECATE_NODE", "ADD_SOURCE_AND_CANONICAL_ADDRES",
    "MOVE_AID", "CREATE_LG"
]

COMMAND_SHEETS = [
    "CREATE_CAMPUS", "MERGE_NODE", "CREATE_BUILDING", "GROUPING",
    "MERGE_BUILDING_NODE", "REPARENT_NODE", "CREATE_DG",
    "MERGE_LG_NODE", "MERGE_DG_NODE", "REPARENT_DG_NODE",
    "REPARENT_LG_NODE", "DEPRECATE_NODE", "ADD_SOURCE_AND_CANONICAL_ADDRES",
    "MOVE_AID", "CREATE_LG"
]

SOURCE_SHEETS = ["CAMPUS", "BUILDING", "UNIT", "AID"]

MANDATORY_SOURCE_COLUMNS = ["Usecase", "Auditor", "Audit Date", "INPUT BPID"]

VALID_USECASES = [
    "CC_AB_Campus_Audit",
    "CC_Cheetah_ORD_AUDIT",
    "CC_ORD_Point_Corrections_EU",
    "CC_ORD_Point_Corrections_NA",
    "CC_ORD_Point_Corrections_Rocket_Stations_NA",
    "CC_SIMS_Audit",
    "DSP_Hierarchy_Building_P0",
    "DSP_Hierarchy_Building_P1",
    "DSP_Hierarchy_Unit_P0",
    "CC_River-DSP_AUDIT"
]

COMMENT_COLUMNS = ["COMMENT", "COMMENTS", "Reviewer comments"]

REVIEWER_COLUMNS = ["Reviewer usecase", "Reviewer alias", " Date", "Reviewer Verdict  ", "Reviewer comments"]

DUPLICATE_KEY_COLUMNS = {
    "CAMPUS": ["Campus Address", "Campus ID(SOURCE)"],
    "BUILDING": ["Building Address", "BPID(SOURCE)"],
    "UNIT": ["Unit Address", "Unit PID(SOURCE)"],
    "AID": ["AID Address", "AID"],
}

MAX_FILE_SIZE_MB = 1

# ============================================================
# TEMPLATE COLUMNS FOR ALL 20 SHEETS
# ============================================================
TEMPLATE_COLUMNS = {
    "CASE DETAILS": [
        "INPUT ID", "AUDIT START DATE", "AUDIT END DATE", "Campus Name",
        "ADDRESS ", "Property Type", "Physical Building Count",
        "SITE MAP AVAILABLE", "SITE MAP LINK", "3P LINK",
        "NO. OF SUB CAMPUSES", "CAMPUS ID PRESENT", "TOTAL NO. OF BUILDING",
        "NO. OF BUILDINGS CREATED", "NO. OF BUILDINGS REPARENT",
        "UNIT-BUILDING RELATION", "UB ESTABLISHED FROM", "TOTAL NO. OF UNIT",
        "NO. OF UNITS CREATED", "NO. OF UNITS REPARENT",
        "CHECK FOR DUPLICATE INPUTS IN FILE", "SANITIZATION CHECK",
        "FILE UPLOAD STATUS", "REFLECTION STATUS", "COMMENT"
    ],
    "CAMPUS": [
        "Reviewer usecase", "Reviewer alias", " Date", "Reviewer Verdict  ",
        "Reviewer comments", "Usecase", "Auditor", "Audit Date", "INPUT BPID",
        "Campus Address", "Campus ID(SOURCE)", "COMMAND", "Error Type",
        "TARGET CAMPUS ID", "COMMENTS", "Uploaded the command(YES/NO/)",
        "UPLOADED THROUGH-CLI/GUI", "REFLECTING YES/NO/CLI-NO", "New ID"
    ],
    "BUILDING": [
        "Reviewer usecase", "Reviewer alias", " Date", "Reviewer Verdict  ",
        "Reviewer comments", "Usecase", "Auditor", "Audit Date", "INPUT BPID",
        "Campus ID", "Building Unique Value", "Building Address", "BPID(SOURCE)",
        "BPID type", "COMMAND", "Error type - Building",
        "Target Node(If action taken is yes)", "COMMENTS",
        "Uploaded the command(YES/NO/)", "UPLOADED THROUGH-CLI/GUI",
        "REFLECTING YES/NO/CLI-NO", "New ID"
    ],
    "UNIT": [
        "Reviewer usecase", "Reviewer alias", " Date", "Reviewer Verdict  ",
        "Reviewer comments", "Usecase", "Auditor", "Audit Date", "INPUT BPID",
        "Campus ID", "Building Unique Value", "BPID", "Unit Address",
        "Unit Label", "Unit PID(SOURCE)", "COMMAND", "Error type - Unit",
        "Target Node(If action taken is yes)",
        "Hierarchy BPID(BPID after taking action)", "COMMENTS",
        "Uploaded the command(YES/NO/)", "UPLOADED THROUGH-CLI/GUI",
        "REFLECTING YES/NO/CLI-NO", "New ID", "Type"
    ],
    "AID": [
        "Reviewer usecase", "Reviewer alias", " Date", "Reviewer Verdict  ",
        "Reviewer comments", "Usecase", "Auditor", "Audit Date", "INPUT BPID",
        "BPID", "UPID", "AID Address", "AID Label", "AID", "COMMAND",
        "Error type - AID", "Target Node(If action taken is yes)",
        "Hierarchy BPID(BPID after taking action)", "COMMENTS",
        "Uploaded the command(YES/NO/)", "UPLOADED THROUGH-CLI/GUI",
        "REFLECTING YES/NO/CLI-NO", "New ID", "Type"
    ],
    "CREATE_CAMPUS": [
        "Usecase", "Auditor", "Audit Date", "INPUT BPID", "COMMAND",
        "isDeliverable", "pid_adminlevel1(country)", "pid_adminlevel3(state)",
        "pid_adminlevel4", "pid_adminlevel5", "pid_adminlevel6(city)",
        "pid_adminlevel7", "pid_adminlevel8", "pid_adminlevel9",
        "pid_adminlevel10", "pid_adminlevel11", "pid_street1", "pid_street2",
        "pid_street3", "pid_unit1name", "pid_unit1number", "pid_unit2name",
        "pid_unit2number", "pid_unit3name", "pid_unit3number", "pid_unit4name",
        "pid_unit4number", "pid_unit5name", "pid_unit5number", "pid_unit6name",
        "pid_unit6number", "pid_unit7name", "pid_unit7number", "pid_unit8name",
        "pid_unit8number", "pid_unit9name", "pid_unit9number",
        "pid_unit10name(Campus label)", "pid_unit10number(campus number)",
        "pid_postalcode1", "pid_postalcode2", "pid_postalcode3", "pid_pobox",
        "pid_addressrole", "parentPid(city node)",
        "Uploaded the command(YES/NO)", "New ID"
    ],
    "MERGE_NODE": [
        "Usecase", "Auditor", "Audit Date", "INPUT BPID", "MERGE_NODE",
        "SOURCE NODE(CAMPUSID)", "TARGET NODE(CAMPUSID)",
        "Uploaded the command(YES/NO/)", "UPLOADED THROUGH-CLI/GUI",
        "REFLECTING YES/NO/CLI-NO"
    ],
    "CREATE_BUILDING": [
        "Usecase", "Auditor", "Audit Date", "INPUT BPID", "Command",
        "isDeliverable", "pid_adminlevel1(country code)",
        "pid_adminlevel3(state)", "pid_adminlevel4", "pid_adminlevel5",
        "pid_adminlevel6(city)", "pid_adminlevel7", "pid_adminlevel8",
        "pid_adminlevel9", "pid_adminlevel10", "pid_adminlevel11",
        "pid_street1(street name)", "pid_street2", "pid_street3",
        "pid_unit1name(building name)", "pid_unit1number(bldg number)",
        "pid_unit2name(unit abbreviation)", "pid_unit2number(unit number)",
        "pid_unit3name", "pid_unit3number", "pid_unit4name", "pid_unit4number",
        "pid_unit5name", "pid_unit5number", "pid_unit6name", "pid_unit6number",
        "pid_unit7name", "pid_unit7number", "pid_unit8name", "pid_unit8number",
        "pid_unit9name", "pid_unit9number", "pid_postalcode1(zipcode)",
        "pid_postalcode2", "pid_postalcode3", "pid_pobox",
        "pid_addressrole(PRIMARY)", "parentPid(CAMPUSID)",
        "Uploaded the command(YES/NO/)", "UPLOADED THROUGH-CLI/GUI",
        "REFLECTING YES/NO/CLI-NO", "New ID"
    ],
    "GROUPING": [
        "Usecase", "Auditor", "Audit Date", "INPUT BPID", "COMMAND",
        "isDeliverable", "pid_adminlevel1(country code)",
        "pid_adminlevel3(state)", "pid_adminlevel4", "pid_adminlevel5",
        "pid_adminlevel6(city)", "pid_adminlevel7", "pid_adminlevel8",
        "pid_adminlevel9", "pid_adminlevel10", "pid_adminlevel11",
        "pid_street1(street name)", "pid_street2", "pid_street3",
        "pid_unit1name(building name)", "pid_unit1number(bldg number)",
        "pid_unit2name(unit abbreviation)", "pid_unit2number(unit number)",
        "pid_unit3name", "pid_unit3number", "pid_unit4name", "pid_unit4number",
        "pid_unit5name", "pid_unit5number", "pid_unit6name", "pid_unit6number",
        "pid_unit7name", "pid_unit7number", "pid_unit8name", "pid_unit8number",
        "pid_unit9name", "pid_unit9number", "pid_postalcode1(zipcode)",
        "pid_postalcode2", "pid_postalcode3", "pid_pobox",
        "pid_addressrole(PRIMARY)", "parentPid(CAMPUSID)",
        "Uploaded the command(YES/NO/)", "UPLOADED THROUGH-CLI/GUI",
        "REFLECTING YES/NO/CLI-NO", "New ID"
    ],
    "MERGE_BUILDING_NODE": [
        "Usecase", "Auditor", "Audit Date", "INPUT BPID", "COMMAND",
        "SOURCE NODE(BPID)", "TARGET NODE(BPID)",
        "Uploaded the command(YES/NO/)", "UPLOADED THROUGH-CLI/GUI",
        "REFLECTING YES/NO/CLI-NO"
    ],
    "REPARENT_NODE": [
        "Usecase", "Auditor", "Audit Date", "INPUT BPID", "COMMAND",
        "SOURCE NODE(BPID)", "TARGET NODE(CAMPUSID)",
        "Uploaded the command(YES/NO/)", "UPLOADED THROUGH-CLI/GUI",
        "REFLECTING YES/NO/CLI-NO"
    ],
    "CREATE_DG": [
        "Usecase", "Auditor", "Audit Date", "INPUT BPID", "CREATE_DG",
        "isDeliverable", "pid_adminlevel1(country code)",
        "pid_adminlevel3(state)", "pid_adminlevel4", "pid_adminlevel5",
        "pid_adminlevel6(city)", "pid_adminlevel7", "pid_adminlevel8",
        "pid_adminlevel9", "pid_adminlevel10", "pid_adminlevel11",
        "pid_street1(street name)", "pid_street2", "pid_street3",
        "pid_unit1name(building name)", "pid_unit1number(bldg number)",
        "pid_unit2name(unit abbreviation)", "pid_unit2number(unit number)",
        "pid_unit3name", "pid_unit3number", "pid_unit4name", "pid_unit4number",
        "pid_unit5name", "pid_unit5number", "pid_unit6name", "pid_unit6number",
        "pid_unit7name", "pid_unit7number", "pid_unit8name", "pid_unit8number",
        "pid_unit9name", "pid_unit9number", "pid_postalcode1(zipcode)",
        "pid_postalcode2", "pid_postalcode3", "pid_pobox",
        "pid_addressrole(PRIMARY)", "parentPid(BPID)",
        "Uploaded the command(YES/NO/)", "UPLOADED THROUGH-CLI/GUI",
        "REFLECTING YES/NO/CLI-NO", "New ID"
    ],
    "MERGE_LG_NODE": [
        "Usecase", "Auditor", "Audit Date", "INPUT BPID", "MERGE_LG_NODE",
        "Source Node (LG)", "Target Node(LG/DG)",
        "Uploaded the command(YES/NO/)", "UPLOADED THROUGH-CLI/GUI",
        "REFLECTING YES/NO/CLI-NO"
    ],
    "MERGE_DG_NODE": [
        "Usecase", "Auditor", "Audit Date", "INPUT BPID", "MERGE_DG_NODE",
        "Source Node (DG)", "Target Node(DG)",
        "Uploaded the command(YES/NO/)", "UPLOADED THROUGH-CLI/GUI",
        "REFLECTING YES/NO/CLI-NO"
    ],
    "REPARENT_DG_NODE": [
        "Usecase", "Auditor", "Audit Date", "INPUT BPID", "SOURCE PPID",
        "REPARENT_DG_NODE", "Source Node (DG)", "Target Node(BPID)",
        "Uploaded the command(YES/NO/)", "UPLOADED THROUGH-CLI/GUI",
        "REFLECTING YES/NO/CLI-NO"
    ],
    "REPARENT_LG_NODE": [
        "Usecase", "Auditor", "Audit Date", "INPUT BPID", "SOURCE PPID",
        "REPARENT_LG_NODE", "Source Node (LG)", "Target Node(BPID)",
        "Uploaded the command(YES/NO/)", "UPLOADED THROUGH-CLI/GUI",
        "REFLECTING YES/NO/CLI-NO"
    ],
    "DEPRECATE_NODE": [
        "Usecase", "Auditor", "Audit Date", "INPUT BPID", "DEPRECATE_NODE",
        "SOURCE_NODE", "NODE TYPE(CAMPUSID/BPID/UPID)"
    ],
    "ADD_SOURCE_AND_CANONICAL_ADDRES": [
        "Usecase", "Auditor", "Audit Date", "INPUT BPID", "COMMAND", "pid",
        "pid_adminlevel1", "pid_adminlevel3", "pid_adminlevel4",
        "pid_adminlevel5", "pid_adminlevel6", "pid_adminlevel7",
        "pid_adminlevel8", "pid_adminlevel9", "pid_adminlevel10",
        "pid_adminlevel11", "pid_street1", "pid_street2", "pid_street3",
        "pid_unit1name", "pid_unit1number", "pid_unit2name", "pid_unit2number",
        "pid_unit3name", "pid_unit3number", "pid_unit4name", "pid_unit4number",
        "pid_unit5name", "pid_unit5number", "pid_unit6name", "pid_unit6number",
        "pid_unit7name", "pid_unit7number", "pid_unit8name", "pid_unit8number",
        "pid_unit9name", "pid_unit9number", "pid_postalcode1", "pid_postalcode2",
        "pid_postalcode3", "pid_pobox", "pid_addressrole"
    ],
    "MOVE_AID": [
        "Usecase", "Auditor", "Audit Date", "INPUT BPID", "MOVE_AID",
        "Source Node (AID)", "Target Node(LG/DG)",
        "Uploaded the command(YES/NO/)", "UPLOADED THROUGH-CLI/GUI",
        "REFLECTING YES/NO/CLI-NO"
    ],
    "CREATE_LG": [
        "Usecase", "Auditor", "Audit Date", "INPUT BPID", "CREATE_LG",
        "Source Node (AID)", "Uploaded the command(YES/NO/)",
        "UPLOADED THROUGH-CLI/GUI", "REFLECTING YES/NO/CLI-NO"
    ]
}

# CASE DETAILS: All columns mandatory EXCEPT "COMMENT" (col Y)
CASE_DETAILS_OPTIONAL = ["COMMENT"]

# ============================================================
# VALIDATION FUNCTIONS
# ============================================================

def check_file_size(file_bytes):
    size_mb = len(file_bytes) / (1024 * 1024)
    if size_mb > MAX_FILE_SIZE_MB:
        return False, f"File is {size_mb:.2f} MB (limit: {MAX_FILE_SIZE_MB} MB)"
    return True, f"File size: {size_mb:.2f} MB ✓"


def check_sheets(xls):
    errors = []
    warnings = []
    actual_sheets = xls.sheet_names

    missing = [s for s in REQUIRED_SHEETS if s not in actual_sheets]
    extra = [s for s in actual_sheets if s not in REQUIRED_SHEETS]

    if missing:
        errors.append(f"MISSING SHEETS: {', '.join(missing)}")
    if extra:
        warnings.append(f"EXTRA SHEETS (unexpected): {', '.join(extra)}")

    return errors, warnings


def get_commands_from_source(xls):
    commands = set()
    for sheet in SOURCE_SHEETS:
        if sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet, header=0)
            if "COMMAND" in df.columns:
                cmds = df["COMMAND"].dropna().astype(str).str.strip().str.upper()
                commands.update(cmds[cmds != ""])
    return commands


def is_excluded_column(col_name, sheet_name):
    col_stripped = col_name.strip()
    if col_stripped in [c.strip() for c in COMMENT_COLUMNS]:
        return True
    if sheet_name in SOURCE_SHEETS:
        if col_stripped in [c.strip() for c in REVIEWER_COLUMNS]:
            return True
    return False


def is_alpha_only(s):
    return bool(re.match(r'^[a-zA-Z]+$', str(s).strip()))


def is_valid_date(val):
    if pd.isna(val):
        return False
    if isinstance(val, (pd.Timestamp, datetime)):
        return True
    s = str(val).strip()
    if not s:
        return False
    try:
        parts = s.split("/")
        if len(parts) == 3:
            m_str, d_str, y_str = parts[0], parts[1], parts[2]
            if len(m_str) in [1, 2] and len(d_str) in [1, 2] and len(y_str) == 4:
                m, d, y = int(m_str), int(d_str), int(y_str)
                if 1 <= m <= 12 and 1 <= d <= 31 and 2020 <= y <= 2030:
                    return True
    except (ValueError, IndexError):
        pass
    return False


def check_mandatory_columns(df, sheet_name):
    errors = []
    if sheet_name not in SOURCE_SHEETS:
        return errors

    for col_name in MANDATORY_SOURCE_COLUMNS:
        if col_name in df.columns:
            blank_count = df[col_name].apply(
                lambda v: pd.isna(v) or str(v).strip() == ""
            ).sum()
            if blank_count > 0:
                blank_rows = df[df[col_name].apply(
                    lambda v: pd.isna(v) or str(v).strip() == ""
                )].index + 2
                sample_rows = blank_rows.tolist()[:5]
                errors.append(
                    f"MANDATORY column '{col_name}' has {blank_count} blank cell(s) "
                    f"at Excel row(s): {sample_rows}{'...' if blank_count > 5 else ''}"
                )
    return errors


def check_duplicate_rows(df, sheet_name):
    if sheet_name not in DUPLICATE_KEY_COLUMNS:
        return None

    key_columns = DUPLICATE_KEY_COLUMNS[sheet_name]
    available_keys = [c for c in key_columns if c in df.columns]

    if not available_keys or len(available_keys) < 2 or len(df) == 0:
        return None

    df_check = df[available_keys].copy()
    non_blank_mask = df_check.apply(
        lambda row: all(not (pd.isna(v) or str(v).strip() == "") for v in row), axis=1
    )
    df_check = df_check[non_blank_mask]

    if len(df_check) == 0:
        return None

    duplicates = df_check[df_check.duplicated(keep=False)]
    if len(duplicates) > 0:
        dup_count = df_check.duplicated(keep='first').sum()
        dup_rows = (duplicates[duplicates.duplicated(keep='first')].index + 2).tolist()[:5]
        return (
            f"{dup_count} DUPLICATE ROW(S) based on "
            f"[{' + '.join(available_keys)}] at Excel row(s): "
            f"{dup_rows}{'...' if dup_count > 5 else ''}"
        )
    return None


# ============================================================
# NEW: EXTRA COLUMNS DETECTION
# ============================================================

def check_extra_columns(df, sheet_name):
    """Check if sheet has columns not present in the template."""
    if sheet_name not in TEMPLATE_COLUMNS:
        return []
    
    expected_cols = TEMPLATE_COLUMNS[sheet_name]
    actual_cols = list(df.columns)
    
    extra_cols = [col for col in actual_cols if col not in expected_cols]
    return extra_cols


# ============================================================
# NEW: DRAG-DOWN DETECTION (+1 increment)
# ============================================================

def get_numeric_suffix(val):
    """Extract trailing numeric characters from a string."""
    suffix = ""
    for ch in reversed(str(val)):
        if ch.isdigit():
            suffix = ch + suffix
        else:
            break
    return suffix


def check_drag_down(df, sheet_name):
    """
    Detect drag-down/fill-down errors where values increment by +1.
    Checks ALL columns except Reviewer and Comment columns.
    Returns list of flagged column names.
    """
    skip = [c.strip() for c in REVIEWER_COLUMNS + COMMENT_COLUMNS]
    check_columns = [col for col in df.columns if col.strip() not in skip]
    
    flagged_columns = []
    
    for col in check_columns:
        values = df[col].dropna().reset_index(drop=True)
        
        # Need at least 4 values to detect pattern (3 consecutive increments)
        if len(values) < 4:
            continue
        
        consecutive_count = 0
        
        for i in range(1, len(values)):
            curr_val = str(values.iloc[i]).strip()
            prev_val = str(values.iloc[i - 1]).strip()
            
            if curr_val == "" or prev_val == "":
                consecutive_count = 0
                continue
            
            is_increment = False
            
            # Case 1: Both are pure numbers
            try:
                curr_num = float(curr_val)
                prev_num = float(prev_val)
                if curr_num - prev_num == 1:
                    is_increment = True
            except (ValueError, TypeError):
                # Case 2: Text ending with numbers (e.g., BPID_001, BPID_002)
                curr_suffix = get_numeric_suffix(curr_val)
                prev_suffix = get_numeric_suffix(prev_val)
                
                if curr_suffix and prev_suffix:
                    curr_prefix = curr_val[:-len(curr_suffix)]
                    prev_prefix = prev_val[:-len(prev_suffix)]
                    
                    if curr_prefix == prev_prefix:
                        try:
                            if float(curr_suffix) - float(prev_suffix) == 1:
                                is_increment = True
                        except (ValueError, TypeError):
                            pass
            
            if is_increment:
                consecutive_count += 1
                if consecutive_count >= 3:
                    flagged_columns.append(col)
                    break
            else:
                consecutive_count = 0
    
    return flagged_columns


# ============================================================
# NEW: CASE DETAILS VALIDATION
# ============================================================

def validate_case_details(xls):
    """Validate CASE DETAILS sheet - all columns mandatory except COMMENT."""
    errors = []
    warnings = []
    passes = []
    
    if "CASE DETAILS" not in xls.sheet_names:
        return errors, warnings, passes
    
    df = pd.read_excel(xls, sheet_name="CASE DETAILS", header=0)
    
    if len(df) == 0:
        errors.append("CASE DETAILS sheet is EMPTY (must have at least 1 row)")
        return errors, warnings, passes
    
    # Check extra columns
    extra_cols = check_extra_columns(df, "CASE DETAILS")
    if extra_cols:
        errors.append(
            f"EXTRA COLUMNS not in template: {', '.join(extra_cols)} "
            f"(only template headers allowed)"
        )
    else:
        passes.append("Column headers match template")
    
    # Check mandatory columns (all except COMMENT)
    mandatory_cols = [c for c in TEMPLATE_COLUMNS["CASE DETAILS"] 
                      if c not in CASE_DETAILS_OPTIONAL]
    
    blank_cols = []
    for col in mandatory_cols:
        if col in df.columns:
            blank_count = df[col].apply(
                lambda v: pd.isna(v) or str(v).strip() == ""
            ).sum()
            if blank_count > 0:
                blank_cols.append(f"'{col}' ({blank_count} blank)")
    
    if blank_cols:
        errors.append(
            f"MANDATORY columns have blank cells: {'; '.join(blank_cols[:5])}"
            f"{'...' if len(blank_cols) > 5 else ''}"
        )
    else:
        passes.append("All mandatory columns filled (COMMENT is optional)")
    
    # Drag-down check on CASE DETAILS
    drag_cols = check_drag_down(df, "CASE DETAILS")
    if drag_cols:
        errors.append(
            f"DRAG-DOWN DETECTED: Values incrementing by +1 in column(s): "
            f"{', '.join(drag_cols)} (likely fill-down error)"
        )
    else:
        passes.append("No drag-down patterns detected")
    
    return errors, warnings, passes


# ============================================================
# MAIN SHEET VALIDATION
# ============================================================

def validate_sheet(xls, sheet_name, commands):
    errors = []
    warnings = []
    passes = []

    if sheet_name not in xls.sheet_names:
        return errors, warnings, passes

    # Skip CASE DETAILS here (handled separately)
    if sheet_name == "CASE DETAILS":
        return errors, warnings, passes

    df = pd.read_excel(xls, sheet_name=sheet_name, header=0)

    if sheet_name in COMMAND_SHEETS:
        sheet_upper = sheet_name.upper()
        command_found = any(
            sheet_upper in cmd or cmd in sheet_upper
            for cmd in commands
        )
        if not command_found:
            passes.append("Skipped (command not used in source sheets)")
            return errors, warnings, passes
        else:
            if len(df) == 0:
                errors.append("EMPTY but command exists in source sheets")
                return errors, warnings, passes

    if len(df) == 0:
        return errors, warnings, passes

    # --- CHECK: Extra Columns ---
    if sheet_name in TEMPLATE_COLUMNS:
        extra_cols = check_extra_columns(df, sheet_name)
        if extra_cols:
            errors.append(
                f"EXTRA COLUMNS not in template: {', '.join(extra_cols)} "
                f"(only template headers allowed)"
            )
        else:
            passes.append("Column headers match template")

    # --- Mandatory columns check (F, G, H, I) for source sheets ---
    if sheet_name in SOURCE_SHEETS:
        mandatory_errors = check_mandatory_columns(df, sheet_name)
        for err in mandatory_errors:
            errors.append(err)
        if not mandatory_errors:
            passes.append("Mandatory columns (Usecase, Auditor, Audit Date, INPUT BPID) all filled")

    # --- Usecase validation ---
    if sheet_name in SOURCE_SHEETS or sheet_name in COMMAND_SHEETS:
        if "Usecase" in df.columns:
            usecase_vals = df["Usecase"].dropna().astype(str).str.strip()
            usecase_vals = usecase_vals[usecase_vals != ""]
            if len(usecase_vals) > 0:
                invalid = usecase_vals[~usecase_vals.isin(VALID_USECASES)]
                if len(invalid) > 0:
                    sample = invalid.head(3).tolist()
                    errors.append(
                        f"INVALID USECASE in {len(invalid)} row(s). "
                        f"Examples: {sample}"
                    )
                else:
                    passes.append("Usecase values valid")

    # --- Auditor validation ---
    if sheet_name in SOURCE_SHEETS or sheet_name in COMMAND_SHEETS:
        if "Auditor" in df.columns:
            auditor_vals = df["Auditor"].dropna().astype(str).str.strip()
            auditor_vals = auditor_vals[auditor_vals != ""]
            if len(auditor_vals) > 0:
                invalid_auditors = auditor_vals[~auditor_vals.apply(is_alpha_only)]
                if len(invalid_auditors) > 0:
                    sample = invalid_auditors.head(3).tolist()
                    errors.append(
                        f"INVALID AUDITOR (must be alphabetical only) in "
                        f"{len(invalid_auditors)} row(s). Examples: {sample}"
                    )
                else:
                    passes.append("Auditor values valid")

    # --- Audit Date validation (M/DD/YYYY or MM/DD/YYYY format) ---
    if sheet_name in SOURCE_SHEETS or sheet_name in COMMAND_SHEETS:
        if "Audit Date" in df.columns:
            date_col = df["Audit Date"]
            non_blank_dates = date_col.dropna()
            non_blank_dates = non_blank_dates[non_blank_dates.astype(str).str.strip() != ""]

            if len(non_blank_dates) > 0:
                invalid_dates = 0
                invalid_rows = []
                for idx, val in date_col.items():
                    if not (pd.isna(val) or str(val).strip() == ""):
                        if not is_valid_date(val):
                            invalid_dates += 1
                            if len(invalid_rows) < 5:
                                invalid_rows.append(idx + 2)

                if invalid_dates > 0:
                    errors.append(
                        f"INVALID AUDIT DATE FORMAT in {invalid_dates} row(s) "
                        f"(must be M/DD/YYYY or MM/DD/YYYY) at Excel row(s): "
                        f"{invalid_rows}{'...' if invalid_dates > 5 else ''}"
                    )
                else:
                    passes.append("Audit Date format valid")

    # --- Blank rows check (excluding comment/reviewer columns) ---
    check_cols = [c for c in df.columns if not is_excluded_column(c, sheet_name)]
    if check_cols:
        df_check = df[check_cols]
        blank_mask = df_check.apply(
            lambda row: all(pd.isna(v) or str(v).strip() == "" for v in row), axis=1
        )
        blank_count = blank_mask.sum()
        if blank_count > 0:
            blank_rows = (blank_mask[blank_mask].index + 2).tolist()[:5]
            errors.append(
                f"{blank_count} BLANK ROW(S) found at Excel row(s): "
                f"{blank_rows}{'...' if blank_count > 5 else ''}"
            )
        else:
            passes.append("No blank rows")

    # --- Duplicate row check (Address + ID) ---
    if sheet_name in DUPLICATE_KEY_COLUMNS:
        dup_result = check_duplicate_rows(df, sheet_name)
        if dup_result:
            warnings.append(dup_result)
        else:
            passes.append("No duplicate rows detected")

    # --- CHECK: Drag-Down Detection ---
    if sheet_name in SOURCE_SHEETS:
        drag_cols = check_drag_down(df, sheet_name)
        if drag_cols:
            errors.append(
                f"DRAG-DOWN DETECTED: Values incrementing by +1 in column(s): "
                f"{', '.join(drag_cols)} (likely fill-down error)"
            )
        else:
            passes.append("No drag-down patterns detected")

    return errors, warnings, passes


def generate_report(filename, all_errors, all_warnings, all_passes):
    lines = []
    lines.append("=" * 60)
    lines.append("CC_AMT_PIPELINE SANITY CHECK REPORT | LMAQ Team")
    lines.append("=" * 60)
    lines.append(f"File: {filename}")
    lines.append(f"Date: {datetime.now().strftime('%d-%b-%Y %H:%M:%S')}")
    lines.append(f"Template Version: v2.2 (20 sheets)")
    lines.append("")

    if all_errors:
        lines.append(f"RESULT: FAILED ({len(all_errors)} error(s))")
    else:
        lines.append("RESULT: PASSED")
    lines.append("")

    if all_errors:
        lines.append("-" * 40)
        lines.append("ERRORS (Must Fix):")
        lines.append("-" * 40)
        for i, (sheet, err) in enumerate(all_errors, 1):
            lines.append(f"  {i}. [{sheet}] {err}")
        lines.append("")

    if all_warnings:
        lines.append("-" * 40)
        lines.append("WARNINGS (Review Recommended):")
        lines.append("-" * 40)
        for i, (sheet, warn) in enumerate(all_warnings, 1):
            lines.append(f"  {i}. [{sheet}] {warn}")
        lines.append("")

    lines.append("-" * 40)
    lines.append(f"PASSED CHECKS ({len(all_passes)}):")
    lines.append("-" * 40)
    for i, (sheet, pas) in enumerate(all_passes, 1):
        lines.append(f"  {i}. [{sheet}] {pas}")

    lines.append("")
    lines.append("=" * 60)
    lines.append("Generated by LMAQ Sanity Checker v2.2")
    lines.append("=" * 60)

    return "\n".join(lines)


# ============================================================
# STREAMLIT UI
# ============================================================

# Title
st.title("🔍 CC_AMT_PIPELINE Sanity Checker")
st.markdown("Upload your Excel file below to validate before WorkDocs upload.")
st.markdown("---")

# Sidebar
with st.sidebar:
    st.markdown("""
    <div style="text-align:center; padding:10px; background:#232F3E; border-radius:8px; margin-bottom:16px;">
        <div style="color:#FF9900; font-size:28px; font-weight:900; letter-spacing:4px;">LMAQ</div>
        <div style="color:#ADB5BD; font-size:11px;">Last Mile Analytics & Quality</div>
    </div>
    """, unsafe_allow_html=True)

    st.header("📋 Validation Rules")
    st.markdown("""
    1. **File size** < 1 MB
    2. **20 sheets** required (no missing)
    3. **Extra columns** flagged (only template headers allowed)
    4. **CASE DETAILS** all columns mandatory (except COMMENT)
    5. **Mandatory columns** (F,G,H,I) in source sheets cannot be blank
    6. **No blank rows** (excl. Comments & Reviewer cols)
    7. **Usecase** must be from approved list (case-sensitive)
    8. **Auditor** must be alphabetical only (alias)
    9. **Audit Date** mandatory, **M/DD/YYYY** or **MM/DD/YYYY** format
    10. **Command sheets** validated only if command in source
    11. **Duplicate rows** flagged (Address + ID combination)
    12. **Drag-down** detection (values +1 increment = fill error)
    """)

    st.markdown("---")
    st.header("✅ Valid Usecases")
    for uc in VALID_USECASES:
        st.code(uc, language=None)

    st.markdown("---")
    st.header("📂 Required Sheets (20)")
    st.markdown("**Source Sheets:**")
    for s in SOURCE_SHEETS:
        st.text(f"  • {s}")
    st.markdown("**Command Sheets:**")
    for s in COMMAND_SHEETS:
        st.text(f"  • {s}")
    st.markdown("**Other:**")
    st.text("  • CASE DETAILS")

    st.markdown("---")
    st.header("🔑 Duplicate Check Keys")
    for sheet, cols in DUPLICATE_KEY_COLUMNS.items():
        st.text(f"  {sheet}: {' + '.join(cols)}")

# File upload
uploaded_file = st.file_uploader(
    "📁 Upload your CC_AMT_PIPELINE Excel file",
    type=["xlsx", "xls"],
    help="File must be .xlsx format and under 1 MB"
)

if uploaded_file is not None:
    st.markdown("---")

    file_bytes = uploaded_file.read()
    uploaded_file.seek(0)

    progress = st.progress(0, text="Starting validation...")

    all_errors = []
    all_warnings = []
    all_passes = []

    # CHECK 1: File Size
    progress.progress(5, text="Checking file size...")
    size_ok, size_msg = check_file_size(file_bytes)
    if not size_ok:
        all_errors.append(("FILE", size_msg))
    else:
        all_passes.append(("FILE", size_msg))

    # Load Excel
    progress.progress(15, text="Loading Excel file...")
    try:
        xls = pd.ExcelFile(BytesIO(file_bytes))
    except Exception as ex:
        st.error(f"❌ Cannot read file: {str(ex)}")
        st.stop()

    # CHECK 2: Sheet Structure
    progress.progress(25, text="Checking sheet structure...")
    sheet_errors, sheet_warnings = check_sheets(xls)
    for err in sheet_errors:
        all_errors.append(("SHEETS", err))
    for warn in sheet_warnings:
        all_warnings.append(("SHEETS", warn))
    if not sheet_errors and not sheet_warnings:
        all_passes.append(("SHEETS", f"All 20 required sheets present ({len(xls.sheet_names)} total)"))

    # CHECK 3: CASE DETAILS validation
    progress.progress(30, text="Validating CASE DETAILS...")
    cd_errors, cd_warnings, cd_passes = validate_case_details(xls)
    for err in cd_errors:
        all_errors.append(("CASE DETAILS", err))
    for warn in cd_warnings:
        all_warnings.append(("CASE DETAILS", warn))
    for pas in cd_passes:
        all_passes.append(("CASE DETAILS", pas))

    # CHECK 4: Get commands from source sheets
    progress.progress(35, text="Reading commands from source sheets...")
    commands = get_commands_from_source(xls)
    if commands:
        all_passes.append(("COMMANDS", f"Found {len(commands)} unique command(s): {', '.join(sorted(commands))}"))

    # CHECK 5+: Per-sheet validation
    total_sheets = len(REQUIRED_SHEETS)
    for idx, sheet_name in enumerate(REQUIRED_SHEETS):
        pct = 35 + int((idx / total_sheets) * 60)
        progress.progress(pct, text=f"Validating: {sheet_name}...")

        sheet_errors, sheet_warnings, sheet_passes = validate_sheet(xls, sheet_name, commands)

        for err in sheet_errors:
            all_errors.append((sheet_name, err))
        for warn in sheet_warnings:
            all_warnings.append((sheet_name, warn))
        for pas in sheet_passes:
            all_passes.append((sheet_name, pas))

    progress.progress(100, text="✅ Validation complete!")

    # DISPLAY RESULTS
    st.markdown("---")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        if len(all_errors) == 0:
            st.metric("Result", "✅ PASSED", delta="Ready for upload")
        else:
            st.metric("Result", "❌ FAILED", delta=f"{len(all_errors)} error(s)")
    with col2:
        st.metric("Errors", len(all_errors))
    with col3:
        st.metric("Warnings", len(all_warnings))
    with col4:
        st.metric("Checks Passed", len(all_passes))

    st.markdown("---")

    if len(all_errors) == 0:
        st.success("## 🎉 ALL CHECKS PASSED!\nYour file is ready for WorkDocs upload.")
    else:
        st.error(
            f"## 🚫 VALIDATION FAILED\n"
            f"{len(all_errors)} error(s) found. Fix and re-upload."
        )

    if all_errors:
        st.markdown("### ❌ Errors (Must Fix)")
        for idx, (sheet, err) in enumerate(all_errors, 1):
            st.markdown(f"**{idx}. [{sheet}]** {err}")

    if all_warnings:
        st.markdown("### ⚠️ Warnings (Review Recommended)")
        for idx, (sheet, warn) in enumerate(all_warnings, 1):
            st.markdown(f"{idx}. **[{sheet}]** {warn}")

    with st.expander(f"✅ Passed Checks ({len(all_passes)})", expanded=False):
        for idx, (sheet, pas) in enumerate(all_passes, 1):
            st.markdown(f"{idx}. **[{sheet}]** {pas}")

    with st.expander("📄 File Information"):
        st.text(f"Filename: {uploaded_file.name}")
        st.text(f"Size: {len(file_bytes) / 1024:.1f} KB")
        st.text(f"Sheets found: {len(xls.sheet_names)}")
        st.text(f"Sheets in file: {', '.join(xls.sheet_names)}")
        st.text(f"Commands in source: {', '.join(sorted(commands)) if commands else 'None'}")
        st.text(f"Validated at: {datetime.now().strftime('%d-%b-%Y %H:%M:%S')}")

    st.markdown("---")
    report_text = generate_report(
        uploaded_file.name, all_errors, all_warnings, all_passes
    )
    st.download_button(
        label="📥 Download Validation Report",
        data=report_text,
        file_name=f"validation_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
        mime="text/plain"
    )

else:
    st.info("👆 Upload your Excel file above to start validation.")

    st.markdown("### 🚀 Quick Start")
    st.markdown("""
    1. Click **Browse files** above
    2. Select your CC_AMT_PIPELINE Excel file
    3. Results appear instantly!
    """)

    st.markdown("### 📝 What's New in v2.2")
    st.markdown("""
    - **Drag-down detection**: Flags columns where values increment by +1 (fill-down error)
    - **Extra columns detection**: Flags columns not in template (ALL 20 sheets)
    - **CASE DETAILS validation**: All columns mandatory except COMMENT
    - **Date format**: Accepts M/DD/YYYY or MM/DD/YYYY
    - All previous checks from v2.1 intact
    """)

