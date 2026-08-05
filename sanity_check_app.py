
import streamlit as st
import pandas as pd
import os
from datetime import datetime

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="LMAQ Sanity Checker",
    page_icon="✅",
    layout="wide"
)

# ============================================================
# CUSTOM CSS - Amazon theme + branding
# ============================================================
st.markdown("""
<style>
    /* Hide Streamlit menu, footer, deploy button */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stDeployButton {display: none;}
    header {visibility: hidden;}
    [data-testid="manage-app-button"] {display: none;}
    .viewerBadge_container__r5tak {display: none;}
    [data-testid="stToolbar"] {display: none;}
    
    /* Background watermark */
    .stApp::before {
        content: "amazon";
        position: fixed;
        bottom: 20px;
        right: 30px;
        font-size: 40px;
        font-weight: bold;
        color: rgba(255, 153, 0, 0.08);
        z-index: 0;
        pointer-events: none;
    }
    .stApp::after {
        content: "LMAQ";
        position: fixed;
        bottom: 70px;
        right: 30px;
        font-size: 20px;
        font-weight: bold;
        color: rgba(35, 47, 62, 0.06);
        z-index: 0;
        pointer-events: none;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
# CONSTANTS
# ============================================================
REQUIRED_SHEETS = [
    "CASE DETAILS", "CAMPUS", "BUILDING", "UNIT", "AID",
    "CREATE_CAMPUS", "MERGE_NODE", "CREATE_BUILDING", "GROUPING",
    "MERGE_BUILDING_NODE", "REPARENT_NODE", "CREATE_DG",
    "MERGE_LG_NODE", "MERGE_DG_NODE", "REPARENT_DG_NODE",
    "REPARENT_LG_NODE", "DEPRECATE_NODE", "ADD_SOURCE_AND_CANONICAL_ADDRES",
    "MOVE_AID", "CREATE_LG"
]

VALID_USECASES = [
    "CC_AB_Campus_Audit", "CC_Cheetah_ORD_AUDIT",
    "CC_ORD_Point_Corrections_EU", "CC_ORD_Point_Corrections_NA",
    "CC_ORD_Point_Corrections_Rocket_Stations_NA", "CC_SIMS_Audit",
    "DSP_Hierarchy_Building_P0", "DSP_Hierarchy_Building_P1",
    "DSP_Hierarchy_Unit_P0", "CC_River-DSP_AUDIT"
]

SOURCE_SHEETS = {
    "CAMPUS": {"address_col": "Campus Address", "id_col": "Campus ID(SOURCE)"},
    "BUILDING": {"address_col": "Building Address", "id_col": "BPID(SOURCE)"},
    "UNIT": {"address_col": "Unit Address", "id_col": "Unit PID(SOURCE)"},
    "AID": {"address_col": "AID Address", "id_col": "AID"}
}

REVIEWER_COLS = ['Reviewer usecase', 'Reviewer alias', ' Date',
                 'Reviewer Verdict  ', 'Reviewer comments']

SKIP_COLS = REVIEWER_COLS + ['COMMENTS']

MAX_FILE_SIZE_MB = 1

# ============================================================
# HELPER FUNCTIONS
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


def check_drag_down(df, sheet_name, skip_cols=None):
    """
    Detect drag-down/fill-down errors where values increment by +1.
    Returns list of column names with sequential patterns.
    """
    if skip_cols is None:
        skip_cols = SKIP_COLS
    
    flagged_columns = []
    check_columns = [col for col in df.columns if col not in skip_cols]
    
    for col in check_columns:
        values = df[col].dropna().reset_index(drop=True)
        
        if len(values) < 4:
            continue
        
        consecutive_count = 0
        detected = False
        
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
                # Case 2: Text ending with numbers
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
                    detected = True
                    break
            else:
                consecutive_count = 0
        
    return flagged_columns


def is_valid_date(val):
    """Check if date is in M/DD/YYYY or MM/DD/YYYY format."""
    if pd.isna(val):
        return False
    
    s = str(val).strip()
    
    # If pandas parsed as Timestamp, format it
    if isinstance(val, pd.Timestamp):
        s = val.strftime("%-m/%d/%Y")  # This may vary by OS
        # Fallback: just check the original
    
    # Must contain exactly 2 slashes
    if s.count('/') != 2:
        return False
    
    parts = s.split('/')
    if len(parts) != 3:
        return False
    
    # Month: 1-2 digits, Day: 1-2 digits, Year: 4 digits
    if len(parts[0]) < 1 or len(parts[0]) > 2:
        return False
    if len(parts[1]) < 1 or len(parts[1]) > 2:
        return False
    if len(parts[2]) != 4:
        return False
    
    try:
        m = int(parts[0])
        d = int(parts[1])
        y = int(parts[2])
    except ValueError:
        return False
    
    # Month 1-12 (catches 23/7/2026 because 23 > 12)
    if m < 1 or m > 12:
        return False
    if d < 1 or d > 31:
        return False
    if y < 2020 or y > 2030:
        return False
    
    return True


def is_alpha_only(val):
    """Check if value contains only alphabetical characters."""
    if pd.isna(val):
        return False
    s = str(val).strip()
    if s == "":
        return False
    return s.isalpha()


# ============================================================
# MAIN VALIDATION FUNCTION
# ============================================================

def run_sanity_check(uploaded_file):
    """Run all sanity checks and return results."""
    
    results = []  # List of dicts: {status, sheet, check, details}
    
    # ----------------------------------------------------------
    # CHECK 1: File Size
    # ----------------------------------------------------------
    file_size_mb = uploaded_file.size / (1024 * 1024)
    if file_size_mb > MAX_FILE_SIZE_MB:
        results.append({
            "status": "ERROR",
            "sheet": "FILE",
            "check": "File Size",
            "details": f"File is {file_size_mb:.2f} MB (limit: {MAX_FILE_SIZE_MB} MB)"
        })
    else:
        results.append({
            "status": "PASS",
            "sheet": "FILE",
            "check": "File Size",
            "details": f"File size: {file_size_mb:.2f} MB"
        })
    
    # ----------------------------------------------------------
    # Read all sheets
    # ----------------------------------------------------------
    try:
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names
    except Exception as e:
        results.append({
            "status": "ERROR",
            "sheet": "FILE",
            "check": "File Read",
            "details": f"Cannot read file: {str(e)}"
        })
        return results
    
    # ----------------------------------------------------------
    # CHECK 2: Sheet Structure
    # ----------------------------------------------------------
    missing_sheets = [s for s in REQUIRED_SHEETS if s not in sheet_names]
    extra_sheets = [s for s in sheet_names if s not in REQUIRED_SHEETS]
    
    if missing_sheets:
        results.append({
            "status": "ERROR",
            "sheet": "SHEETS",
            "check": "Missing Sheets",
            "details": f"MISSING: {', '.join(missing_sheets)}"
        })
    
    if extra_sheets:
        results.append({
            "status": "WARNING",
            "sheet": "SHEETS",
            "check": "Extra Sheets",
            "details": f"EXTRA: {', '.join(extra_sheets)}"
        })
    
    if not missing_sheets and not extra_sheets:
        results.append({
            "status": "PASS",
            "sheet": "SHEETS",
            "check": "Sheet Structure",
            "details": "All 20 required sheets present, no extra"
        })
    
    # ----------------------------------------------------------
    # CHECK 3-9: Source Sheet Validations
    # ----------------------------------------------------------
    for sheet_name, config in SOURCE_SHEETS.items():
        if sheet_name not in sheet_names:
            results.append({
                "status": "ERROR",
                "sheet": sheet_name,
                "check": "Sheet Missing",
                "details": "Sheet not found in workbook"
            })
            continue
        
        try:
            df = pd.read_excel(xl, sheet_name=sheet_name, dtype=str)
        except Exception as e:
            results.append({
                "status": "ERROR",
                "sheet": sheet_name,
                "check": "Read Error",
                "details": f"Cannot read sheet: {str(e)}"
            })
            continue
        
        if len(df) == 0:
            results.append({
                "status": "PASS",
                "sheet": sheet_name,
                "check": "All Checks",
                "details": "Sheet is empty (no data rows) - skipped"
            })
            continue
        
        # --- CHECK 3: Mandatory Columns ---
        mandatory_cols = ['Usecase', 'Auditor', 'Audit Date', 'INPUT BPID']
        mandatory_errors = []
        
        for col in mandatory_cols:
            if col in df.columns:
                blank_count = df[col].isna().sum() + (df[col].astype(str).str.strip() == '').sum()
                # Subtract rows where NaN counted twice
                blank_count = df[col].apply(
                    lambda x: pd.isna(x) or str(x).strip() == '' or str(x).strip().lower() == 'nan'
                ).sum()
                if blank_count > 0:
                    mandatory_errors.append(f"{col}: {blank_count} blank")
            else:
                mandatory_errors.append(f"{col}: column not found")
        
        if mandatory_errors:
            for err in mandatory_errors:
                results.append({
                    "status": "ERROR",
                    "sheet": sheet_name,
                    "check": f"Mandatory - {err.split(':')[0]}",
                    "details": err
                })
        else:
            results.append({
                "status": "PASS",
                "sheet": sheet_name,
                "check": "Mandatory Columns",
                "details": "All mandatory columns filled"
            })
        
        # --- CHECK 4: Blank Rows (cols F onwards, excl COMMENTS) ---
        check_cols = [c for c in df.columns if c not in SKIP_COLS]
        if check_cols:
            blank_rows = []
            for idx, row in df.iterrows():
                is_blank = True
                for col in check_cols:
                    val = row.get(col, None)
                    if not pd.isna(val) and str(val).strip() != '' and str(val).strip().lower() != 'nan':
                        is_blank = False
                        break
                if is_blank:
                    blank_rows.append(idx + 2)  # +2 for header + 0-index
            
            if blank_rows:
                display_rows = blank_rows[:5]
                results.append({
                    "status": "ERROR",
                    "sheet": sheet_name,
                    "check": "Blank Rows",
                    "details": f"{len(blank_rows)} blank row(s) at row(s): {', '.join(map(str, display_rows))}"
                })
            else:
                results.append({
                    "status": "PASS",
                    "sheet": sheet_name,
                    "check": "Blank Rows",
                    "details": "No blank rows found"
                })
        
        # --- CHECK 5: Usecase Validation ---
        if 'Usecase' in df.columns:
            usecase_vals = df['Usecase'].dropna()
            usecase_vals = usecase_vals[usecase_vals.astype(str).str.strip() != '']
            usecase_vals = usecase_vals[usecase_vals.astype(str).str.strip().str.lower() != 'nan']
            
            invalid_uc = usecase_vals[~usecase_vals.isin(VALID_USECASES)]
            if len(invalid_uc) > 0:
                examples = invalid_uc.head(3).tolist()
                results.append({
                    "status": "ERROR",
                    "sheet": sheet_name,
                    "check": "Usecase Validation",
                    "details": f"{len(invalid_uc)} invalid value(s). Examples: {', '.join(map(str, examples))}"
                })
            else:
                results.append({
                    "status": "PASS",
                    "sheet": sheet_name,
                    "check": "Usecase Validation",
                    "details": "All usecase values from approved list"
                })
        
        # --- CHECK 6: Auditor Validation (alpha only) ---
        if 'Auditor' in df.columns:
            auditor_vals = df['Auditor'].dropna()
            auditor_vals = auditor_vals[auditor_vals.astype(str).str.strip() != '']
            auditor_vals = auditor_vals[auditor_vals.astype(str).str.strip().str.lower() != 'nan']
            
            invalid_aud = auditor_vals[~auditor_vals.astype(str).str.strip().str.isalpha()]
            if len(invalid_aud) > 0:
                examples = invalid_aud.head(3).tolist()
                results.append({
                    "status": "ERROR",
                    "sheet": sheet_name,
                    "check": "Auditor Validation",
                    "details": f"{len(invalid_aud)} invalid value(s). Must be alpha only. Examples: {', '.join(map(str, examples))}"
                })
            else:
                results.append({
                    "status": "PASS",
                    "sheet": sheet_name,
                    "check": "Auditor Validation",
                    "details": "All auditor aliases valid (alpha only)"
                })
        
        # --- CHECK 7: Audit Date Format (M/DD/YYYY or MM/DD/YYYY) ---
        if 'Audit Date' in df.columns:
            date_vals = df['Audit Date'].dropna()
            date_vals = date_vals[date_vals.astype(str).str.strip() != '']
            date_vals = date_vals[date_vals.astype(str).str.strip().str.lower() != 'nan']
            
            invalid_dates = []
            for idx, val in date_vals.items():
                if not is_valid_date(val):
                    invalid_dates.append(f"Row {idx + 2} ('{val}')")
            
            if invalid_dates:
                display_dates = invalid_dates[:5]
                results.append({
                    "status": "ERROR",
                    "sheet": sheet_name,
                    "check": "Audit Date Format",
                    "details": f"{len(invalid_dates)} invalid date(s). Must be M/DD/YYYY or MM/DD/YYYY. At: {', '.join(display_dates)}"
                })
            else:
                results.append({
                    "status": "PASS",
                    "sheet": sheet_name,
                    "check": "Audit Date Format",
                    "details": "All dates valid"
                })
        
        # --- CHECK 8: Duplicate Detection (Address + ID) ---
        addr_col = config["address_col"]
        id_col = config["id_col"]
        
        if addr_col in df.columns and id_col in df.columns:
            dup_df = df[[addr_col, id_col]].dropna(how='all')
            dup_df = dup_df[
                (dup_df[addr_col].astype(str).str.strip() != '') & 
                (dup_df[id_col].astype(str).str.strip() != '')
            ]
            duplicates = dup_df[dup_df.duplicated(keep='first')]
            
            if len(duplicates) > 0:
                results.append({
                    "status": "WARNING",
                    "sheet": sheet_name,
                    "check": "Duplicate Rows",
                    "details": f"{len(duplicates)} duplicate(s) based on [{addr_col} + {id_col}]"
                })
            else:
                results.append({
                    "status": "PASS",
                    "sheet": sheet_name,
                    "check": "Duplicate Rows",
                    "details": f"No duplicates on [{addr_col} + {id_col}]"
                })
        
        # --- CHECK 9: Drag-Down Detection ---
        drag_cols = check_drag_down(df, sheet_name, SKIP_COLS)
        
        if drag_cols:
            results.append({
                "status": "ERROR",
                "sheet": sheet_name,
                "check": "Drag-Down Detected",
                "details": f"Values incrementing by +1 in: {', '.join(drag_cols)} (likely fill-down error)"
            })
        else:
            results.append({
                "status": "PASS",
                "sheet": sheet_name,
                "check": "Drag-Down Check",
                "details": "No sequential +1 patterns detected"
            })
    
    # ----------------------------------------------------------
    # CHECK 10: Command Sheet Validation
    # ----------------------------------------------------------
    # Collect all COMMAND values from source sheets
    all_commands = set()
    for sheet_name_src in SOURCE_SHEETS.keys():
        if sheet_name_src in sheet_names:
            try:
                df_src = pd.read_excel(xl, sheet_name=sheet_name_src, dtype=str)
                if 'COMMAND' in df_src.columns:
                    cmds = df_src['COMMAND'].dropna()
                    cmds = cmds[cmds.astype(str).str.strip() != '']
                    cmds = cmds[cmds.astype(str).str.strip().str.lower() != 'nan']
                    for cmd in cmds:
                        all_commands.add(str(cmd).strip().upper())
            except Exception:
                pass
    
    command_sheets = [s for s in REQUIRED_SHEETS if s not in 
                      ["CASE DETAILS", "CAMPUS", "BUILDING", "UNIT", "AID"]]
    
    for cmd_sheet in command_sheets:
        if cmd_sheet not in sheet_names:
            continue
        
        # Check if this command is used in source sheets
        command_found = False
        for cmd in all_commands:
            if cmd_sheet.upper() in cmd or cmd in cmd_sheet.upper():
                command_found = True
                break
        
        if command_found:
            try:
                df_cmd = pd.read_excel(xl, sheet_name=cmd_sheet, dtype=str)
                if len(df_cmd) == 0:
                    results.append({
                        "status": "ERROR",
                        "sheet": cmd_sheet,
                        "check": "Command Sheet",
                        "details": "EMPTY but command exists in source sheets"
                    })
                else:
                    results.append({
                        "status": "PASS",
                        "sheet": cmd_sheet,
                        "check": "Command Sheet",
                        "details": f"Has data ({len(df_cmd)} rows)"
                    })
            except Exception:
                results.append({
                    "status": "ERROR",
                    "sheet": cmd_sheet,
                    "check": "Command Sheet",
                    "details": "Cannot read sheet"
                })
        else:
            results.append({
                "status": "PASS",
                "sheet": cmd_sheet,
                "check": "Command Sheet",
                "details": "Skipped (command not used in source)"
            })
    
    return results


# ============================================================
# STREAMLIT UI
# ============================================================

# Header
st.markdown("""
<div style="background-color: #232F3E; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
    <h2 style="color: #FF9900; margin: 0; text-align: center;">
        LMAQ | Last Mile Analytics & Quality
    </h2>
</div>
""", unsafe_allow_html=True)

st.title("CC_AMT_PIPELINE Sanity Checker")
st.caption("Upload your Excel file to validate before WorkDocs upload. Your file is NEVER modified.")

# File uploader
uploaded_file = st.file_uploader(
    "Select your CC_AMT_PIPELINE Excel file",
    type=["xlsx", "xls"],
    help="Upload the pipeline Excel file to run sanity checks"
)

if uploaded_file is not None:
    
    st.markdown("---")
    
    # Run validation
    with st.spinner("🔍 Running sanity checks..."):
        start_time = datetime.now()
        results = run_sanity_check(uploaded_file)
        elapsed = (datetime.now() - start_time).total_seconds()
    
    # Count results
    error_count = sum(1 for r in results if r["status"] == "ERROR")
    warning_count = sum(1 for r in results if r["status"] == "WARNING")
    pass_count = sum(1 for r in results if r["status"] == "PASS")
    
    # Summary banner
    if error_count == 0:
        st.success(f"✅ **ALL CHECKS PASSED!** | Errors: 0 | Warnings: {warning_count} | Passed: {pass_count} | Time: {elapsed:.1f}s")
        st.balloons()
    else:
        st.error(f"❌ **VALIDATION FAILED** | Errors: {error_count} | Warnings: {warning_count} | Passed: {pass_count} | Time: {elapsed:.1f}s")
    
    # Metrics row
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("❌ Errors", error_count)
    col2.metric("⚠️ Warnings", warning_count)
    col3.metric("✅ Passed", pass_count)
    col4.metric("⏱️ Time", f"{elapsed:.1f}s")
    
    st.markdown("---")
    
    # Detailed results
    st.subheader("📋 Detailed Results")
    
    # Filter options
    filter_option = st.radio(
        "Show:",
        ["All", "Errors Only", "Warnings Only", "Passed Only"],
        horizontal=True
    )
    
    # Filter results
    if filter_option == "Errors Only":
        display_results = [r for r in results if r["status"] == "ERROR"]
    elif filter_option == "Warnings Only":
        display_results = [r for r in results if r["status"] == "WARNING"]
    elif filter_option == "Passed Only":
        display_results = [r for r in results if r["status"] == "PASS"]
    else:
        display_results = results
    
    # Display results
    for r in display_results:
        status = r["status"]
        sheet = r["sheet"]
        check = r["check"]
        details = r["details"]
        
        if status == "ERROR":
            st.markdown(f"🔴 **ERROR** | `{sheet}` | **{check}** | {details}")
        elif status == "WARNING":
            st.markdown(f"🟡 **WARNING** | `{sheet}` | **{check}** | {details}")
        else:
            st.markdown(f"🟢 **PASS** | `{sheet}` | **{check}** | {details}")
    
    # Results table
    st.markdown("---")
    st.subheader("📊 Results Table")
    
    df_results = pd.DataFrame(results)
    
    # Color the status column
    def color_status(val):
        if val == "ERROR":
            return "background-color: #FFE6E6; color: #CC0000; font-weight: bold"
        elif val == "WARNING":
            return "background-color: #FFF5DC; color: #B46400; font-weight: bold"
        else:
            return "background-color: #DCFFDC; color: #008200; font-weight: bold"
    
    styled_df = df_results.style.applymap(color_status, subset=["status"])
    st.dataframe(styled_df, use_container_width=True, hide_index=True)
    
    # File info
    st.markdown("---")
    st.caption(f"📁 File: {uploaded_file.name} | Validated: {datetime.now().strftime('%d-%b-%Y %H:%M:%S')} | LMAQ Sanity Checker v2.5")

else:
    # Instructions when no file uploaded
    st.markdown("---")
    st.subheader("📝 Validation Checks Performed")
    
    checks_info = """
    | # | Check | Description |
    |---|-------|-------------|
    | 1 | File Size | Must be < 1 MB |
    | 2 | Sheet Structure | All 20 required sheets present, flags extra |
    | 3 | Mandatory Columns | Usecase, Auditor, Audit Date, INPUT BPID filled |
    | 4 | Blank Rows | No blank rows (excl. Reviewer cols & COMMENTS) |
    | 5 | Usecase | Must be from 10 approved values |
    | 6 | Auditor | Must be alpha only (a-z alias) |
    | 7 | Audit Date | Must be M/DD/YYYY or MM/DD/YYYY format |
    | 8 | Duplicates | Address + ID combination must be unique |
    | 9 | Drag-Down | Flags columns with values incrementing by +1 |
    | 10 | Command Sheets | If command used, command sheet must have data |
    """
    st.markdown(checks_info)
    
    st.info("💡 Checks 3-9 run on ALL 4 source sheets: CAMPUS, BUILDING, UNIT, AID")
    
    st.markdown("---")
    st.caption("LMAQ Sanity Checker v2.5 | Last Mile Analytics & Quality | Your file is NEVER modified")

