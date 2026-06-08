import pandas as pd
import sys
from pathlib import Path


EAID_FILE    = "EAID AT 7 May for Ops team.xlsx"
DMM_OPS_FILE = "DMM Ops Report_1-29 April (as on 30 April).xlsx"
MTD_FILE     = "MTD Revenue_by_Resource_11-05-2026.xlsx"
OUTPUT_FILE  = "final_report.xlsx"

SUMMARY_SHEETS = [
    "1.1 Partner Summary",
    "1.2Partner Detailed Information",
    "1.3 RPH < 1000",
    "1.4 Discount>65%",
]

# Columns 93–94 in the original Data tab (beyond col CN=92) that must be preserved
KEPT_DATA_COLS = ["Practitioner L4", "Engg Participating Partner Function"]


def main():

    # ──────────────────────────────────────────────
    print("STEP 1: LOAD FILES")
    # ──────────────────────────────────────────────

    try:
        # EAID: real headers are on row 3 (0-indexed: header=2)
        eaid_df = pd.read_excel(EAID_FILE, sheet_name=0, header=2, dtype=str)
        eaid_df.columns = [str(c).strip() for c in eaid_df.columns]

        # DMM Ops template: read the Data tab only to capture its last 2 columns (beyond CN)
        dmm_data_df = pd.read_excel(DMM_OPS_FILE, sheet_name="Data", dtype=str)
        dmm_data_df.columns = [str(c).strip() for c in dmm_data_df.columns]

        # MTD Revenue by Resource
        rbr_df = pd.read_excel(MTD_FILE, sheet_name=0, dtype=str)
        rbr_df.columns = [str(c).strip() for c in rbr_df.columns]

        # Summary sheets (preserved as-is in the output)
        summary_sheets = {}
        for sheet in SUMMARY_SHEETS:
            try:
                summary_sheets[sheet] = pd.read_excel(
                    DMM_OPS_FILE, sheet_name=sheet, header=None
                )
            except Exception as e:
                print(f"  ⚠  Could not read summary sheet '{sheet}': {e}")

    except FileNotFoundError as e:
        print(f"Error loading files: {e}")
        sys.exit(1)

    # ──────────────────────────────────────────────
    print("STEP 2: FILTER EAID FOR DMM")
    # ──────────────────────────────────────────────

    eaid_dmm = eaid_df[
        eaid_df["L4"].str.upper().str.strip() == "DATA MODERNIZATION & MIGRATION"
    ].copy()
    print(f"  → {len(eaid_dmm)} DMM resources found in EAID.")

    # ──────────────────────────────────────────────
    print("STEP 3: FILTER RBR FOR ENG")
    # ──────────────────────────────────────────────

    # Exact column name in MTD file is "Employee sub-function"
    rbr_eng = rbr_df[
        rbr_df["Employee sub-function"].str.upper().str.strip() == "ENG"
    ].copy()
    print(f"  → {len(rbr_eng)} ENG rows found in MTD (of {len(rbr_df)} total).")

    # ──────────────────────────────────────────────
    print("STEP 4: BUILD ADMM TAB DATA")
    # ──────────────────────────────────────────────

    # Required EAID source columns
    required_eaid = ["Emp ID", "Emp Name", "Emp Email", "L4", "DOJ",
                     "Location", "Designation", "Category"]
    missing = [c for c in required_eaid if c not in eaid_dmm.columns]
    if missing:
        print(f"ERROR: EAID file is missing columns: {missing}")
        sys.exit(1)

    admm_df = eaid_dmm[required_eaid].rename(columns={
        "Emp ID":      "Employee ID",
        "Emp Name":    "Employee Name",
        "Emp Email":   "Email ID",          # ADMM header is "Email ID", not "Employee Email"
        "L4":          "L4",
        "DOJ":         "Date of Joining",   # ADMM header is "Date of Joining", not "DOJ"
        "Location":    "Office Location",
        "Designation": "Designation",
        "Category":    "Category",
    }).copy()

    # L4 is always "DMM" in the ADMM tab
    admm_df["L4"] = "DMM"

    # Insert blank RM Name column at position 4 (after Email ID, before Date of Joining)
    admm_df.insert(4, "RM Name", "")

    admm_df = admm_df[[
        "Employee ID", "Employee Name", "Email ID", "L4", "RM Name",
        "Date of Joining", "Office Location", "Designation", "Category",
    ]].reset_index(drop=True)

    # ──────────────────────────────────────────────
    print("STEP 5: TAKE RBR COLS A:CN (first 92 columns)")
    # ──────────────────────────────────────────────

    # CN is the 92nd column (1-based), so iloc[:, :92] in pandas
    rbr_cn = rbr_eng.iloc[:, :92].copy().reset_index(drop=True)

    # ──────────────────────────────────────────────
    print("STEP 6: KEEP LAST TWO TEMPLATE COLUMNS (beyond CN)")
    # ──────────────────────────────────────────────

    # The original Data tab has 94 cols; cols 93–94 are "Practitioner L4"
    # and "Engg Participating Partner Function". We preserve them as blank
    # columns appended to the MTD data (they contain no per-row data to carry over).
    for col in KEPT_DATA_COLS:
        rbr_cn[col] = ""

    # ──────────────────────────────────────────────
    print("STEP 7: VALIDATIONS")
    # ──────────────────────────────────────────────

    print(f"  DMM Employees (ADMM tab) : {len(admm_df)} rows, {len(admm_df.columns)} columns")
    print(f"  ENG Rows (Data tab)      : {len(rbr_cn)} rows, {len(rbr_cn.columns)} columns")

    # ──────────────────────────────────────────────
    print("STEP 8: EXPORT")
    # ──────────────────────────────────────────────

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        # Preserve the 4 summary/dashboard sheets from the original template
        for sheet_name, df in summary_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

        # ADMM tab: DMM resource list from EAID
        admm_df.to_excel(writer, sheet_name="ADMM", index=False)

        # Data tab: MTD ENG revenue rows (cols A–CN) + 2 appended structural cols
        rbr_cn.to_excel(writer, sheet_name="Data", index=False)

    # ──────────────────────────────────────────────
    print("STEP 9: SUCCESS")
    # ──────────────────────────────────────────────

    print(f"  Report generated : {OUTPUT_FILE}")
    print(f"  ADMM tab         : {admm_df.shape[0]} rows × {admm_df.shape[1]} columns")
    print(f"  Data tab         : {rbr_cn.shape[0]} rows × {rbr_cn.shape[1]} columns")


if __name__ == "__main__":
    main()
