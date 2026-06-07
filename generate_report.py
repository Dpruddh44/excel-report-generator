import pandas as pd

def main():
    print("STEP 1: LOAD FILES")
    # Load files using the exact names provided in the requirement
    try:
        eaid_df = pd.read_excel("EAID AT 7 May for Ops team.xlsx", header=2)
        dmm_template_df = pd.read_excel("DMM Ops Report_1-29 April (as on 30 April).xlsx", sheet_name="Data")
        rbr_df = pd.read_excel("MTD Revenue_by_Resource_11-05-2026.xlsx")

        # Rename columns to match the standard business logic
        eaid_df = eaid_df.rename(columns={
            "Emp ID": "Employee ID",
            "Emp Name": "Employee Name",
            "Emp Email": "Employee Email"
        })

        rbr_df = rbr_df.rename(columns={
            "Personnel No./Employee code": "Employee ID",
            "Employee sub-function": "Employee Subfunction"
        })
    except FileNotFoundError as e:
        print(f"Error loading files: {e}")
        print("Please ensure the input files are in the same directory as the script with the exact names:")
        print("1. EAID AT 7 May for Ops team.xlsx")
        print("2. DMM Ops Report_1-29 April (as on 30 April).xlsx")
        print("3. MTD Revenue_by_Resource_11-05-2026.xlsx")
        return

    print("STEP 2: FILTER EAID")
    eaid_dmm = eaid_df[eaid_df["L4"] == "DATA MODERNIZATION & MIGRATION"].copy()

    print("STEP 3: FILTER RBR")
    rbr_eng = rbr_df[rbr_df["Employee Subfunction"] == "ENG"].copy()

    print("STEP 4: STANDARDIZE EMPLOYEE IDS")
    eaid_dmm["Employee ID"] = eaid_dmm["Employee ID"].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    rbr_eng["Employee ID"] = rbr_eng["Employee ID"].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

    print("STEP 5: FIND COMMON EMPLOYEES")
    common_ids = set(eaid_dmm["Employee ID"]) & set(rbr_eng["Employee ID"])
    eaid_dmm = eaid_dmm[eaid_dmm["Employee ID"].isin(common_ids)].copy()
    rbr_eng = rbr_eng[rbr_eng["Employee ID"].isin(common_ids)].copy()
    
    # Drop duplicates to ensure exactly one row per employee for concatenation
    eaid_dmm = eaid_dmm.drop_duplicates(subset=["Employee ID"]).copy()
    rbr_eng = rbr_eng.drop_duplicates(subset=["Employee ID"]).copy()

    print("STEP 6: BUILD EMPLOYEE MASTER SECTION")
    employee_cols = ["Employee ID", "Employee Name", "Employee Email", "L4", "DOJ", "Location", "Designation", "Category"]
    employee_section = eaid_dmm[employee_cols].copy()
    employee_section = employee_section.rename(columns={"Location": "Office Location"})

    print("STEP 7: PROCESS RBR DATA")
    rbr_cn = rbr_eng.iloc[:, :92].copy()

    print("STEP 9: ALIGN DATA")
    # Sort both datasets by Employee ID and reset index before removing duplicate columns to properly align
    employee_section = employee_section.sort_values("Employee ID").reset_index(drop=True)
    rbr_cn = rbr_cn.sort_values("Employee ID").reset_index(drop=True)

    print("STEP 8: REMOVE DUPLICATE EMPLOYEE COLUMNS FROM RBR")
    # The employee section already contains:
    final_employee_cols = ["Employee ID", "Employee Name", "Employee Email", "L4", "DOJ", "Office Location", "Designation", "Category"]
    
    # Identify and drop columns from RBR that are in the employee section to prevent _x / _y duplicates
    # Also removing "Location" just in case RBR has "Location" which corresponds to "Office Location"
    cols_to_drop = [col for col in rbr_cn.columns if col in final_employee_cols or col == "Location"]
    rbr_cn = rbr_cn.drop(columns=cols_to_drop)

    print("STEP 10: CREATE FINAL REPORT")
    # Dynamically identify the last two template columns after CN (index 92 onwards)
    template_last_cols = dmm_template_df.iloc[:, 92:].copy().reset_index(drop=True)
    print(f"Retained template columns: {list(template_last_cols.columns)}")
    
    # Horizontally combine employee section and RBR A:CN
    final_report = pd.concat([employee_section, rbr_cn], axis=1)
    
    # Append the retained template columns
    # Reindex ensures we preserve the exact final_report rows without adding rows if the template has more rows
    template_last_cols = template_last_cols.reindex(final_report.index)
    final_report = pd.concat([final_report, template_last_cols], axis=1)

    print("STEP 11: VALIDATIONS")
    total_dmm = len(eaid_df[eaid_df["L4"] == "DATA MODERNIZATION & MIGRATION"])
    total_eng = len(rbr_df[rbr_df["Employee Subfunction"] == "ENG"])
    total_common = len(common_ids)
    
    print(f"Total DMM employees: {total_dmm}")
    print(f"Total ENG employees: {total_eng}")
    print(f"Total common employees: {total_common}")
    print(f"Final report row count: {final_report.shape[0]}")
    print(f"Final report column count: {final_report.shape[1]}")

    if final_report.shape[0] != total_common:
        raise ValueError("Validation Error: Final rows do not match common employees count.")

    print("STEP 12: EXPORT")
    output_filename = "final_report.xlsx"
    with pd.ExcelWriter(output_filename, engine="openpyxl") as writer:
        final_report.to_excel(writer, sheet_name="Final Report", index=False)

    print("STEP 13: SUCCESS MESSAGE")
    print("Report generated successfully.")
    print(f"Rows: {final_report.shape[0]}")
    print(f"Columns: {final_report.shape[1]}")

if __name__ == "__main__":
    main()
