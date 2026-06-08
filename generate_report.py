import pandas as pd

def main():
    print("STEP 1: LOAD FILES")

    try:
        eaid_df = pd.read_excel(
            "EAID AT 7 May for Ops team.xlsx",
            header=2
        )

        dmm_template_df = pd.read_excel(
            "DMM Ops Report_1-29 April (as on 30 April).xlsx",
            sheet_name="Data"
        )

        rbr_df = pd.read_excel(
            "MTD Revenue_by_Resource_11-05-2026.xlsx"
        )

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
        return

    print("STEP 2: FILTER EAID FOR DMM")

    eaid_dmm = eaid_df[
        eaid_df["L4"]
        .astype(str)
        .str.strip()
        .eq("DATA MODERNIZATION & MIGRATION")
    ].copy()

    print("STEP 3: FILTER RBR FOR ENG")

    rbr_eng = rbr_df[
        rbr_df["Employee Subfunction"]
        .astype(str)
        .str.strip()
        .eq("ENG")
    ].copy()

    print("STEP 4: BUILD EMPLOYEE SECTION")

    employee_cols = [
        "Employee ID",
        "Employee Name",
        "Employee Email",
        "L4",
        "DOJ",
        "Location",
        "Designation",
        "Category"
    ]

    employee_section = eaid_dmm[employee_cols].copy()

    employee_section = employee_section.rename(columns={
        "Location": "Office Location"
    })

    employee_section["L4"] = "DMM"

    employee_section = employee_section.reset_index(drop=True)

    print("STEP 5: TAKE RBR A:CN")

    rbr_cn = rbr_eng.iloc[:, :92].copy()
    rbr_cn = rbr_cn.reset_index(drop=True)

    print("STEP 6: REMOVE DUPLICATE EMPLOYEE COLUMNS FROM RBR")

    final_employee_cols = [
        "Employee ID",
        "Employee Name",
        "Employee Email",
        "L4",
        "DOJ",
        "Office Location",
        "Designation",
        "Category"
    ]

    cols_to_drop = [
        col
        for col in rbr_cn.columns
        if col in final_employee_cols
        or col == "Location"
    ]

    rbr_cn = rbr_cn.drop(
        columns=cols_to_drop,
        errors="ignore"
    )

    print("STEP 7: KEEP LAST TWO TEMPLATE COLUMNS")

    template_last_cols = (
        dmm_template_df
        .iloc[:, 92:]
        .copy()
        .reset_index(drop=True)
    )

    print("STEP 8: ALIGN ROW COUNTS FOR PASTE OPERATION")

    max_rows = max(
        len(employee_section),
        len(rbr_cn)
    )

    employee_section = employee_section.reindex(range(max_rows))
    rbr_cn = rbr_cn.reindex(range(max_rows))
    template_last_cols = template_last_cols.reindex(range(max_rows))

    print("STEP 9: CREATE FINAL REPORT")

    final_report = pd.concat(
        [
            employee_section.reset_index(drop=True),
            rbr_cn.reset_index(drop=True),
            template_last_cols.reset_index(drop=True)
        ],
        axis=1
    )

    print("STEP 10: VALIDATIONS")

    print(f"DMM Employees: {len(employee_section)}")
    print(f"ENG Rows: {len(rbr_cn)}")
    print(f"Final Rows: {len(final_report)}")
    print(f"Final Columns: {len(final_report.columns)}")

    print("STEP 11: EXPORT")

    output_filename = "final_report.xlsx"

    with pd.ExcelWriter(
        output_filename,
        engine="openpyxl"
    ) as writer:
        final_report.to_excel(
            writer,
            sheet_name="Final Report",
            index=False
        )

    print("STEP 12: SUCCESS")

    print(f"Report generated: {output_filename}")
    print(f"Rows: {final_report.shape[0]}")
    print(f"Columns: {final_report.shape[1]}")

if __name__ == "__main__":
    main()
