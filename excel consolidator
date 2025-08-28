import os
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter


def normalize_header(header):
    """Normalize header for comparison (case-insensitive, remove dots/spaces)."""
    if header is None:
        return ""
    return str(header).strip().lower().replace(".", "").replace(" ", "")


def read_excel_with_hidden(file, include_hidden):
    """
    Reads Excel file and extracts only the contiguous table starting from A1.
    Optionally excludes hidden rows/columns.
    Returns list of (df, sheet_name).
    """
    wb = openpyxl.load_workbook(file, data_only=True)
    results = []

    for sheet in wb.sheetnames:
        ws = wb[sheet]

        max_row = ws.max_row
        max_col = ws.max_column

        # Determine bottom row of table starting from A1 (contiguous block)
        data_end_row = 1
        for r in range(1, max_row + 1):
            if any(ws.cell(row=r, column=c).value is not None for c in range(1, max_col + 1)):
                data_end_row = r
            else:
                break

        # Determine right-most col of table starting from A1 (contiguous block)
        data_end_col = 1
        for c in range(1, max_col + 1):
            if any(ws.cell(row=r, column=c).value is not None for r in range(1, data_end_row + 1)):
                data_end_col = c
            else:
                break

        data = []
        for r in range(1, data_end_row + 1):
            row = [ws.cell(row=r, column=c).value for c in range(1, data_end_col + 1)]
            data.append(row)

        if len(data) < 2:
            continue  # Not enough data for table

        df = pd.DataFrame(data[1:], columns=data[0])

        if not include_hidden:
            # Hidden cols
            hidden_cols = {
                get_column_letter(col[0].column)
                for col in ws.iter_cols(min_row=1, max_col=data_end_col, max_row=1)
                if ws.column_dimensions[get_column_letter(col[0].column)].hidden
            }

            # Hidden rows
            hidden_rows = {
                row[0].row
                for row in ws.iter_rows(min_row=2, max_row=data_end_row)
                if ws.row_dimensions[row[0].row].hidden
            }

            # Remove hidden cols
            visible_cols = []
            for idx, col_name in enumerate(df.columns):
                col_letter = get_column_letter(idx + 1)
                if col_letter not in hidden_cols:
                    visible_cols.append(col_name)
            df = df[visible_cols]

            # Remove hidden rows
            hidden_row_indexes = [i for i in range(len(df)) if i + 2 in hidden_rows]
            df = df.drop(index=hidden_row_indexes, errors='ignore')

        results.append((df, sheet))

    return results


def consolidate_excels(folder_path, include_hidden, match_identical_only):
    all_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith((".xlsx", ".xls"))]
    consolidated_df = pd.DataFrame()
    progress_text = st.empty()
    warnings = []

    # Step 1: Collect column occurrences across all sheets (not just per file)
    column_counts = {}
    file_sheet_columns = []

    for file in all_files:
        sheet_dataframes = read_excel_with_hidden(file, include_hidden)

        for df, sheet_name in sheet_dataframes:
            if df.empty:
                continue

            cols_lower = set([normalize_header(col) for col in df.columns])
            file_sheet_columns.append((file, sheet_name, cols_lower))

            for col in cols_lower:
                column_counts[col] = column_counts.get(col, 0) + 1

    # Decide which columns to consolidate
    if match_identical_only:
        # Include columns that appear in at least 2 sheets (across any file)
        common_columns_lower = {col for col, count in column_counts.items() if count >= 2}
    else:
        # Include union of all columns
        common_columns_lower = set(column_counts.keys())

    # Step 2: Consolidate data
    for file, sheet_name, sheet_cols in file_sheet_columns:
        progress_text.text(f"Processing: {os.path.basename(file)} > {sheet_name} ...")
        sheet_dataframes = read_excel_with_hidden(file, include_hidden)

        for df, sname in sheet_dataframes:
            if sname != sheet_name or df.empty:
                continue

            df.insert(0, "Sheet Name", sname)
            df.insert(0, "Source File", os.path.basename(file))

            if match_identical_only:
                # Normalize df columns for lookup
                df_cols_lower_map = {normalize_header(c): c for c in df.columns}

                filtered_cols = [df_cols_lower_map[col] for col in common_columns_lower if col in df_cols_lower_map]

                # Always add Source File and Sheet Name
                for special_col in ["Source File", "Sheet Name"]:
                    if special_col not in filtered_cols and special_col in df.columns:
                        filtered_cols.insert(0, special_col)

                if len(filtered_cols) <= 2:  # Means only Source File and Sheet Name or less
                    warnings.append(
                        f"⚠️ Skipped '{os.path.basename(file)}' > Sheet '{sname}': no common columns to consolidate."
                    )
                    continue

                df_filtered = df[filtered_cols]
                consolidated_df = pd.concat([consolidated_df, df_filtered], ignore_index=True)

            else:
                # Union of all columns
                consolidated_df = pd.concat([consolidated_df, df], ignore_index=True, sort=False)

    progress_text.text("✅ Consolidation Completed Successfully!")
    return consolidated_df, warnings


def main():
    st.title("📊 Excel Consolidation Tool")

    folder_path = st.text_input("📂 Enter folder path containing Excel files:")

    include_hidden = st.checkbox("Include hidden rows & columns?", value=False)
    match_identical_only = st.checkbox("Consolidate only identical headers (columns present in multiple files)?", value=True)

    if st.button("Start Consolidation"):
        if not folder_path or not os.path.isdir(folder_path):
            st.error("Please enter a valid folder path.")
        else:
            result_df, warnings = consolidate_excels(folder_path, include_hidden, match_identical_only)

            if warnings:
                for warning in warnings:
                    st.warning(warning)

            if result_df.empty:
                st.error("No valid data was consolidated. Please check your files.")
            else:
                output_file = os.path.join(folder_path, "Consolidated_Output.xlsx")
                result_df.to_excel(output_file, index=False)

                st.success(f"Consolidation complete! File saved at: {output_file}")
                st.dataframe(result_df.head(50))


if __name__ == "__main__":
    main()
