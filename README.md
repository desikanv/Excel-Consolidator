**Excel Consolidation Tool**

This tool provides an easy-to-use interface to consolidate data from multiple Excel files into one. It can optionally filter out hidden rows and columns, and it supports two modes of consolidation:
1. Identical Columns Mode: Only columns present in multiple files are included in the final consolidated dataset.
2. Union of All Columns Mode: All columns from every file are included, even if they are not identical across files.

**Features**
1. Consolidate multiple Excel files: Handles files with multiple sheets and multiple formats (XLSX).
2. Hidden Rows & Columns: Optionally includes or excludes hidden rows and columns from the consolidation.
3. Flexible Column Matching: Choose whether to consolidate only columns that appear in more than one sheet/file or include all columns.
4. Progress Updates: Real-time progress feedback during the consolidation process.
5. Warnings: Displays warnings when a sheet does not contain common columns when consolidating with identical headers.

**How It Works**
1. Reading Excel Files: The script reads all Excel files in the provided folder path. It supports .xlsx and .xls formats.
2. Processing Each Sheet: For each sheet in each file: It identifies the contiguous block of data starting from the top-left (A1). Optionally excludes hidden rows and columns based on user preference.
3. Column Matching: If the "Consolidate only identical headers" option is checked, only columns that are present across multiple sheets/files are included in the output. If unchecked, all columns are included, even if they are not shared across files.
4. Consolidation: The data from all sheets is merged into a single Pandas DataFrame.
5. Exporting: The consolidated data is saved as a new Excel file in the same directory as the source files.

**Streamlit Interface**
This tool is built with Streamlit, allowing for an interactive web-based interface.
Use: $ streamlit run excel_consolidator.py "OR" python -m streamlit run "_Insert File Location_" to run the file

