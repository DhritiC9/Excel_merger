import io
import re
from pathlib import Path

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Excel Merger", page_icon="📊", layout="wide")

st.title("📊 Excel Merger")
st.write(
    "Combine Excel or CSV files into one consistently aligned Excel workbook. "
    "You can also append new files to a previously generated combined workbook."
)

SUPPORTED_TYPES = ["xlsx", "xls", "csv"]


def clean_header(value):
    """Create a comparison key without changing the displayed column name."""
    text = str(value).strip().lower()
    return re.sub(r"[^a-z0-9]", "", text)


def make_unique_headers(columns):
    """Avoid duplicate/blank headers while preserving readable original names."""
    result = []
    seen = {}
    for index, value in enumerate(columns, start=1):
        name = str(value).strip() if value is not None else ""
        if not name or name.lower().startswith("unnamed:"):
            name = f"Column {index}"

        key = clean_header(name)
        count = seen.get(key, 0) + 1
        seen[key] = count
        if count > 1:
            name = f"{name}_{count}"
        result.append(name)
    return result


def read_uploaded_file(uploaded_file):
    """Read every sheet from Excel files, or the single table from a CSV."""
    filename = uploaded_file.name
    suffix = Path(filename).suffix.lower()

    if suffix == ".csv":
        uploaded_file.seek(0)
        return {"CSV": pd.read_csv(uploaded_file)}

    if suffix in (".xlsx", ".xls"):
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file, sheet_name=None)

    raise ValueError(f"Unsupported file type: {filename}")


def standardize_headers(df, canonical_headers):
    """
    Match columns case-insensitively and regardless of spaces/punctuation.
    The first encountered spelling becomes the displayed column name.
    """
    df = df.copy()
    df.columns = make_unique_headers(df.columns)

    rename_map = {}
    for col in df.columns:
        key = clean_header(col)
        if key not in canonical_headers:
            canonical_headers[key] = col
        rename_map[col] = canonical_headers[key]

    df = df.rename(columns=rename_map)

    # If two source columns normalize to the same final name, combine their
    # non-empty values rather than silently dropping either column.
    if df.columns.duplicated().any():
        merged = {}
        for col in dict.fromkeys(df.columns):
            same_name = df.loc[:, df.columns == col]
            if same_name.shape[1] == 1:
                merged[col] = same_name.iloc[:, 0]
            else:
                merged[col] = same_name.bfill(axis=1).iloc[:, 0]
        df = pd.DataFrame(merged, index=df.index)

    return df


def workbook_to_dataframe(uploaded_file, selected_sheet=None):
    sheets = read_uploaded_file(uploaded_file)
    if selected_sheet is not None:
        if selected_sheet not in sheets:
            raise ValueError(
                f"Sheet '{selected_sheet}' was not found in {uploaded_file.name}."
            )
        sheets = {selected_sheet: sheets[selected_sheet]}

    frames = []
    for sheet_name, frame in sheets.items():
        if frame is None or frame.empty:
            continue
        frame = frame.dropna(how="all")
        if frame.empty:
            continue
        frames.append(frame)

    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True, sort=False)


def to_excel_bytes(df, sheet_name="Combined Data"):
    output = io.BytesIO()
    safe_sheet_name = sheet_name[:31] or "Combined Data"
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=safe_sheet_name)
        worksheet = writer.sheets[safe_sheet_name]
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        # Readable column widths, capped to avoid excessively wide columns.
        for column_cells in worksheet.columns:
            values = [str(cell.value) if cell.value is not None else "" for cell in column_cells[:200]]
            width = min(max(max((len(v) for v in values), default=10) + 2, 12), 45)
            worksheet.column_dimensions[column_cells[0].column_letter].width = width

    output.seek(0)
    return output.getvalue()


mode = st.radio(
    "What would you like to do?",
    ["Create a new combined Excel", "Append files to an existing combined Excel"],
    horizontal=True,
)

existing_df = pd.DataFrame()

if mode == "Append files to an existing combined Excel":
    st.subheader("1. Upload your existing combined workbook")
    existing_file = st.file_uploader(
        "Choose the Excel file you want to add rows to",
        type=["xlsx", "xls"],
        key="existing_workbook",
    )

    if existing_file is not None:
        try:
            existing_sheets = read_uploaded_file(existing_file)
            sheet_names = list(existing_sheets.keys())
            chosen_sheet = st.selectbox(
                "Choose the sheet to append data to",
                sheet_names,
                key="existing_sheet",
            )
            existing_df = existing_sheets[chosen_sheet].copy()
            st.success(
                f"Loaded '{existing_file.name}' — sheet '{chosen_sheet}' "
                f"contains {len(existing_df):,} rows."
            )
        except Exception as exc:
            st.error(f"Could not read the existing workbook: {exc}")

st.subheader("2. Upload the files to merge")
new_files = st.file_uploader(
    "Select one or more Excel or CSV files",
    type=SUPPORTED_TYPES,
    accept_multiple_files=True,
    key="new_files",
    help="There is no app-level file-count limit. Practical limits depend on your "
         "computer/server memory and Streamlit's configured per-file upload size.",
)

add_source_info = st.checkbox(
    "Add Source File and Source Sheet columns for traceability",
    value=False,
)

if new_files:
    st.caption(f"{len(new_files)} file(s) selected.")

if st.button("Merge files", type="primary", use_container_width=True):
    if not new_files and existing_df.empty:
        st.warning("Upload at least one source file to continue.")
    elif mode == "Append files to an existing combined Excel" and existing_df.empty:
        st.warning("Upload a valid existing combined Excel workbook first.")
    else:
        try:
            canonical_headers = {}
            frames_to_merge = []

            # Existing data is included first when append mode is selected.
            if not existing_df.empty:
                existing_copy = existing_df.dropna(how="all").copy()
                existing_copy = standardize_headers(existing_copy, canonical_headers)
                frames_to_merge.append(existing_copy)

            progress = st.progress(0, text="Reading files...")
            total = len(new_files)

            for index, uploaded_file in enumerate(new_files, start=1):
                sheets = read_uploaded_file(uploaded_file)

                for sheet_name, frame in sheets.items():
                    if frame is None or frame.empty:
                        continue
                    frame = frame.dropna(how="all")
                    if frame.empty:
                        continue

                    frame = standardize_headers(frame, canonical_headers)

                    if add_source_info:
                        frame["Source File"] = uploaded_file.name
                        frame["Source Sheet"] = str(sheet_name)

                    frames_to_merge.append(frame)

                progress.progress(
                    index / total,
                    text=f"Processed {index} of {total}: {uploaded_file.name}",
                )

            if not frames_to_merge:
                st.warning("The uploaded files did not contain any data rows.")
            else:
                # sort=False keeps the first-seen column order and adds any
                # newly encountered columns at the end. Missing values remain blank.
                combined_df = pd.concat(
                    frames_to_merge,
                    ignore_index=True,
                    sort=False,
                )
                combined_df = combined_df.dropna(how="all")

                st.session_state["merged_dataframe"] = combined_df
                st.session_state["merged_filename"] = (
                    "combined_data.xlsx"
                    if mode == "Create a new combined Excel"
                    else "combined_data_updated.xlsx"
                )
                st.success(
                    f"Merge complete: {len(combined_df):,} rows and "
                    f"{len(combined_df.columns):,} columns."
                )

        except Exception as exc:
            st.error(f"Something went wrong while merging: {exc}")

if "merged_dataframe" in st.session_state:
    result = st.session_state["merged_dataframe"]
    st.subheader("3. Review and download")
    col1, col2 = st.columns(2)
    col1.metric("Total rows", f"{len(result):,}")
    col2.metric("Total columns", f"{len(result.columns):,}")

    st.dataframe(result.head(100), use_container_width=True)
    st.caption("Preview shows up to the first 100 rows. The downloaded file contains all rows.")

    excel_data = to_excel_bytes(result)
    st.download_button(
        label="⬇️ Download combined Excel",
        data=excel_data,
        file_name=st.session_state["merged_filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

st.divider()
st.caption(
    "Column matching ignores differences in capitalization, spaces and punctuation "
    "(for example, 'Beneficiary Name' and 'beneficiary_name' are treated as the same "
    "column). Columns found in only some files are retained; their missing cells stay blank."
)
