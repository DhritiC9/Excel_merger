# Excel Merger Streamlit App

## Run locally

1. Install Python 3.10 or newer.
2. Open a terminal in this folder.
3. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

4. Start the app:

   ```bash
   streamlit run app.py
   ```

## How to use

- Choose **Create a new combined Excel** to merge a fresh batch.
- Choose **Append files to an existing combined Excel** to upload a previously
  generated workbook, select its target sheet, and add new rows.
- Upload as many `.xlsx`, `.xls`, or `.csv` files as your machine/server can handle.
- Download the combined workbook when the merge is complete.

## Column matching

Headers are matched without regard to capitalization, spaces, underscores, or
punctuation. For example, `Beneficiary Name`, `beneficiary_name`, and
`BENEFICIARY-NAME` are treated as the same column. New columns found in later
files are retained and added to the combined output; cells with no corresponding
value are left blank.

The app does not remove duplicate rows. It appends all source rows.

## Upload-size note

There is no hard-coded file-count limit in the app. Streamlit and the machine
running it still have practical upload-size and memory limits. To raise the
per-file limit, create `.streamlit/config.toml` beside `app.py`, for example:

```toml
[server]
maxUploadSize = 500
```

The value is in MB per file. Choose a value appropriate for your server.
