# Google Sheets Manager

This Python script allows you to manage Google Sheets in various ways, such as appending rows from JSON data, overwriting data, creating new sheets, sorting sheets by columns, and exporting sheet data as JSON files. The script uses Google Sheets API v4 and requires client secret and token files for authentication.

## Features

- Append rows to an existing sheet from a JSON file or stdin
- Overwrite existing data in a sheet with data from a JSON file
- Create a new sheet and upload data from a JSON file
- Sort sheets by specified columns
- Export sheet data as a JSON file or output to stdout
- Automatically update sheet formatting (header filters, alternating row colors)
- Create new columns if the JSON data contains keys not present in the sheet
- Manage Google Sheets using the command line

## Requirements

- Python 3
  - `pip install google-auth google-auth-oauthlib google-api-python-client`
- A Google API client secret file (client_secret.json) for authentication

## Setup

### Use requirements.txt

To install the required packages using the requirements.txt file, follow these steps:

- Open a terminal or command prompt.
- Navigate to the directory containing the requirements.txt file and the script.
- Run the following command:

```bash
pip install -r requirements.txt
```

This will install the required packages listed in the requirements.txt file. Make sure you have Python 3.6 or higher installed on your machine, and that pip is also installed and up-to-date.

### Google Cloud Platform (client_secret.json)

- Go to the Google Cloud Console to create a new project or select an existing one.
- Click "Create credentials" and select "OAuth 2.0 Client ID".
- Choose "Desktop app" for the Application type, and click "Create".
- Click "Download" next to your newly created Client ID and save the client_secret.json file in the same directory as the script.

## Usage

To run the script, use the following command-line format:

```bash
python3 sheets_manager.py \
    --sheet-id <Google-Sheet-ID-or-URL> \
    --sheet-name <sheet-name> \
    --in-json <path-to-json-file> \
    --out-json <path-to-json-file> \
    --overwrite \
    --new-sheet \
    --sort-columns <column-name-1> <column-name-2> ...
```

### Arguments

#### Required:

- `--sheet-id` (required): Google Sheet ID or URL.
- `--sheet-name`: Name of the sheet in the Google Sheet (default: Sheet1).
- `--in-json`: Path to the JSON file containing the rows to append or '-' for stdin.

#### Optional:

- `--out-json`: Path to the output JSON file where the updated sheet data will be saved.
- `--overwrite`: Overwrite existing data in the sheet instead of appending.
- `--new-sheet`: Create a new sheet and upload data to that sheet.
- `--sort-columns`: List of columns to sort by.

### Examples

1. Append rows from a JSON file to an existing sheet:

```bash
python3 sheets_manager.py --sheet-id "your-sheet-id" --sheet-name "Sheet1" --in-json "input.json"
```

2. Overwrite existing data in a sheet with data from a JSON file:

```bash
python3 sheets_manager.py --sheet-id "your-sheet-id" --sheet-name "Sheet1" --in-json "input.json" --overwrite
```

3. Create a new sheet and upload data from a JSON file:

```bash
python3 sheets_manager.py --sheet-id "your-sheet-id" --sheet-name "NewSheet" --in-json "input.json" --new-sheet
```

4. Sort a sheet by specific columns:

```bash
python3 sheets_manager.py --sheet-id "your-sheet-id" --sheet-name "Sheet1" --sort-columns "Column1" "Column2"
```

5. Save the sheet data as a JSON file:

```bash
python3 sheets_manager.py --sheet-id "your-sheet-id" --sheet-name "Sheet1" --out-json "output.json"
```

6. Read rows from stdin and append to an existing sheet:

```bash
cat input.json | python3 sheets_manager.py --sheet-id "your-sheet-id" --sheet-name "Sheet1" --in-json "-"
```

7. Output the sheet data as JSON to stdout:

```bash
python3 sheets_manager.py --sheet-id "your-sheet-id" --sheet-name "Sheet1" --out-json "-"
```
