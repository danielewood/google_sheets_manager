#!/usr/bin/env python3
import argparse
import json
import os
import sys
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build as googleapiclient_build

CLIENT_SECRET_FILE = "client_secret.json"
TOKEN_FILE = "token.pickle"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def get_credentials():
    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "rb") as token:
            creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, "wb") as token:
            token.write(creds.to_json().encode())

    return creds


def write_sheet_to_json(sheet_id, sheet_name, headers, service, output_path):
    range_name = f"{sheet_name}!A1:{chr(ord('A') + len(headers) - 1)}"
    result = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=sheet_id, range=range_name)
        .execute()
    )
    values = result.get("values", [])

    if not values:
        print("No data found.")
    else:
        data = []
        for row in values[1:]:
            data.append({header: value for header, value in zip(headers, row)})

        if output_path == sys.stdout:
            json.dump(data, output_path, ensure_ascii=False, indent=2)
        else:
            with open(output_path, "w") as outfile:
                json.dump(data, outfile, ensure_ascii=False, indent=2)


def get_sheet_headers(sheet_id, sheet_name, service):
    range_name = f"{sheet_name}!1:1"
    result = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=sheet_id, range=range_name)
        .execute()
    )
    return result.get("values", [[]])[0]


def append_rows(sheet_id, sheet_name, rows, headers, service):
    data = []

    for row in rows:
        data.append([row.get(header, "") for header in headers])

    body = {"values": data}
    range_name = f"{sheet_name}!A:{chr(ord('A') + len(headers) - 1)}"
    result = (
        service.spreadsheets()
        .values()
        .append(
            spreadsheetId=sheet_id, range=range_name, valueInputOption="RAW", body=body
        )
        .execute()
    )
    print(f"{result.get('updates').get('updatedCells')} cells appended.")


def extract_sheet_id(url_or_id):
    if "https://docs.google.com/spreadsheets/d/" in url_or_id:
        start_index = url_or_id.find("https://docs.google.com/spreadsheets/d/") + len(
            "https://docs.google.com/spreadsheets/d/"
        )
        end_index = url_or_id.find("/", start_index)
        return url_or_id[start_index:end_index]
    else:
        return url_or_id


def sheet_exists(sheet_id, sheet_name, service):
    sheet_metadata = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    sheets = sheet_metadata.get("sheets", "")
    return any(s["properties"]["title"] == sheet_name for s in sheets)


def get_available_sheet_names(sheet_id, service):
    sheet_metadata = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    sheets = sheet_metadata.get("sheets", "")
    return [s["properties"]["title"] for s in sheets]


def add_new_columns(sheet_id, sheet_name, new_headers, service):
    headers_range = f"{sheet_name}!1:1"
    current_headers = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=sheet_id, range=headers_range)
        .execute()
        .get("values", [[]])[0]
    )

    updated_headers = current_headers + new_headers
    updated_headers_range = f"{sheet_name}!1:1"
    body = {
        "range": updated_headers_range,
        "majorDimension": "ROWS",
        "values": [updated_headers],
    }
    result = (
        service.spreadsheets()
        .values()
        .update(
            spreadsheetId=sheet_id,
            range=updated_headers_range,
            valueInputOption="RAW",
            body=body,
        )
        .execute()
    )


def update_sheet_formatting(sheet_id, sheet_name, headers, service):
    sheet_metadata = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    sheets = sheet_metadata.get("sheets", "")
    sheet = [s for s in sheets if s["properties"]["title"] == sheet_name][0]

    sheet_id_num = sheet["properties"]["sheetId"]

    set_filter_request = {
        "setBasicFilter": {
            "filter": {
                "range": {
                    "sheetId": sheet_id_num,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": len(headers),
                }
            }
        }
    }

    set_conditional_format_rules_request = {
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [
                    {
                        "sheetId": sheet_id_num,
                        "startRowIndex": 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": len(headers),
                    }
                ],
                "booleanRule": {
                    "condition": {
                        "type": "CUSTOM_FORMULA",
                        "values": [{"userEnteredValue": "=ISEVEN(ROW())"}],
                    },
                    "format": {
                        "backgroundColor": {
                            "red": 0.9,
                            "green": 0.9,
                            "blue": 0.9,
                        }
                    },
                },
            },
            "index": 0,
        }
    }

    body = {"requests": [set_filter_request, set_conditional_format_rules_request]}
    service.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body=body).execute()


def sort_sheet(sheet_id, sheet_name, sort_columns, service):
    sheet_metadata = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    sheets = sheet_metadata.get("sheets", "")
    sheet = [s for s in sheets if s["properties"]["title"] == sheet_name][0]

    sheet_id_num = sheet["properties"]["sheetId"]

    headers = get_sheet_headers(sheet_id, sheet_name, service)

    sort_specs = []
    for column_name in sort_columns:
        if column_name in headers:
            sort_specs.append(
                {
                    "dimensionIndex": headers.index(column_name),
                    "sortOrder": "ASCENDING",
                }
            )

    body = {
        "requests": [
            {
                "sortRange": {
                    "range": {
                        "sheetId": sheet_id_num,
                        "startRowIndex": 1,
                        "endRowIndex": sheet["properties"]["gridProperties"][
                            "rowCount"
                        ],
                        "startColumnIndex": 0,
                        "endColumnIndex": len(headers),
                    },
                    "sortSpecs": sort_specs,
                }
            }
        ]
    }

    service.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body=body).execute()


def column_name_to_index(column_name):
    index = 0
    for char in column_name:
        index = index * 26 + (ord(char.upper()) - ord("A")) + 1
    return index - 1


def delete_sheet(sheet_id, sheet_name, service):
    sheet_metadata = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    sheets = sheet_metadata.get("sheets", "")
    sheet = [s for s in sheets if s["properties"]["title"] == sheet_name][0]
    sheet_id_num = sheet["properties"]["sheetId"]
    sheet_index = sheet["properties"]["index"]

    body = {"requests": [{"deleteSheet": {"sheetId": sheet_id_num}}]}
    service.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body=body).execute()

    return sheet_index


def create_new_sheet(sheet_id, sheet_name, service, sheet_index=None):
    properties = {"title": sheet_name}
    if sheet_index is not None:
        properties["index"] = sheet_index

    body = {"requests": [{"addSheet": {"properties": properties}}]}
    service.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body=body).execute()


def handle_in_jsonput(args, sheet_id, sheet_name, service):
    if args.overwrite:
        sheet_index = delete_sheet(sheet_id, sheet_name, service)
        create_new_sheet(sheet_id, sheet_name, service, sheet_index=sheet_index)

    if args.in_json == "-":
        rows = json.load(sys.stdin)
    else:
        with open(args.in_json, "r") as json_file:
            rows = json.load(json_file)

    headers = get_sheet_headers(sheet_id, sheet_name, service)
    new_headers = []

    for row in rows:
        for key in row.keys():
            if key not in headers:
                headers.append(key)
                new_headers.append(key)

    if new_headers:
        add_new_columns(sheet_id, sheet_name, new_headers, service)

    append_rows(sheet_id, sheet_name, rows, headers, service)
    update_sheet_formatting(sheet_id, sheet_name, headers, service)


def parse_arguments():
    parser = argparse.ArgumentParser(
        description="Append rows to a Google Sheet from a JSON file and/or output the sheet data as a JSON file."
    )
    parser.add_argument("--sheet-id", required=True, help="Google Sheet ID or URL.")
    parser.add_argument(
        "--sheet-name",
        default="Sheet1",
        help="Name of the sheet in the Google Sheet (default: Sheet1).",
    )
    parser.add_argument(
        "--in-json",
        help="Path to the JSON file containing the rows to append or '-' for stdin.",
    )
    parser.add_argument("--sort-columns", nargs="*", help="List of columns to sort by.")
    parser.add_argument(
        "--out-json",
        help="Path to the output JSON file where the updated sheet data will be saved.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite existing data in the sheet instead of appending.",
    )
    parser.add_argument(
        "--new-sheet",
        action="store_true",
        help="Create a new sheet and upload data to that sheet.",
    )

    return parser.parse_args()


def main():
    args = parse_arguments()

    creds = get_credentials()
    service = googleapiclient_build("sheets", "v4", credentials=creds)

    sheet_id = extract_sheet_id(args.sheet_id)
    sheet_name = args.sheet_name

    if args.new_sheet:
        if sheet_exists(sheet_id, sheet_name, service):
            print(f"Sheet '{sheet_name}' already exists. Exiting.")
            sys.exit(1)
        else:
            create_new_sheet(sheet_id, sheet_name, service)

    elif not sheet_exists(sheet_id, sheet_name, service):
        available_sheet_names = get_available_sheet_names(sheet_id, service)
        print(f"Sheet '{sheet_name}' not found.")
        print(f"Available sheets: {', '.join(available_sheet_names)}")
        sys.exit(1)

    if args.in_json:
        handle_in_jsonput(args, sheet_id, sheet_name, service)

    if args.sort_columns:
        sort_sheet(sheet_id, sheet_name, args.sort_columns, service)

    if args.out_json:
        headers = get_sheet_headers(sheet_id, sheet_name, service)
        if args.out_json == "-":
            write_sheet_to_json(sheet_id, sheet_name, headers, service, sys.stdout)
        else:
            write_sheet_to_json(sheet_id, sheet_name, headers, service, args.out_json)


if __name__ == "__main__":
    main()
