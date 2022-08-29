from typing import Dict, Any, List
import json
from dataclasses import dataclass
from datetime import datetime, date, timedelta

import pytz
import gspread

from docassemble.base.util import get_config, DADict
from oauth2client.service_account import ServiceAccountCredentials

__all__ = ["read_sheet", "GoogleSheetAppender", "AppendRow"]

credential_json = get_config("google", {}).get("service account credentials", None)
if credential_json is None:
    credential_info = None
else:
    credential_info = json.loads(credential_json, strict=False)


scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]


def read_sheet(sheet_name, worksheet_index=0):
    creds = ServiceAccountCredentials.from_json_keyfile_dict(credential_info, scope)
    client = gspread.authorize(creds)
    sheet = client.open(sheet_name).get_worksheet(worksheet_index)
    return sheet.get_all_records()


@dataclass
class AppendRow:
    sheet_id: int
    data: Dict[str, Any]


def _flatten_values(values):
    flat = []
    for field, value in values:
        if isinstance(value, DADict) or isinstance(value, dict):
            for sub_field, sub_value in value.items():
                flat.append((f"{field}[{sub_field}]", sub_value))
        else:
            flat.append((field, value))
    return flat


def _extended_value(value: Any, tz: Any):
    if isinstance(value, str):
        return { "userEnteredValue": {"stringValue": value}}
    elif isinstance(value, bool):
        return { "userEnteredValue": {"boolValue": value}}
    elif isinstance(value, (float, int)):
        return { "userEnteredValue": {"numberValue": value}}
    elif isinstance(value, datetime):
        if value.tzinfo:
            value = value.astimezone(tz).replace(tzinfo=None)
        epoch = datetime(1899, 12, 30)
        serial_number = (value - epoch) / timedelta(days=1)
        return { 
            "userEnteredFormat": {
                "numberFormat": {
                    "type": "DATE_TIME",
                }
            },
            "userEnteredValue": {"numberValue": serial_number}}
    elif isinstance(value, date):
        serial_number = (value - date(1899, 12, 30)).days
        return { 
            "userEnteredFormat": {
                "numberFormat": {
                    "type": "DATE",
                }
            },
            "userEnteredValue": {"numberValue": serial_number}}
    else:
        return { "userEnteredValue": {"stringValue": str(value)}}


def _add_columns_request(sheet_id, num_columns):
    return {
        "appendDimension": {
            "sheetId": sheet_id,
            "dimension": "COLUMNS",
            "length": num_columns,
        }
    }


def _append_row_request(sheet_id, values, tz):
    return {
        "appendCells": {
            "sheetId": sheet_id,
            "rows": [
                {
                    "values": [
                        _extended_value(value, tz)
                        for value in values
                    ]
                }
            ],
            "fields": "*",
        }
    }


def _update_headers_request(sheet_id, headers, tz):
    return {
        "updateCells": {
            "rows": [
                {
                    "values": [
                        _extended_value(header, tz)
                        for header in headers
                    ]
                }
            ],
            "start": {
                "sheetId": sheet_id,
                "rowIndex": 0,
                "columnIndex": 0,
            },
            "fields": "*",
        }
    }


class GoogleSheetAppender:
    def __init__(self):
        creds = ServiceAccountCredentials.from_json_keyfile_dict(credential_info, scope)
        self.client = gspread.authorize(creds)

    def append_batch(self, spreadsheet_key: str, batch: List[AppendRow]):
        """
        append_batch is a high level function that appends rows to a spreadsheet
        using as few requests as possible. It handles column ordering,
        adding columns, and multiple sheet updates.
        """

        ss = self.client.open_by_key(spreadsheet_key)

        sheet_meta = ss.fetch_sheet_metadata()
        sheet_tz_name = sheet_meta.get('properties', {}).get('timezone', 'America/Vancouver')
        sheet_tz = pytz.timezone(sheet_tz_name)

        sheet_names_by_id = {
            x["properties"]["sheetId"]: x["properties"]["title"]
            for x in sheet_meta["sheets"]
        }

        # unique sheet ids in the batch
        sheet_ids_to_update = list(set(row.sheet_id for row in batch))

        # unique sheet titles in the batch
        sheet_titles_to_update = [sheet_names_by_id[id] for id in sheet_ids_to_update]

        # Fetch header rows
        header_row_result = ss.values_batch_get(
            ranges=[f"{sheet_title}!1:1" for sheet_title in sheet_titles_to_update]
        )

        # Map sheet_id -> header list
        headers_by_sheet_id = {
            sheet_id: value_range.get("values", [[]])[0]
            for sheet_id, value_range in zip(
                sheet_ids_to_update, header_row_result["valueRanges"]
            )
        }

        # Holds a list of new headers in a dict keyed by sheet id
        new_headers = {}

        # Holds appendCells requests
        append_row_requests = []

        for row in batch:
            headers = headers_by_sheet_id[row.sheet_id]
            ordered_values = [""] * len(headers)

            for field, value in _flatten_values(row.data.items()):
                try:
                    headers.index
                    field_idx = headers.index(field)
                except ValueError:
                    # New Field
                    field_idx = len(headers)
                    headers.append(field)
                    ordered_values.append(value)
                    new_headers.setdefault(row.sheet_id, []).append(field)
                else:
                    # Existing field
                    ordered_values[field_idx] = value

            append_row_requests.append(
                _append_row_request(row.sheet_id, ordered_values, sheet_tz)
            )

        # Build requests to add colunns and set headers
        header_requests = []
        for sheet_id, new_headers in new_headers.items():
            header_requests.append(_add_columns_request(sheet_id, len(new_headers)))
            header_requests.append(
                _update_headers_request(sheet_id, headers_by_sheet_id[sheet_id], sheet_tz)
            )

        update_requests = []

        # First submit header related requests
        update_requests.extend(header_requests)

        # Then we can append rows now that the headers are sorted out
        update_requests.extend(append_row_requests)

        if update_requests:
            ss.batch_update(
                {
                    "requests": update_requests,
                    "includeSpreadsheetInResponse": False,
                    "responseIncludeGridData": False,
                }
            )
        
        return self