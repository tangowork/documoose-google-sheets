# [Docassemble Google Sheets Integration](https://docassemble.org/) from [DocuMoose](https://documoose.ca/)

google_sheets.py is a python module for Docassemble to send data from an interview to google sheets.

## Benefits
This module allows you to:
- batch many row entries per sheet tab for quicker speed and more efficient use of api calls.
- post to different tabs by id
- write to correct cell by named headers even if column is moved since last data entry (not during api call)
- write object data with prepended names based on object name
- have values written into sheet correctly based on datatypes
- set datetimes


## Installation
1. Copy the google_sheets.py file into your DA package folder.
2. Import the module into your interview file

## Usage
GoogleSheetAppender has only one method to use, append_batch.
It takes a spreadsheet_key which is the GS sheet ID, and batch, which is an array of AppendRow instances.

AppendRow has a sheet_id and data key.
'sheet_id' is integer value for the Google Sheet tab id
'data' is a array of dictionaries of the actual column and row data.
Each item in the data array represents a row entry and the keys of the dictionary are the column headers.

### Example
```yml
objects:
    - user: Person
    - second_user: Person
---
code: |
    SHEET_ID = "1234abcd5678efgh9090"


    example_data = {
        'time_submitted': datetime.datetime.now(),
        'submitter': user,
        'favourite_colour': favourite_color,
        'likes_animals': likes_animals,
    }

    another_example_data = {
        'time_submitted': datetime.datetime.now(),
        'submitter': user,
        'favourite_animal': favourite_animal,
        'likes_animals': likes_animals,
    }

    third_example_data = {
        'event_date': event.datetime,
        'event_name': event.name,
        'event_host': event.host,
        'event_durattion': event.duration
    }

    SHEET_TABS = {
        'example_tab_1': 0,
        'example_tab_2': 1234567890
    }

    example_sheet_data = []
    example_sheet_data.append(AppendRow(sheet_id=SHEET_TABS['example_tab_1']), data=example_data))
    example_sheet_data.append(AppendRow(sheet_id=SHEET_TABS['example_tab_1']), data=another_example_data))
    example_sheet_data.append(AppendRow(sheet_id=SHEET_TABS['example_tab_2']), data=third_example_data))
    
    GoogleSheetAppender().append_batch(
        spreadsheet_key=SHEET_ID,
        batch=example_sheet_data
    )

    sheet_data_sent = True
```

If you wanted to write to two spreadsheets instead of tabs (or in addition), set up the data the same as the third example but add an additional data array

```yml
    second_spreadsheet_data = []
    example_sheet_data.append(AppendRow(sheet_id=SHEET_TABS['example_tab_1'], data=example_data))
    example_sheet_data.append(AppendRow(sheet_id=SHEET_TABS['example_tab_1'], data=another_example_data))
    second_spreadsheet_data.append(AppendRow(sheet_id=0, data=third_example_data))

    GoogleSheetAppender().append_batch(
        spreadsheet_key=SHEET_ID,
        batch=example_sheet_data
    ).append_batch(
        spreadsheet_key=SHEET_ID,
        batch=second_spreadsheet_data
    )

    sheet_data_sent = True
```