# [Docassemble Google Sheets Integration](https://docassemble.org/) from [DocuMoose](https://documoose.ca/)

`google_sheets.py` is a python module for Docassemble to send data from an interview to google sheets.

## Benefits
This module allows you to:
- batch many row entries per sheet tab for quicker speed and more efficient use of api calls.
- post to different tabs by id
- write to correct cell by named headers even if column is moved since last data entry (not during api call)
- write object data with prepended names based on object name
- have values written into sheet correctly based on datatypes
- set datetimes


## Installation
1. Copy the `google_sheets.py` file into your DA package folder.
2. Import the module into your interview file

## Usage
`GoogleSheetAppender` has only one method to use, `append_batch`.
It takes a `spreadsheet_key` which is the GS sheet ID, and `batch`, which is an array of AppendRow instances.

`AppendRow` has a `sheet_id` and `data` key.
'sheet_id' is integer value for the Google Sheet tab id
'data' is a array of dictionaries of the actual column and row data.
Each item in the data array represents a row entry and the keys of the dictionary are the column headers.

### Example
`example_interview.yml`
```yml
imports:
  - datetime
---
modules:
  - .google_sheets
---
objects:
  - events: DAList.using(object_type=Thing, ask_number=True)
---
mandatory: True
code: |
  likes_animals
  favorite_vegetable
  favorite_color
  events.gather()
  sheet_data_sent
  final_screen
---
question: Do you like animals?
fields:
  - Animal's are great!: likes_animals
    datatype: yesnoradio
  - Animal: favorite_animal
    show if: likes_animals
    required: True
---
question: What is your favorite vegetable?
fields:
  - Vegetable: favorite_vegetable
---
show if: favorite_animal
question: What a coincidence!
subquestion: |
  My favorite animal is the ${ favorite_animal }, too!
buttons:
  - Submit: continue
---
question: What is your favorite color
fields:
  - Color: favorite_color
---
question: How many events?
fields:
  - number: events.target_number
    min: 0
    max: 2
    datatype: integer
---
question: Event ${ ordinal(i) } details.
fields:
  - Name: events[i].name.text
  - Date: events[i].date
    datatype: date
  - Host: events[i].host
  - duration: events[i].duration
    datatype: integer
---
id: events_thanks
mandatory: |
  events.target_number > 0
question: |
  Thank you for telling us about these events.
subquestion: |
  % for event in events:
  ${ event }: Held on ${ event.date } with ${ event.host } as the host.

  % endfor
buttons:
  - Continue: continue
---
include:
  - append-data.yml
---
id: final_screen
event: final_screen
prevent going back: True
question: Thank you for your submission!
buttons:
  - Restart: restart
  - Exit: exit
---
```

`append-data.yml`
```yml
code: |
  SHEET_ID = "1234abcdef567890-ab12"
  SHEET_TABS = {
      'example_tab_1': 0,
      'example_tab_2': 123450033
  }

  example_sheet_data = []

  example_data = {
      'time_submitted': datetime.datetime.now(),
      'favorite_color': favorite_color,
      'likes_animals': likes_animals,
  }
  if (likes_animals):
    example_data['favorite_animal'] = favorite_animal

  example_sheet_data.append(AppendRow(sheet_id=SHEET_TABS['example_tab_1'], data=example_data))

  another_example_data = {
      'time_submitted': datetime.datetime.now(),
      'favorite_color': favorite_color
  }

  example_sheet_data.append(AppendRow(sheet_id=SHEET_TABS['example_tab_1'], data=another_example_data))

  if len(events.elements) > 0:
    for event in events:
      event_data = {
        'event': event.name.text,
        'host': event.host,
        'date': event.date,
        'duration': event.duration
      }
      example_sheet_data.append(AppendRow(sheet_id=SHEET_TABS['example_tab_2'], data=event_data))
  
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
