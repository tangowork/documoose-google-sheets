# documoose-google-sheets

## Enhanced Google Sheets integration for [Docassemble](https://docassemble.org), by [DocuMoose](https://documoose.ca/)

`google_sheets.py` is a Python module for Docassemble to send data from an interview to Google Sheets.

We started with the Docassemble [demo code for Google Sheets](https://github.com/jhpyle/docassemble/blob/master/docassemble_demo/docassemble/demo/google_sheets.py) and added the enhancements listed below. 


## Benefits

- Refer to sheet by id instead of name
- Post to different tabs by tab id
- Post row updates in batches for 10X - 50X faster performance
- Write to correct cell even if column is moved (column name lookup)
- Write object data (breaks objects into multiple columns)
- Set cell format based on datatype
- Set datetimes



## Installation

1. Follow the [Docassemble documentation for using Google Sheets](https://docassemble.org/docs/functions.html#google%20sheets%20example):
   1. Set up your Google developer account as described
   2. Select a Google Sheet as described
   3. Add your service account credentials to Configuration as described
2. Go to the modules folder of your Docassemble Playground. Copy the file [`google_sheets.py`](https://github.com/tangowork/docassemble-google_sheets_integration/blob/main/google_sheets.py) into your modules folder
3. Reference the module in your interview file, e.g.:
```
modules:
  - .google_sheets
```

## Usage

`GoogleSheetAppender` has only one method to use: `append_batch`. It expects:
- `spreadsheet_key` (the Google Sheet spreadsheet id) 
- `batch` (an array of `AppendRow` instances)

`AppendRow` has a `sheet_id` and `data` key.
- `sheet_id` is the integer value for the Google Sheet tab id.
- `data` is an array of dictionaries of the actual column and row data.

Each item in the data array represents a row entry, and the keys of the dictionary are the column headers.


### Example

[Try this demo live](https://app.documoose.ca/start/google-sheets-demo)

To try this example:
1. Install `google_sheets.py` as described above
2. Create two interview files using the code below, `example_interview.yml` and `append-data.yml`
3. In `append_data.yml`, update SHEET_ID and SHEET_TABS with the actual ids from your Google Sheet.

#### example_interview.yml
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
  - Animals are great!: likes_animals
    datatype: yesnoradio
  - What's your favourite animal?: favorite_animal
    show if: likes_animals
    required: True
---
question: What's your favorite vegetable?
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
question: What's your favorite colour?
fields:
  - Colour: favorite_color
---
question: How many events did you attend last month?
sub: Max 2
fields:
  - Number of events: events.target_number
    min: 0
    max: 2
    datatype: integer
---
question: Please provide details of the ${ ordinal(i) } event.
fields:
  - Event name: events[i].name.text
  - Date: events[i].date
    datatype: date
  - Host: events[i].host
  - Duration: events[i].duration
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

#### append-data.yml
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

If you want to write to two spreadsheets instead of tabs (or in addition), set up the data the same as the third example but add an additional data array:
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
