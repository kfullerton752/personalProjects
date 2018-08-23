"""
File: gsheets.py
Author: Kyle Fullerton
Purpose: File that is used to handle functionality with google sheets.
For additional documentation lookup Google's Sheets API. A good starting
place would be the batchUpdate page which has information and links
to many of the update requests made with this file.
"""
from excelScript import utils
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import httplib2
import sys


"""
Method: execute
Purpose: Makes calls to authorize the request credentials, update
properties with the spreadsheets, and then add requests to update
the spreadsheet headers and data. 

Parameters: 
part_dict- Dictionary mapping part numbers to the rest of the wanted data
header_list- List of headers that are on the output excel file
config_dict- Dictionary of configuration parameters

Variables:
book_id- file id for the workbook
requests- will be a list of dictionaries (Json objects) used to make 
requests to edit the spreadsheet
sheet1_id- unique id to the first sheet in the workbook
total_id- unique id to the totals sheet in the workbook
"""


def execute(part_dict, header_list, config_dict):
    book_id = config_dict["gbook_id"]
    requests = []

    service = authorize()
    sheet1_id, total_id = update_sheets(service, book_id, config_dict)

    requests.append(update_headers(header_list, sheet1_id, config_dict["header_row"]))
    requests.append(update_headers(config_dict["total_sheet_headers"],
                                   total_id, config_dict["total_header_row"]))

    requests.extend(update_values(part_dict, header_list, config_dict, sheet1_id, total_id))

    requests.append(resize_columns(header_list, sheet1_id))
    requests.append(resize_columns(config_dict["total_sheet_headers"], total_id))

    request_body = {"requests": requests}

    try:
        service.spreadsheets().batchUpdate(spreadsheetId=book_id, body=request_body).execute()
    except HttpError as error:
        print(error)
        sys.exit(1)


"""
Method: authorize
Purpose: Used to authorize the credentials in order to use
the google sheets API. 

Variables:
scope- amount of access to the API needed for the program
to function
credentials- ServiceAccountCredentials object 
http- HTTP object
service- googleapiclient.discovery.Resource object

Return:
service- googleapiclient.discovery.Resource object
"""


def authorize():
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name("service_key.json", scope)

    http = httplib2.Http()
    http = credentials.authorize(http)

    service = build("sheets", "v4", http=http)

    return service


"""
Method: update_sheets
Purpose: Used to edit the title of the work book. Also, 
adds additonal work sheets if the specified work sheets
are not found in the workbook.

Parameters:
service- googleapiclient.discovery.Resource object
book_id- file id for the workbook

Variables:
total_id- unique id to the totals sheet in the workbook
sheet1_id- unique id to the first sheet in the workbook
result- json response after the request
body- request body for sending a request to the API
response- json response after the request

Returns:
total_id- unique id to the totals sheet in the workbook
sheet1_id- unique id to the first sheet in the workbook
"""


def update_sheets(service, book_id, config_dict):
    total_id, sheet1_id = "", ""

    # get information about worksheet
    try:
        result = service.spreadsheets().get(spreadsheetId=book_id).execute()

    except HttpError as error:
        print(error)
        sys.exit(1)

    workbook_title = result["properties"]["title"]

    # change workbook title if necessary
    if workbook_title != config_dict["book_title"]:
        body = {"requests": [{"updateSpreadsheetProperties":
               {"properties": {"title": config_dict["book_title"]},
                "fields": "*"}}]}
        try:
            service.spreadsheets().batchUpdate(spreadsheetId=book_id, body=body).execute()
        except HttpError as err:
            print(err)
            sys.exit(1)

    # loops through the sheets in the json response
    # looks for the specified sheet names
    for sheet in result["sheets"]:
        if sheet["properties"]["title"] == config_dict["total_sheet_name"]:
            total_id = sheet["properties"]["sheetId"]

        if sheet["properties"]["title"] == config_dict["sheet1_title"]:
            sheet1_id = sheet["properties"]["sheetId"]

    # create the totals sheet since it wasn't found
    if total_id == "":
        body = {"requests": [{"addSheet": {"properties": {"title": "Card Totals"}}}]}

        try:
            response = service.spreadsheets().batchUpdate(
                       spreadsheetId=book_id, body=body).execute()
            total_id = response["replies"][0]["addSheet"]["properties"]["sheetId"]
        except HttpError as error:
            print(error)
            sys.exit(1)

    # create sheet1 since it wasn't found
    if sheet1_id == "":
        body = {"requests": [{"addSheet": {"properties": {"title": "Sheet1"}}}]}

        try:
            response = service.spreadsheets().batchUpdate(
                       spreadsheetId=book_id, body=body).execute()
            sheet1_id = response["replies"][0]["addSheet"]["properties"]["sheetId"]
        except HttpError as error:
            print(error)
            sys.exit(1)

    return sheet1_id, total_id


"""
Method: update_headers
Purpose: Makes the request to update the header formatting 
and values of the spreadsheet. 

Parameters:
headers- list of header strings
sheet_id- unique id for the particular sheet
in the workbook
start_row- row where the headers start

Variables:
values- list of dictionaries that dictate the formatting
and value for the cells
black- dictionary that specifies rgb values for
the color black
gray- same as black only for gray this time
border- dictionary that dictates the border color
and border type
borders- dictionary that dictates which borders of
the cell will be formatted
cell_range- range of cells where the headers will
be located
cell_format- dictionary that dictates the formatting
for the cells
cell_value- value that will be inserted into the cell
header- Dictionary that dictates what attributes of 
the cells will be updated. "*" for fields means that 
any change will be applied to the particular cell.

Return: Returns an updateCells request via a dictionary
of "updateCells" mapped to the header dictionary. 
"""


def update_headers(headers, sheet_id, start_row):
    values = []
    black = {"red": 0,
             "blue": 0,
             "green": 0
             }

    gray = {"red": .72,
            "blue": .72,
            "green": .72
            }

    border = {"style": "SOLID", "color": black}
    borders = {"top": border,
               "bottom": border,
               "left": border,
               "right": border
               }
    cell_range = {"sheetId": sheet_id,
                  "startRowIndex": start_row - 1,
                  "endRowIndex": start_row,
                  "startColumnIndex": 0,
                  "endColumnIndex": len(headers)
                  }

    cell_format = {"borders": borders,
                   "backgroundColor": gray,
                   "horizontalAlignment": "CENTER"
                   }

    for header in headers:
        cell_value = {"stringValue": header}
        values.append({"userEnteredFormat": cell_format,
                       "userEnteredValue": cell_value})

    header = {"range": cell_range,
              "rows": [{"values": values}],
              "fields": "*"}

    return {"updateCells": header}


"""
Method: resize_columns
Purpose: To make a request to auto resize columns in the
spreadsheet. 

Parameters:
header_list- list of header strings
sheet_id- unique id for a sheet in the workbook

Variables:
sheet_dimensions- dictionary that dictates which
columns in the spreadsheet should be auto resized

Returns: 
A dictionary that makes an autoResizeDimensions request.
"""


def resize_columns(header_list, sheet_id):
    sheet_dimensions = {"sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": 0,
                        "endIndex": len(header_list)
                        }

    return {"autoResizeDimensions": {"dimensions": sheet_dimensions}}


"""
Method: update_header 
Purpose: Makes updateCells requests per row of 
data that needs to be entered into both the first sheet
and the totals sheet. 

Parameters:
part_dict- dictionary mapping part numbers to the rest of the wanted data
header_list- list of header strings
config_dict- dictionary of configuration parameters 
sheet1_id- unique id for the first sheet in the workbook
totals_id- unique id for the totals sheet in the workbook

Variables:
qty_start- starting index for the qty of parts
qty_end- ending index for the qty of parts
header_row- row where the headers are located
row_updates- list of dictionaries that specify
the updates to make per row of cells
row_num- row number that corresponds to the part info
found in the excel sheet
data- list of updateCells requests
sheet_range- dictionary that specifies the range of 
cells for the row that will be altered
row- dictionary that specifies the attributes of the
row that will change
total_data- same as data
total_range- same as range only particular to the total sheet
totals_row_num- same as row_num only particular
to the total sheet
range1- cell range 1 for the formula
range1- cell range 2 for the formula
sheet- sheet name for the formula
formula- sumproduct formula used to calculate the number of
parts needed for a specified number of cards
total_row- same as row only particular to the total sheet

Returns: 
A list of cellUpdates requests.
"""


def update_values(part_dict, header_list, config_dict, sheet1_id, totals_id):
    qty_start = config_dict["qty_start"] - 1
    qty_end = len(header_list) - 1
    header_row = config_dict["header_row"]

    row_updates = []
    for keys, values in part_dict.items():
        row_num = values[0] - 1
        data = []
        data.append({"userEnteredValue": {"stringValue": keys}})

        sheet_range = {"sheetId": sheet1_id,
                       "startRowIndex": row_num,
                       "endRowIndex": row_num + 1,
                       "startColumnIndex": 0,
                       "endColumnIndex": len(header_list)
                       }

        for i in range(1, len(values)):
            if isinstance(values[i], list):
                if isinstance(values[i][0], str):
                    cell_value = {"stringValue": values[i][0]}

                else:
                    cell_value = {"numberValue": values[i][0]}

            else:
                cell_value = {"stringValue": values[i]}
            data.append({"userEnteredValue": cell_value})

        row = {"range": sheet_range,
               "rows": [{"values": data}],
               "fields": "*"}

        row_updates.append({"updateCells": row})

        total_data = []
        total_row_num = row_num - config_dict["header_row"] + config_dict["total_header_row"]
        totals_range = {"sheetId": totals_id,
                        "startRowIndex": total_row_num,
                        "endRowIndex": total_row_num + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": len(config_dict["total_sheet_headers"])
                        }

        range1 = utils.get_range(qty_start, qty_end, row_num + 1, row_num + 1)
        range2 = utils.get_range(qty_start, qty_end, header_row - 1, header_row - 1)
        sheet = config_dict["out_sheet_name"] + "!"
        formula = "=SUMPRODUCT({0}{1}, {2}{3})".format(sheet, range1, sheet, range2)

        total_data.append({"userEnteredValue": {"stringValue": keys}})
        total_data.append({"userEnteredValue": {"formulaValue": formula}})

        total_row = {"range": totals_range,
                     "rows": [{"values": total_data}],
                     "fields": "*"}

        row_updates.append({"updateCells": total_row})

    return row_updates
