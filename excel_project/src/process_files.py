"""
File: process_files.py
Author: Kyle Fullerton
Purpose: File that contains functions that process the .xslx files.
Whether that be for reading a spreadsheet or writing to one.
"""

import os
import sys
import openpyxl
import xlrd

from openpyxl.utils.exceptions import InvalidFileException
from xlrd.biffh import XLRDError

"""
Method: find_write_file
Purpose: Finds the specified write file in the directory and
pops it from the file list. This happens if
the filename or path with the filename is contained in the 
directory.

Parameters: 
directory_path- file path to the directory of files
out_file- excel file that will be used for writing data to

Variables:
files- list of files in the directory
file- filename to find in the directory
idx- index in the list of files where the file was found
write_file- file path of the output file

Returns: files- list of remaining files in the directory minus
the output file
write_file- excel file that will be used for writing data to
"""


def find_write_file(directory_path, out_file):
    files = os.listdir(directory_path)
    file = os.path.basename(out_file)

    try:
        idx = files.index(file)
        write_file = os.path.abspath(os.path.join(directory_path, files.pop(idx)))

    except ValueError:
        write_file = out_file

    return files, write_file


"""
Method: get_valid_writebook
Purpose: Uses the output file path and tries to
get valid Openpyxl workbook and worksheet objects.
Various error messages are printed should an error
occur with getting the workbook and worksheet objects.

Parameters: 
file_path- file that will be written to
sheet_title- specified sheet title in the workbook

Variables:
out_write_book- work book from the write excel file
out_write_sheet- worksheet from the write excel file

Returns- out_write_book- work book from the write excel file
out_write_sheet- worksheet from the write excel file
"""


def get_valid_writebook(file_path, sheet_title):
    file = os.path.basename(file_path)

    # Returns a valid workbook
    # or prints out an error message and exits the program

    try:
        out_write_book = openpyxl.load_workbook(file_path)

    except InvalidFileException:
        print("Error: file from {0} is not an .xslx file".format(file))
        sys.exit(1)

    except FileNotFoundError:
        print("Error: file {0} cannot be found from path {1}".format(file, file_path))
        sys.exit(1)

    # Returns a valid worksheet
    # or prints out an error message and exits the program

    try:
        out_write_sheet = out_write_book[sheet_title]

    except KeyError:
        print("Error: sheet {0} doesn't exist in the output file".format(sheet_title))
        sys.exit(1)

    return out_write_book, out_write_sheet


"""
Method: get_valid_readbook
Purpose: Tries to create a valid XLRD workbook object from the passed
in filepath. Then tries to create a valid XLRD worksheet object.
If any errors occur than an exception is raised and passed up 
to the caller.  

Parameters: 
file_path- path to the file that will be read in
configs- dictionary of configuration parameters
file- filename used to help provide a descriptive error message

Returns: 
read_book- work book object from the read excel file
read_sheet- worksheet object from the read excel file
"""


def get_valid_readbook(file_path, configs):
    try:
        read_book = xlrd.open_workbook(file_path)

    except XLRDError:
        raise XLRDError

    try:
        read_sheet = check_read_sheet(read_book, configs)

    except RuntimeError as error:
        raise RuntimeError(error)

    except IndexError as parameter:
        raise IndexError(parameter)

    return read_book, read_sheet


"""
Method: check_read_sheet
Purpose: Checks that the specified sheet name is 
contained in the workbook object. Then performs 
various checks to make sure that the worksheet
object is formatted correctly. If an error occurs
in the formatting then an exception is raised and
passed up to the caller. 

Parameters:
work_book- work book object from the read excel file
configs- dictionary of configuration parameters

Returns:
sheet- worksheet object from the read excel file
"""


def check_read_sheet(work_book, configs):
    sheet_list = work_book.sheets()
    # Check first sheet name
    if sheet_list[0].name != configs["in_sheet_name"]:
        raise RuntimeError("Error: {0} doesn't have the specified first sheet name")

    sheet = work_book.sheet_by_name(configs["in_sheet_name"])

    # Check that document labels are there
    try:
        doc_labels_col = sheet.col_values(configs["doc_labels_column"] - 1)
        doc_labels = doc_labels_col[configs["label_start_row"] - 1: configs["label_end_row"]]

    except IndexError:
        raise IndexError("'doc_labels_column'")

    if doc_labels != configs["doc_labels"]:
        raise RuntimeError("Error: {0} doesn't have the right specified doc labels")

    try:
        doc_info_column = sheet.col_values(configs["doc_values_column"] - 1)
        doc_info = doc_info_column[configs["label_start_row"] - 1:configs["label_end_row"]]

    except IndexError:
        raise IndexError("'doc_values_column'")

    # Check that there are values corresponding to the document labels
    if "" in doc_info:
        raise RuntimeError("Error: {0} is missing one or more document values")

    try:
        header_list = sheet.row_values(configs["headers_row"] - 1)

    except IndexError:
        raise IndexError("'headers_row'")

    # Check that the headers are correct
    if len(header_list) != len(configs["header_list"]):
        raise RuntimeError("Error: {0} header lists are different sizes")

    for i in range(0, len(header_list)):
        header1 = str(header_list[i]).split("\n")
        header1 = " ".join(header1)
        header1 = header1.split()

        header2 = configs["header_list"][i]
        header2 = header2.split()

        if header1 != header2:
            raise RuntimeError("Error: {0} header list of sheet is different than specified")

    return sheet


"""
Method: create_part_dict
Purpose: Creates a dictionary of part number mapped to the 
rest of the specified row data from the read excel sheet. 

Parameters: 
out_read_sheet- the XLRD worksheet from the output excel file
configs- dictionary of configuration parameters

Returns: 
parts_dict- dictionary of part number mapped to the 
rest of the specified row data from the read excel sheet. 
"""


def create_part_dict(out_read_sheet, configs):
    parts_dict = {}

    # if sheet is empty return an empty dict
    if out_read_sheet.ncols == 0:
        return parts_dict

    else:
        try:
            part_num = out_read_sheet.col_values(configs["serial_num_column"] - 1)

        except IndexError:
            print("Error: config parameter {0} defines an out of range column"
                  .format("'serial_num_column'"))
            sys.exit(1)

        # Loop through all values except for the headers
        for row_num in range(configs["header_row"], len(part_num)):
            row = out_read_sheet.row_values(row_num)

            # index 0 is the row number
            row[0] = row_num + 1

            if configs["qty_start"] >= len(row):
                print("Error: {0} specifies an out of range column".format("'qty_start'"))
                sys.exit(1)
            # index 2 to n is all of the qtys
            for i in range(configs["qty_start"] - 1, len(row)):
                row[i] = [row[i], i]

            parts_dict[str(part_num[row_num]).strip()] = row
    return parts_dict
