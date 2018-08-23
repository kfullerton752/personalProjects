"""
File: main.py
Author: Kyle Fullerton
Purpose: File that is used to control the main flow of the
program.
"""

import xlrd
import os

from excelScript import (args,
                         process_files,
                         utils,
                         configs,
                         excel,
                         gsheets
                         )

from xlrd.biffh import XLRDError
import sys
# import time


"""
Method: main
Purpose: Used to control the main flow of the program. 
Makes function calls to parse the arguments, set up
configuration parameters, read in an output file,
and finally process the other read in files in the directory

Variables: 
directory_path- Path to the specified directory on the command line
arg_config_file- None or a filename/path for a specified configuration file
parameters- List of specified parameters to change
configs_dict- Dictionary of configuration parameters
files- List of files to read in the directory
write_file- Output excel file that will be written to
out_write_book- Openpyxl workbook object of the output excel file
out_write_sheet- Openpyxl worksheet of the output excel file
out_read_book- XLRD workbook object of the output excel file
out_read_sheet- XLRD workbook object of the output excel file
part_dict- Dictionary mapping part numbers to the rest of the wanted data
header_list- List of headers that are on the output excel file 
file_path- Full file path for the other read in files in the directory 
in_read_book- XLRD workbook object of an input excel file
in_read_sheet- XLRD worksheet object of an input excel file
"""


def main():
    # start_time = time.clock()

    directory_path, arg_config_file, parameters = args.parse_arguments()
    configs_dict = configs.make_config_dict(arg_config_file, parameters)

    files, write_file = process_files.find_write_file(directory_path, configs_dict["out_file"])
    out_write_book, out_write_sheet = process_files.\
        get_valid_writebook(write_file, configs_dict["out_sheet_name"])

    out_read_book = xlrd.open_workbook(write_file)
    out_read_sheet = out_read_book.sheet_by_name(configs_dict["out_sheet_name"])

    part_dict = process_files.create_part_dict(out_read_sheet, configs_dict)
    header_list = excel.add_headers(out_read_sheet, out_write_sheet, configs_dict)

    for file in files:
        file_path = os.path.abspath(os.path.join(directory_path, file))

        try:
            in_read_book, in_read_sheet = process_files.get_valid_readbook(file_path, configs_dict)
            utils.add_assembly_num(header_list, out_write_sheet, in_read_sheet, configs_dict)
            part_dict = excel.update_master(in_read_sheet, out_write_sheet, header_list,
                                            part_dict, configs_dict, file, out_write_book)
        except XLRDError:
            print("Error: {0} was not read in since it's not a .xlsx file".format(file))
            continue

        except RuntimeError as error:
            print(str(error).format(file))
            continue

        except IndexError as parameter:
            print("Error: config parameter {0} defines an out of "
                  "range column for file {1}".format(parameter, file))
            continue

    utils.edit_column_width(out_write_sheet, header_list, configs_dict)
    title = configs_dict["total_sheet_name"]
    utils.edit_column_width(out_write_book[title],
                            configs_dict["total_sheet_headers"], configs_dict)
    out_write_book.save(write_file)

    if configs_dict["use_gsheets"]:
        gsheets.execute(part_dict, header_list, configs_dict)

    # print(time.clock() - start_time, "seconds")


main()
