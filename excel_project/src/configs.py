"""
File: configs.py
Author: Kyle Fullerton
Purpose: File that includes functions pertaining to configuration parameters.
"""
import configparser
import os
import sys

# Dictionaries for the various types of configuration parameters

CONFIG_STRS = {"out_file": "",
               "out_sheet_name": "",
               "in_serial_num_column": "",
               "in_sheet_name": "",
               "total_sheet_name": "",
               "gbook_id": "",
               "sheet1_title": "",
               "book_title": ""}

CONFIG_INTS = {"serial_num_column": "",
               "header_row": "",
               "qty_start": "",
               "out_remarks_index": "",
               "column_width": "",
               "doc_labels_column": "",
               "doc_values_column": "",
               "label_start_row": "",
               "label_end_row": "",
               "headers_row": "",
               "part_num_row": "",
               "part_num_column": "",
               "data_start": "",
               "total_header_row": ""}

CONFIG_BOOLS = {"add_mode": "",
                "lines_skipped": "",
                "use_gsheets": ""}

CONFIG_LISTS = {"doc_labels": "",
                "header_list": "",
                "out_default_headers": "",
                "total_sheet_headers": ""}

CONGIG_INT_LISTS = {"wanted_columns": "",
                    "check_columns": ""}

"""
Method: make_config_dict
Purpose: Creates the dictionary of configuration options.
First, creates an initial dictionary mapped to empty strings.
Then uses the default configuration file to fill the dictionary
with valid configuration options. If another config file or parameter
was specified on the command line, and the corresponding configuration 
option(s) are valid, they will overwrite the default configuration
options. Lastly, a check is made to see that all configuration 
options have been specified/are valid.

Variables:
config_file- default configuration file
default_config- file path to the default configuration file
init_dict- initial dictionary of configuration options mapped to
empty strings
config_dict- dictionary of configuration options

Parameters: 
arg_config_file- argument configuration file path or None
parameters- list of specified parameters to change or None

Return:
config_dict- dictionary of configuration options
"""


def make_config_dict(arg_config_file, parameters):
    config_file = "test_config.ini"
    # gets the absolute path of the configuration file no matter the os
    default_config = os.path.abspath(os.path.join
                                     (os.path.dirname(__file__), os.pardir, config_file))

    init_dict = init_config_dict()
    config_dict = add_configs(init_dict, default_config)

    if arg_config_file is not None:
        config_dict = add_configs(config_dict, arg_config_file)

    elif parameters is not None:
        config_dict = add_parameters(config_dict, parameters)

    check_config_dict(config_dict)
    return config_dict


"""
Method: init_config_dict
Purpose: Creates the initial configuration dictionary 
mapped to empty strings by updating an empty dictionary
with the other various type dictionaries. 

Parameter: config_dict- dictionary of configuration options

Return: config_dict- dictionary of configuration options
"""


def init_config_dict():
    config_dict = {}

    config_dict.update(CONFIG_STRS)
    config_dict.update(CONFIG_BOOLS)
    config_dict.update(CONFIG_INTS)
    config_dict.update(CONFIG_LISTS)
    config_dict.update(CONGIG_INT_LISTS)

    return config_dict


"""
Method: add_configs
Purpose: Checks that there is a valid configuration file. If not
an error message is produced.
Then loops through the key value pairs of the configuration
parameters, adding valid values to the main dictionary
of configuration parameters. Invalid values are skipped and
an error message is produced. 

Parameters: 
config_dict- current dictionary of configuration parameters for the program
file_path- configuration file

Variables:
file- filename
config- ConfigParser object 
config_file- first section in the INI file
value- value in the key:value pair in the INI file

Return: config_dict- current dictionary of configuration parameters 
"""


def add_configs(config_dict, file_path):
    file = os.path.basename(file_path)
    config = configparser.ConfigParser()

    try:
        config.read(file_path)

    except configparser.ParsingError:
        print("Error: file {0} doesn't match INI formatting.".format(file))
        return config_dict

    if len(config.sections()) == 0:
        print("Error: couldn't read in config file {0}".format(file))
        return config_dict

    config_file = config.sections()[0]

    for key in config[config_file]:
        if key in config_dict:
            value = config[config_file][key]

            try:
                value = check_entry(key, value)

            except ValueError:
                print("{0} in {1} couldn't be converted to an int".format(key, file))
                continue

            except IndexError:
                print("{0} in {1} can't be a negative value".format(key, file))
                continue

            config_dict[key] = value

    return config_dict


"""
Method: check_entry
Purpose: Checks to see which dictionary the key
is in and provide a necessary type conversion.
 Then provides error checking when converting
strings to integers and string processing when 
converting strings to lists. 

Parameters:
key- string that maps a key to some value
value- string that will be converted to a type 
based on which dictionary its corresponding key is found

Variable: new_value- value converted to its new type

Return: new_value- value converted to its new type 
"""


def check_entry(key, value):
    if key in CONFIG_INTS:
        try:
            new_value = int(value)

            if new_value < 0:
                raise IndexError

        except ValueError:
            raise ValueError

    elif key in CONFIG_BOOLS:
        new_value = (value == "True")

    elif key in CONFIG_LISTS:
        new_value = []

        for item in value.split(","):
            values = item.split("\n")
            new_value.append(" ".join(values).strip())

    elif key in CONGIG_INT_LISTS:
        new_value = []
        for item in value.split(","):
            try:
                int_item = int(item.strip())
                if int_item < 0:
                    raise IndexError

                new_value.append(int_item)

            except ValueError:
                raise ValueError

    else:
        new_value = value

    return new_value


"""
Method: add_parameters
Purpose: Goes through the list of parameters and checks that 
the current parameter is formatted correctly. Then the parameter
is split into key and value. Lastly, the parameter is
checked to see if it is contained in the configuration parameter 
dictionary. If found, the check_entry function is performed to
do a possible type conversion and error checking. 

Parameters: 
config_dict- dictionary of configuration parameters
parameters- list of configuration parameters to overwrite

Variables:
updated_dict- configuration parameter dictionary that has
its entries updated according to the specified valid paramters
key- string that maps a key to some value
value- string that will be converted to a type 

Return: updated_dict- updated dictionary of configuration parameters 
"""


def add_parameters(config_dict, parameters):
    updated_dict = config_dict
    for entry in parameters:
        if "=" not in entry and ":" not in entry:
            print("Invalid parameter pair {0}. Must be separated by \":\" or \"=\"".format(entry))
            continue

        if "=" in entry:
            key = entry.split("=")[0]
            value = entry.split("=")[1]

        else:
            key = entry.split(":")[0]
            value = entry.split(":")[1]

        if key in config_dict:
            try:
                new_value = check_entry(key, value)

            except ValueError:
                print("Value '{0}' for key '{1}' couldn't "
                      "be converted to an int".format(value, key))
                continue

            except IndexError:
                print("Value '{0}' for key '{1}' can't be a negative value".format(value, key))
                continue

            updated_dict[key] = new_value

    return updated_dict


"""
Method: check_config_dict
Purpose: Checks that the configuration parameter dictionary
has all of the configuration parameters specified. Any offending
configuration parameters are outputted to the console and the
program is exited.

Parameter: config_dict- dictionary of configuration parameters
"""


def check_config_dict(config_dict):
    if "" in config_dict.values():
        print("Error: not given all valid configuration parameters.\n\nNot given:")
        for key, value in config_dict.items():
            if value == "":
                print(key)

        sys.exit(1)
