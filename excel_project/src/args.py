"""
File: args.py
Author: Kyle Fullerton
Purpose: Handles parsing the command line arguments and puts the given
arguments into their respective variables.
"""

import argparse
# The argparse library does its own error checking. Also the last argument given is what will
# be used for the arguments respective variable. Ex: -c file_path -c other_file will result in
# other_file used as the configuration file for the program.

"""
Method: parse_arguments
Purpose: Parses the command line arguments, 
puts those arguments into variables, and then
returns the variables to the caller. 

Returns: 
directory- filepath to a directory of files
config_file- filepath to an alternate configuration file
pars- A list of configuration parameters the user wants to change.
Must be like INI format with no spaces in between "foo=2" or "foo:2".
"""


def parse_arguments():
    parser = argparse.ArgumentParser()
    add_options(parser)

    args = parser.parse_args()
    directory = args.directory
    config_file = args.config_file
    parameters = args.parameters

    return directory, config_file, parameters


"""
Method: add_options
Purpose: Adds command line options to store specific
arguments into their respective variables.

Parameter: parser- ArgumentParser object
"""


def add_options(parser):
    # dest is where the argument is stored
    # required means the argument is required or an error occurs
    # nargs is the number of arguments read in after the commands
    # help is the help message that appears after -h

    parser.add_argument("-d", "--directory", dest="directory", required=True,
                        help="file path to a directory to read in excel files")

    parser.add_argument("-c", "--config-file", dest="config_file",
                        help="specify a different configuration file besides the default")

    parser.add_argument("-p", "--change-parameters", dest="parameters", nargs="*",
                        help="change one or more of the default configuration parameters")

