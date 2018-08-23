Command Line:

-d or -directory followed by the directory path is required for the program
to work.
-c with a configuration filename or path and -p followed by a list of parameters in
INI key=value format are optional arguments that can be put on the command line.
Also, -h can be used to find help with the usage of the command line arguments.
These additional options are used to overwrite the default configuration parameters.
See the default configuration file "test_config.ini" for documentation on these parameters.

Functionality:

Takes in a directory of files to read in as a command line argument. Every file read in
the directory is checked that it has the correct specified formatting. Invalid files
will be outputted to the console. Then the specified columns from the read in spreadsheets are
added to the output file.

    Default:
    For default functionality you will have to provide the path or the filename of the excel
    spreadsheet. The filename option can only be used if the excel spreadsheet is located in
    the same directory given to the program via command line arguments.

    Google Sheets Mode:
    If using google sheets mode you will have to create a new spreadsheet or supply
    the spreadsheet id for the existing spreadsheet. The spreadsheet id is located in
    the URL after the d/ and before the /edit portion of the URL. An example URL is supplied below.
    https://docs.google.com/spreadsheets/d/1VeRWOUJNw-vvkChCo4u2ZV9fM-eMQb8ttVrUGImCEY8/edit#gid=0
