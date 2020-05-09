# Convert To YNAB
This is a Python project I created to convert CSV files (and now also an Excel file) from various banks to a format that loads into my budget software (called YNAB or "You Need A Budget"). It has the following major functions:

## Using the Application
The dist/ directory contains a runtime executable of the Python script based on compiling with PyInstaller using default settings. The entire directory needs to be placed on the harddrive and then call the executable. The executable has command line parameters as follows:
  * requires a path to source files
  * requires a path to output files
  * NOTE: source files are moved to a "\loaded" subdirectory under the source file path

## Find Valid Source Files
* Retrieves a list of files from the source path
* Parses the file list for any files that have a predefined prefix + "_" 

## Main Processing
There are two parts to processing a file: an overall process flow and "transforming" a source row to a target row.
### Overall Processing
* opens the file
    * for each row in file do some initial data quality checks to skip garbage rows else sends a row for mapping
* when all rows are processed, generates an output file name (the source file name + "_output" + currentdatetime
* writes processed rows to output file
* move source file to archive directory

### Map Source Row to Target Row
* assign source columns to variables and do some basic cleaning
* build a target row from above variables with additional cleaning or formatting
* return target row to process function
