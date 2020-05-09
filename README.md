# Convert To YNAB
This is a Python project I created to convert CSV files (and now also an Excel file) from various banks to a format that loads into my budget software (called YNAB or "You Need A Budget").

**NOTE: this code represents a specific point in time in my ongoing learning of Python. Certain code usage or patterns don't necessarily represent how I might code today.**

# Using the Application
I created an executable with PyInstaller using default settings. This creates a dist directory which needs to be placed on the harddrive and then call the executable. The executable has command line parameters as follows:
  * requires a path for source files (excluding a final "\\")
  * requires a path for output files  (excluding a final "\\")
  * NOTE: source files are archived to a "\loaded" subdirectory under the source file path
  
Both of the above paths should exist already. The "loaded" subdirectory does not have to exist. Source files cannot already exist in the loaded directory. The pathing could potentially be either relative or fully qualified based on how the executable is run and/or located. Fully qualified would be safest.

The source files themselves are pretty specific to the banks I use. The code would need to be modified for other bank source files. In most cases, these are CSV format. One particular bank export is in Excel format.

# Major Functions
The application performs the following major functions.

## Find Valid Source Files
* Retrieves a list of files from the source path
* Parses the file list for any files that have a predefined prefix + "_" 

## Main Processing
There are two parts to processing a file: an overall process flow and "transforming" a source row to a target row.

### Overall Processing
* opens the file
    * for each row in file do some initial data quality checks to skip garbage rows else sends a row for mapping
* when all rows are processed, generates an output file name (the source file name + "_output" + currentdatetime
* writes processed rows to output file in output directory
* move source file to archive directory

### Map Source Row to Target Row
* assign source columns to variables and do some basic cleaning
* build a target row from above variables with additional cleaning or formatting
* return target row to process function
