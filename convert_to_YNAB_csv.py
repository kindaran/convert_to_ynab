import os
import sys
import csv
import logging
import glob
import shutil
import xlrd
from decimal import Decimal
from datetime import datetime


#####
# Functions
#####

def getArgs():

    args = []

    try:
        logging.info("***RETRIEVING COMMAND LINE ARGS")
        if len(sys.argv) == 3:
            sourceFilePath = sys.argv[1]
            args.append(sourceFilePath)
            targetFilePath = sys.argv[2]
            args.append(targetFilePath)
            logging.debug("Command line arguments used: input_path: %s  output_path: %s " % (
                sourceFilePath, targetFilePath))
            return args
        else:
            logging.error("Not enough arguments provided.")
            print(
                "Incorrect arguments provided\r\nPlease include path to source files and path for output")
            return None
    except Exception as e:
        msg = str(e)
        logging.error("*****Error in getArgs. Args: %s  Error: %s" %
                      (sys.argv, msg))
        return None
# END DEF


def getFileList(p_path):
    try:
        logging.info("***GETTING INPUT DIR FILE LIST")
        filelist = [file for file in glob.glob(p_path + r"/*") if
                    not os.path.isdir(file) and "_OUTPUT" not in file.upper()]
        logging.debug("Found %s files in directory %s" % (len(filelist), p_path))
        if len(filelist) > 0:
            return filelist
        else:
            logging.warning(
                "***WARNING: no files found in path %s matching search criteria" % (p_path))
            return None
        # end if
    except Exception as e:
        msg = str(e)
        logging.error("*****Error in getFileList. p_path: %s  Error: %s" % (p_path, msg))
        return None


# end getFileList

def writeFile(p_filename, p_rows):
    global g_OutputPath
    global g_CSVHeader

    try:
        logging.info("***WRITING TO OUTPUT FILE")
        if len(p_rows) > 0:
            output_file = g_OutputPath + "/" + p_filename
            logging.debug('Writing to file: %s' % (output_file))
            with open(output_file, 'w') as hFile:
                hFile.write(g_CSVHeader)
                hFile.write('\n'.join(p_rows))
            logging.info('Done writing')
        else:
            logging.warning("WARNING: No data to write***")
        # end if
    except Exception as e:
        msg = str(e)
        logging.error("*****Error in writeFile. p_filename: %s  p_rows: %s  Error: %s" %
                      (p_filename, p_rows, msg))


# end writeFile

def generateOutputFilename(p_filename):
    try:
        # strips the raw filename out of file string
        filename = p_filename.split(".")[0].split("/")[-1]
        current_datetime = datetime.strftime(datetime.now(), "%Y%m%d%H%M%S")
        output_filename = filename + "_output_" + current_datetime + ".csv"
        logging.debug("Output filename: %s" % (output_filename))
        return output_filename
    except Exception as e:
        msg = str(e)
        logging.error("*****Error in generateOutputFilename. p_filename: %s  Error: %s" %
                      (p_filename, msg))
        return None


# end generateOutputFilename

def moveFile(p_source, p_dest):

    try:
        logging.info("***MOVING FILE FROM %s TO %s" % (p_source, p_dest))
        if os.path.isdir(p_dest) is False:
            logging.debug("%s not found - creating it" % (p_dest))
            os.mkdir(p_dest)
        # end if
        shutil.move(p_source, p_dest)
        logging.debug("File moved.")
    except Exception as e:
        msg = str(e)
        logging.error("*****Error in moveFile. p_source: %s  p_dest: %s  Error: %s" %
                      (p_source, p_dest, msg))
        return None

# end moveFile


def parseTDRow(p_row):
    global g_Months
    global g_Delim

    try:
        # process data row
        trade_date = p_row[0].strip()
        description = p_row[2].strip()
        action = p_row[3].strip()
        net_amount = p_row[7].strip()

        # build output row
        temp = trade_date.split()
        trans_date = str(g_Months.index(temp[1].upper()) + 1).zfill(2)
        # get index of trans month from months list, then left pad with zero
        trans_date += "/" + temp[0] + "/" + temp[2]  # format: MM/DD/YYYY
        data_row = trans_date
        data_row += g_Delim
        data_row += action
        data_row += g_Delim
        data_row += '"' + description + '"'  # double quote to handle text with commas
        data_row += g_Delim
        if net_amount.startswith('-'):
            data_row += net_amount[1:]  # strip off the negative sign
            data_row += g_Delim
        else:
            data_row += g_Delim
            data_row += net_amount

        logging.debug("Data row: %s" % (data_row))
        return data_row
    except Exception as e:
        msg = str(e)
        logging.error("*****Error in parseTDRow. p_row: %s  Error: %s" % (p_row, msg))
        return None


# END parseTDRow

def processTDFile(p_file):
    global g_Delim
    global g_LoadedPath

    file_data = []
    file = p_file

    try:
        # parse file and generate converted rows
        logging.info("***PROCESSING FILE: %s" % (file))
        with open(file) as csvfile:
            filebuf = csv.reader(csvfile, delimiter=g_Delim)

            for cols in filebuf:
                # print cols
                # skip header rows
                if cols[0].find('As of') >= 0:
                    continue
                elif cols[0].find('Account') >= 0:
                    continue
                elif cols[0].find('Trade') >= 0:
                    continue
                elif len(cols) == 2:  # empty row
                    continue
                # process data rows
                else:
                    data = parseTDRow(cols)
                    if data is None:
                        logging.warning("Error parsing data row")
                        return
                    else:
                        file_data.append(data)
                    # end if
                # end if parse row
            # end for
            logging.info("File: %s  %d rows processed" % (file, len(file_data)))

            # output converted rows
            output_filename = generateOutputFilename(file)
            if output_filename is not None:
                writeFile(output_filename, file_data)
            else:
                logging.warning("WARNING: Invalid output filename***")
            # end if
            file_data = []
        # end with
        moveFile(file, g_LoadedPath)
    except Exception as e:
        msg = str(e)
        logging.error(
            "*****Error in processTDFile. Current file: %s  Error: %s" % (file, msg))
# end processTDFile


def parseEQRow(p_row):
    global g_Delim

    try:
        # process data row
        trans_date = p_row[0].strip()
        description = p_row[1].strip()
        amount = p_row[2].strip()
        # debit_amount = p_row[3].strip()
        balance = p_row[3].strip()
        payee = '""'

        # build output row
        """
            Looks like trans_date format changed. Now DD MMM YYYY
        """
        temp = trans_date.split()
        trans_date = str(g_Months.index(temp[1].upper()) + 1).zfill(2)
        #get index of trans month from months list, then left pad with zero
        trans_date += "/" + temp[0] + "/" + temp[2]  # format: MM/DD/YYYY

        data_row = trans_date
        data_row += g_Delim
        data_row += payee
        data_row += g_Delim
        data_row += '"' + description + '"'  # double quote to handle text with commas
        data_row += g_Delim
        amount = amount.replace("$", "").replace('"', "").replace(",", "")  # get rid of $ or " or ,
        if amount.startswith('-'):
            data_row += amount[1:]
            data_row += g_Delim
        else:
            data_row += g_Delim
            data_row += amount
        # end if

        logging.debug("Data row: %s" % (data_row))
        return data_row
    except Exception as e:
        msg = str(e)
        logging.error("*****Error in parseEQRow. p_row: %s Error: %s" % (p_row, msg))
        return None


# END parseEQRow

def processEQFile(p_file):
    global g_Delim
    global g_LoadedPath

    file_data = []
    file = p_file

    try:
        # parse file and generate converted rows
        logging.info("***PROCESSING FILE: %s" % (file))
        with open(file) as csvfile:
            filebuf = csv.reader(csvfile, delimiter=g_Delim)

            for cols in filebuf:
                logging.debug("Processing row: %s" % (cols))
                # skip header rows
                if len(cols) == 0:  # empty row
                    continue
                elif cols[0].find('Date') >= 0:
                    continue
                elif len(cols[0].strip()) == 0:  # empty row
                    continue
                # process data rows
                else:
                    data = parseEQRow(cols)
                    if data is None:
                        logging.warning("Error parsing data row")
                        return
                    else:
                        file_data.append(data)
                    # end if
                # end if parse row
            # end for
            logging.info("File: %s  %d rows processed" % (file, len(file_data)))

            # output converted rows
            output_filename = generateOutputFilename(file)
            if output_filename is not None:
                writeFile(output_filename, file_data)
            else:
                logging.warning("WARNING: Invalid output filename***")
            # end if
            file_data = []
        # end with
        moveFile(file, g_LoadedPath)
    except Exception as e:
        msg = str(e)
        logging.error(
            "*****Error in processEQFile. urrent file: %s  Error: %s" % (file, msg))

# end processEQFile


def parseEQTRow(p_row):
    global g_Delim

    try:
        # process data row
        trans_date = p_row[0].strip()
        description = p_row[1].strip()
        amount = p_row[2].strip()
        # debit_amount = p_row[3].strip()
        # balance = p_row[4].strip()
        payee = '""'

        # build output row
        """
            Looks like trans_date format changed. Now YYYY-MM-DD
            which should be acceptable as is
        """
        # temp = trans_date.split("-")
        # trans_date = str(g_Months.index(temp[1].upper()) + 1).zfill(2)
        # #get index of trans month from months list, then left pad with zero
        # trans_date += "/" + temp[0] + "/" + temp[2]  # format: MM/DD/YYYY

        data_row = trans_date
        data_row += g_Delim
        data_row += payee
        data_row += g_Delim
        data_row += '"' + description + '"'  # double quote to handle text with commas
        data_row += g_Delim
        amount = amount[1:].strip()  # strip off dollar sign
        if amount.startswith('-'):
            data_row += amount[1:]
            data_row += g_Delim
        else:
            data_row += g_Delim
            data_row += amount
        # end if

        logging.debug("Data row: %s" % (data_row))
        return data_row
    except Exception as e:
        msg = str(e)
        logging.error("*****Error in parseEQTRow. p_row: %s Error: %s" % (p_row, msg))
        return None


# END parseEQRow

def processEQTFile(p_file):
    global g_Delim
    global g_LoadedPath

    file_data = []
    file = p_file

    try:
        # parse file and generate converted rows
        logging.info("***PROCESSING FILE: %s" % (file))
        with open(file) as csvfile:
            filebuf = csv.reader(csvfile, delimiter=g_Delim)

            for cols in filebuf:
                logging.debug("Processing row: %s" % (cols))
                # skip header rows
                if len(cols) == 0:  # empty row
                    continue
                elif cols[0].find('Date') >= 0:
                    continue
                elif len(cols[0].strip()) == 0:  # empty row
                    continue
                # process data rows
                else:
                    data = parseEQTRow(cols)
                    if data is None:
                        logging.warning("Error parsing data row")
                        return
                    else:
                        file_data.append(data)
                    # end if
                # end if parse row
            # end for
            logging.info("File: %s  %d rows processed" % (file, len(file_data)))

            # output converted rows
            output_filename = generateOutputFilename(file)
            if output_filename is not None:
                writeFile(output_filename, file_data)
            else:
                logging.warning("WARNING: Invalid output filename***")
            # end if
            file_data = []
        # end with
        moveFile(file, g_LoadedPath)
    except Exception as e:
        msg = str(e)
        logging.error(
            "*****Error in processEQTFile. Current file: %s  Error: %s" % (file, msg))

# end processEQFile


def parseMBRow(p_row):
    global g_Delim

    try:
        # process data row
        post_date = p_row[1].strip()
        description = p_row[2].strip()
        amount = p_row[4].strip()
        payee = description

        # build output row
        logging.debug('Trans date: %s' % (post_date))

        data_row = post_date
        data_row += g_Delim
        data_row += '"' + payee + '"'
        data_row += g_Delim
        data_row += '"' + description + '"'  # double quote to handle text with commas
        data_row += g_Delim
        amount = amount.replace("$", "").replace('"', "").replace(",", "")  # get rid of $ or " or ,
        if amount.startswith('-'):
            data_row += g_Delim
            data_row += amount[1:]
        else:
            data_row += amount
            data_row += g_Delim
        # end if
        logging.debug("Data row: %s" % (data_row))
        return data_row
    except Exception as e:
        msg = str(e)
        logging.error("*****Error in parseMBRow. p_row: %s Error: %s" % (p_row, msg))
        return None
# END parseMBRow


def processMBFile(p_file):
    global g_Delim
    global g_LoadedPath

    file_data = []
    file = p_file

    try:
        # parse file and generate converted rows
        logging.info("***PROCESSING FILE: %s" % (file))
        with open(file) as csvfile:
            filebuf = csv.reader(csvfile, delimiter=g_Delim)

            count = 1
            for cols in filebuf:
                logging.debug(f"Current row is {cols}")
                if count == 1:
                    row = []
                    row += [cols[0]]
                    count += 1
                    continue
                elif count < 5:
                    row += [cols[0]]
                    count += 1
                    continue
                else:
                    row += [cols[0]]
                    count = 1
                # END IF

                logging.debug("Processing row: %s" % (row))
                logging.debug("Cols length: %s" % (len(row)))
                # skip header rows
                if len(row) == 0:  # empty row
                    continue
                # elif cols[0].find('Date') >= 0:
                #     continue
                # elif len(cols[0].strip()) == 0:     #empty row
                #     continue
                # process data rows
                else:
                    data = parseMBRow(row)
                    if data is None:
                        logging.warning("Error parsing data row")
                        return
                    else:
                        file_data.append(data)
                    # end if
                # end if parse row
            # end for
            logging.info("File: %s  %d rows processed" % (file, len(file_data)))

            # output converted rows
            output_filename = generateOutputFilename(file)
            if output_filename is not None:
                writeFile(output_filename, file_data)
            else:
                logging.warning("WARNING: Invalid output filename***")
            # end if
            file_data = []
        # end with
        moveFile(file, g_LoadedPath)
    except Exception as e:
        msg = str(e)
        logging.error(
            "*****Error in processMBFile. Current file: %s  Error: %s" % (file, msg))

# end processMBFile


def parseMBSRow(p_row):
    global g_Delim

    try:
        # process data row
        post_date = p_row[0].strip()
        description = p_row[1].strip()
        amount = p_row[3].strip()
        payee = description

        # build output row
        logging.debug('Trans date: %s' % (post_date))

        data_row = post_date
        data_row += g_Delim
        data_row += '"' + payee + '"'
        data_row += g_Delim
        data_row += '"' + description + '"'  # double quote to handle text with commas
        data_row += g_Delim
        amount = amount.replace("$", "").replace('"', "").replace(",", "")  # get rid of $ or " or ,
        if amount.startswith('-'):
            data_row += amount[1:]
            data_row += g_Delim
        else:
            data_row += g_Delim
            data_row += amount
        # end if
        logging.debug("Data row: %s" % (data_row))
        return data_row
    except Exception as e:
        msg = str(e)
        logging.error("*****Error in parseMBRow. p_row: %s Error: %s" % (p_row, msg))
        return None

# END parseMBRow


def processMBSFile(p_file):
    global g_Delim
    global g_LoadedPath

    file_data = []
    file = p_file

    try:
        # parse file and generate converted rows
        logging.info("***PROCESSING FILE: %s" % (file))
        with open(file) as csvfile:
            filebuf = csv.reader(csvfile, delimiter=g_Delim)

            for cols in filebuf:
                logging.debug("Processing row: %s" % (cols))
                logging.debug("Cols length: %s" % (len(cols)))
                # skip header rows
                if len(cols) == 0:  # empty row
                    continue
                elif cols[0].find('Date') >= 0:
                    continue
                elif len(cols[0].strip()) == 0:  # empty row
                    continue
                # process data rows
                else:
                    data = parseMBSRow(cols)
                    if data is None:
                        logging.warning("Error parsing data row")
                        return
                    else:
                        file_data.append(data)
                    # end if
                # end if parse row
            # end for
            logging.info("File: %s  %d rows processed" % (file, len(file_data)))

            # output converted rows
            output_filename = generateOutputFilename(file)
            if output_filename is not None:
                writeFile(output_filename, file_data)
            else:
                logging.warning("WARNING: Invalid output filename***")
            # end if
            file_data = []
        # end with
        moveFile(file, g_LoadedPath)
    except Exception as e:
        msg = str(e)
        logging.error(
            "*****Error in processMBSFile. Current file: %s  Error: %s" % (file, msg))

# end processMBFile


def parseQTRow(p_row):
    global g_Delim

    try:
        # process data row
        trans_date = p_row[0].strip()[:10]
        payee = p_row[2].strip()
        description = p_row[4].strip()
        amount = p_row[9].strip()

        # build output row
        data_row = trans_date
        data_row += g_Delim
        data_row += '"' + payee + '"'
        data_row += g_Delim
        data_row += '"' + description + '"'  # double quote to handle text with commas
        data_row += g_Delim
        amount = amount.replace("$", "").replace('"', "").replace(",", "")  # get rid of $ or " or ,
        num_amount = Decimal(amount) * -1  # need to flip the sign for YNAB account
        amount = str(num_amount)
        data_row += amount

        logging.debug("Data row: %s" % (data_row))
        return data_row
    except Exception as e:
        msg = str(e)
        logging.error("*****Error in parseMBRow. p_row: %s Error: %s" % (p_row, msg))
        return None

# END parseMBRow


def processQTFile(p_file):
    global g_Delim
    global g_LoadedPath

    file_data = []
    file = xlrd.open_workbook(p_file)
    worksheet = file.sheet_by_index(0)

    try:
        # parse file and generate converted rows
        logging.info("***PROCESSING FILE: %s with %d rows" % (p_file, worksheet.nrows))

        for row in range(worksheet.nrows):
            logging.debug("Processing row: %s" % (worksheet.row(row)))
            # skip header rows
            if row == 0:
                continue
            # process data rows
            else:
                logging.debug("Row values: %s" % (worksheet.row_values(row)))
                data = parseQTRow(worksheet.row_values(row))

                if data is None:
                    logging.warning("Error parsing data row")
                    return
                else:
                    file_data.append(data)
                # end if
            # end if parse row
        # end for
        logging.info("File: %s  %d rows processed" % (p_file, len(file_data)))

        # output converted rows
        output_filename = generateOutputFilename(p_file)
        if output_filename is not None:
            writeFile(output_filename, file_data)
        else:
            logging.warning("WARNING: Invalid output filename***")
        # end if
        file_data = []

        moveFile(p_file, g_LoadedPath)

    except Exception as e:
        msg = str(e)
        logging.error(
            "*****Error in processQTFile. Current file: %s  Error: %s" % (p_file, msg))

# end processMBFile

#####
# Main
#####


def main():

    global g_InputPath
    global g_OutputPath
    global g_LoadedPath

    try:
        args = getArgs()
        if args is None:
            logging.error("Unable to retrieve command line args - EXITING")
            return
        # END IF

        g_InputPath = args[0]
        g_OutputPath = args[1]
        g_LoadedPath = g_InputPath + r"/loaded"

        # get file list using FileTypes filters
        filelist = getFileList(g_InputPath)
        if filelist is None:
            logging.info('*****PROGRAM END')
            return
        # end if

        filelist = [file for file in filelist for prefix in g_FileTypes if prefix + "_" in file.upper()]
        logging.debug("Filelist: %s" % (filelist))

        for file in filelist:
            if "TD_" in file.upper():
                processTDFile(file)
            elif "EQ_" in file.upper():
                processEQFile(file)
            elif "EQT_" in file.upper():
                processEQTFile(file)
            elif "MB_" in file.upper():
                processMBFile(file)
            elif "MBS_" in file.upper():
                processMBSFile(file)
            elif "QU_" in file.upper():
                processQTFile(file)
            # end if
        # end for
    except Exception as e:
        msg = str(e)
        logging.error("*****Error in Main. input path: %s  output path: %s  Error: %s" % (
            g_InputPath, g_OutputPath, msg))


# end main()

#####
# Globals
#####
g_LoggingLevel = logging.INFO

logging.basicConfig(level=g_LoggingLevel,
                    format="%(levelname)s: %(asctime)s %(message)s", datefmt="%m/%d/%Y %I:%M:%S %p")

logging.info('*****PROGRAM START')

g_Months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
g_Delim = ','
g_CSVHeader = 'Date,Payee,Memo,Outflow,Inflow\n'
g_FileTypes = ['TD', 'EQ', 'MB', 'MBS', 'QU', 'EQT']

g_InputPath = None
g_OutputPath = None
g_LoadedPath = None

if __name__ == '__main__':

    main()

# end if main()

logging.info('*****PROGRAM END')
