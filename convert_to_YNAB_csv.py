import os
import sys
import csv
import logging
import glob
from datetime import datetime


##########
## Functions
##########
def getFileList(p_path):
    try:
        logging.info("***GETTING INPUT DIR FILE LIST")
        filelist = [file for file in glob.glob(p_path + r"\*") if
                    not os.path.isdir(file) and not "_OUTPUT" in file.upper()]
        logging.debug("Found %s files in directory %s" % (len(filelist), p_path))
        if len(filelist) > 0:
            return filelist
        else:
            logging.warning("***WARNING: no files found in path %s matching search criteria" % (p_path))
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
            output_file = g_OutputPath + "\\" + p_filename
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
        logging.error("*****Error in writeFile. p_filename: %s  p_rows: %s  Error: %s" % (p_filename, p_rows, msg))


# end writeFile

def generateOutputFilename(p_filename):
    try:
        filename = p_filename.split(".")[0].split("\\")[-1]  ##strips the raw filename out of file string
        current_datetime = datetime.strftime(datetime.now(), "%Y%m%d%H%M%S")
        output_filename = filename + "_output_" + current_datetime + ".csv"
        logging.debug("Output filename: %s" % (output_filename))
        return output_filename
    except Exception as e:
        msg = str(e)
        logging.error("*****Error in generateOutputFilename. p_filename: %s  Error: %s" % (p_filename, msg))
        return None


# end generateOutputFilename

def parseTDRow(p_row):
    global g_Months
    global g_Delim

    try:
        # print 'process data row'
        trade_date = p_row[0].strip()
        # settle_date = p_row[1].strip()
        description = p_row[2].strip()
        action = p_row[3].strip()
        # quantity = p_row[4].strip()
        # price = p_row[5].strip()
        # commission = p_row[6].strip()
        net_amount = p_row[7].strip()

        ##build output row
        temp = trade_date.split()
        trans_date = str(g_Months.index(temp[1]) + 1).zfill(2)
        ##get index of trans month from months list, then left pad with zero
        trans_date += "/" + temp[0] + "/" + temp[2]  # format: MM/DD/YYYY
        data_row = trans_date
        data_row += g_Delim
        data_row += action
        data_row += g_Delim
        data_row += '"' + description + '"'  ##double quote to handle text with commas
        data_row += g_Delim
        if net_amount.startswith('-'):
            data_row += net_amount[1:]  ##strip off the negative sign
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

def processTDFile(p_filelist):
    global g_Delim

    file_data = []

    try:
        for file in p_filelist:
            ##parse file and generate converted rows
            logging.info("***PROCESSING FILE: %s" % (file))
            with open(file) as csvfile:
                filebuf = csv.reader(csvfile, delimiter=g_Delim)

                for cols in filebuf:
                    # print cols
                    ##skip header rows
                    if cols[0].find('As of') >= 0:
                        continue
                    elif cols[0].find('Account') >= 0:
                        continue
                    elif cols[0].find('Trade') >= 0:
                        continue
                    elif len(cols) == 2:  ##empty row
                        continue
                    ##process data rows
                    else:
                        file_data.append(parseTDRow(cols))
                    # end if parse row
                # end for
                logging.info("File: %s  %d rows processed" % (file, len(file_data)))

                ##output converted rows
                output_filename = generateOutputFilename(file)
                if output_filename != None:
                    writeFile(output_filename, file_data)
                else:
                    logging.warning("WARNING: Invalid output filename***")
                # end if
                file_data = []
            # end with
        # end for
    except Exception as e:
        msg = str(e)
        logging.error(
            "*****Error in processTDFile. p_filelist: %s  Current file: %s  Error: %s" % (p_filelist, file, msg))


# end processTDFile

def parseEQRow(p_row):
    global g_Delim

    try:
        # print 'process data row'
        trans_date = p_row[0].strip()
        description = p_row[1].strip()
        credit_amount = p_row[2].strip()
        debit_amount = p_row[3].strip()
        # balance = p_row[4].strip()
        payee = '""'

        ##build output row
        trans_date = trans_date[4:6] + '/' + trans_date[6:8] + '/' + trans_date[:4]
        # logging.debug('Trans date: %s'%(trans_date))

        data_row = trans_date
        data_row += g_Delim
        data_row += payee
        data_row += g_Delim
        data_row += '"' + description + '"'  ##double quote to handle text with commas
        data_row += g_Delim
        data_row += debit_amount[1:].strip()  ##strip off dollar sign
        data_row += g_Delim
        data_row += credit_amount[1:].strip()  ##strip off dollar sign

        logging.debug("Data row: %s" % (data_row))
        return data_row
    except Exception as e:
        msg = str(e)
        logging.error("*****Error in parseEQRow. p_row: %s Error: %s" % (p_row, msg))
        return None


# END parseEQRow

def processEQFile(p_filelist):
    global g_Delim

    file_data = []

    try:
        for file in p_filelist:
            ##parse file and generate converted rows
            logging.info("***PROCESSING FILE: %s" % (file))
            with open(file) as csvfile:
                filebuf = csv.reader(csvfile, delimiter=g_Delim)

                for cols in filebuf:
                    logging.debug("Processing row: %s" % (cols))
                    logging.debug("Cols length: %s" % (len(cols)))
                    ##skip header rows
                    if len(cols) == 0:                  ##empty row
                        continue
                    elif cols[0].find('Date') >= 0:
                        continue
                    elif len(cols[0].strip()) == 0:     ##empty row
                        continue
                    ##process data rows
                    else:
                        file_data.append(parseEQRow(cols))
                    # end if parse row
                # end for
                logging.info("File: %s  %d rows processed" % (file, len(file_data)))

                ##output converted rows
                output_filename = generateOutputFilename(file)
                if output_filename != None:
                    writeFile(output_filename, file_data)
                else:
                    logging.warning("WARNING: Invalid output filename***")
                # end if
                file_data = []
            # end with
        # end for
    except Exception as e:
        msg = str(e)
        logging.error(
            "*****Error in processEQFile. p_filelist: %s  Current file: %s  Error: %s" % (p_filelist, file, msg))


# end processTDFile

##########
## Main
##########

def main():

    filetype = sys.argv[3]

    logging.basicConfig(level=g_LoggingLevel, format='LOG: %(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')

    logging.info('*****PROGRAM START')

    logging.debug("Command line arguments used: input_path: %s  output_path: %s  filetype: %s" % (
        g_InputPath, g_OutputPath, filetype))

    try:
        ##get file list using filetype as filter
        filelist = [file for file in getFileList(g_InputPath) if filetype.upper() + "_" in file.upper()]
        logging.debug("Filelist: %s" % (filelist))

        if len(filelist) > 0:
            if filetype.upper() == "TD":
                processTDFile(filelist)
            elif filetype.upper() == "EQ":
                processEQFile(filelist)
            else:
                logging.warning("WARNING: filetype %s is not recognized" % (filetype))
        else:
            logging.warning("WARNING: no files found to process for filetype %s***" % (filetype))
        # end if

    except Exception as e:
        msg = str(e)
        logging.error("*****Error in Main. input path: %s  output path: %s  filetype: %s  Error: %s" % (
            g_InputPath, g_OutputPath, filetype, msg))

    logging.info('*****PROGRAM END')
#end main()

##########
## Globals
##########
g_LoggingLevel = logging.DEBUG

g_Months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
g_Delim = ','
g_CSVHeader = 'Date,Payee,Memo,Outflow,Inflow\n'

g_InputPath = sys.argv[1]
g_OutputPath = sys.argv[2]

if __name__ == '__main__':

    main()
    
# end if main()
