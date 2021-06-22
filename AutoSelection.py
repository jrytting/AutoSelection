from __future__ import with_statement

import contextlib
import datetime

import xlsxwriter as xlsxwriter

try:
    from urllib.parse import urlencode
except ImportError:
    from urllib import urlencode
try:
    from urllib.request import urlopen
except ImportError:
    from urllib2 import urlopen

from bs4 import BeautifulSoup
from selenium import webdriver

from datetime import datetime
import winsound
import pandas as pd
import time
import sys
import logging

import ControlGroupCollection
import ControlGroupRecords
import GoogleSearch


product_collection = ControlGroupCollection.ControlGroupMasterDataSet
retry_collection = pd.DataFrame()
google_search = GoogleSearch


total_records_from_file = 0
zero_product_count = 0
attribute_errors = 0
start_control_timer = 0

"""
def duplicate_check(current_control_group, control, source_description):
    global product_collection
    if product_collection.empty:
        return False

    result = product_collection.loc[product_collection['Searched For This Item'] == source_description]

    if result.empty:  # No Duplicate Record Found
        return False

    if len(result.index) > 1:     # when multiple records are returned, use data from last row in selection
        try:
            current_control_group.control_number = control
            current_control_group.search_for_description = source_description
            current_control_group.found_description = result["Selected Item Description"][0:1]
            current_control_group.found_price = float(result["Price"][0:1])
            current_control_group.low_price = float(result["Low"][0:1])
            current_control_group.high_price = float(result["High"][0:1])
            current_control_group.source_url = result["URL Used For Search"][0:1]
            current_control_group.found_search_url = result["Selected Item URL"][0:1]
        except:
            return False
    else:   # only 1 duplicate record found in the control_group
        try:
            current_control_group.control_number = control
            current_control_group.search_for_description = source_description
            current_control_group.found_description = result["Selected Item Description"]
            current_control_group.found_price = float(result["Price"])
            current_control_group.low_price = float(result["Low"])
            current_control_group.high_price = float(result["High"])
            current_control_group.source_url = result["URL Used For Search"]
            current_control_group.found_search_url = result["Selected Item URL"]
        except:
            return False

    current_control_group.add_record()
    return True
"""


class ReCaptchaError(Exception):
    """Exception raised when ReCaptcha Screen Encountered

    Attributes:
        salary -- input salary which caused the error
        message -- explanation of the error
    """
    def __init__(self, message="reCaptcha Screen is blocking our search"):
        self.message = message
        super().__init__(self.message)


class NoResultsReturnedError(Exception):
    """Exception raised when ReCaptcha Screen Encountered

    Attributes:
        salary -- input salary which caused the error
        message -- explanation of the error
    """
    def __init__(self, message="No Results Returned for search"):
        self.message = message
        super().__init__(self.message)
# --------------------------  End of Exception Definitions --------------------------------

def check_and_handle_duplicate_records(product_collection, control_group_df, new_control_number, new_search_description):

    if not product_collection.empty:
        result = product_collection.loc[product_collection['Searched For This Item'] == new_search_description]

        if not result.empty:
            if len(result.index) > 1:  # when multiple records are returned, use data from last row in selection
                try:
                    control_group_df.control_number = new_control_number
                    control_group_df.search_for_description = new_search_description
                    control_group_df.found_description = result["Selected Item Description"][0:1]
                    control_group_df.found_price = float(result["Price"][0:1])
                    control_group_df.low_price = float(result["Low"][0:1])
                    control_group_df.high_price = float(result["High"][0:1])
                    control_group_df.source_url = result["URL Used For Search"][0:1]
                    control_group_df.found_search_url = result["Selected Item URL"][0:1]
                except Exception:
                    print('Warning -- Error checking for duplicates: result.index > 1 ')
                    e = sys.exc_info()[0]
                    print("<p>Error: %s</p>" % e)
                    exit(2)
            else:  # only 1 duplicate record found in the control_group
                try:
                    control_group_df.control_number = new_control_number
                    control_group_df.search_for_description = new_search_description
                    control_group_df.found_description = result["Selected Item Description"]
                    control_group_df.found_price = float(result["Price"])
                    control_group_df.low_price = float(result["Low"])
                    control_group_df.high_price = float(result["High"])
                    control_group_df.source_url = result["URL Used For Search"]
                    control_group_df.found_search_url = result["Selected Item URL"]
                except Exception:
                    print('Warning -- Error checking for duplicates: result.index = 1 ')
                    e = sys.exc_info()[0]
                    print("<p>Error: %s</p>" % e)
                    exit(2)

            # add this record to the control group control_group
            control_group_df.add_record_to_control_group()
            return True
    return False


def read_file(input_file_name):
    """
    Read data from an excel spreadsheet and initiate the web search for those items
    :param input_file_name:
    """
    global start_control_timer
    global total_records_from_file
    global product_collection
    global google_search
    cgc = ControlGroupCollection.ControlGroupMasterDataSet()
    product_collection = cgc.get_collections()

    gs = GoogleSearch.GoogleRecordSearch()
    last_control_number = ''
    source_data_file = ''
    control_group = ''
    new_control_number = ''
    new_search_for_description = ''

    try:
        source_data_file = pd.read_excel(input_file_name)
    except Exception:
        print("Error Reading Spreadsheet \n " +
              "-- Control # must be in column 1 and Description of item to search for must be in column 2.\n" +
              "-- Also ensure the Source File is not open by another application.")
        e = sys.exc_info()[0]
        print("<p>Error: %s</p>" % e)
        print("Program will terminate. ")
        exit(1)

    first_pass = True

    # Read/Process the source file records one at a time
    for index, row in source_data_file.iterrows():
        start_control_timer = time.perf_counter()   # Start Timer for each Control #
        total_records_from_file += 1

        # Create a new control group object
        control_group = ControlGroupRecords.ControlGroupDataSet()

        new_control_number = row[0]
        new_search_for_description = row[1]

        if not first_pass:
            if check_and_handle_duplicate_records(product_collection, control_group, new_control_number, new_search_for_description):
                product_collection = control_group.add_selected_control_group_record_to_collection(product_collection)
                continue  # if duplicate record found, -- no need to look up again

        url_to_search = "https://www.google.com/search?tbm=shop&q=" + row[1]

        if first_pass:
            last_control_number = new_control_number
            first_pass = False

        #if last_control_number != new_control_number:
            # Choose the correct record from the control group and add to master data set
            #control_group.add_selected_control_group_record_to_collection(product_collection)
            #control_group = ControlGroupRecords.ControlGroupDataSet()  #re-init control-group

        control_group.control_number            = new_control_number
        control_group.search_for_description    = new_search_for_description
        control_group.source_url                = url_to_search
        last_control_number                     = new_control_number

        # Search the Web for this data
        try:
            gs.item_search(control_group)
        except ReCaptchaError:
            logging.info("**Control Number: %s was blocked by reCaptcha Screen", control_group.control_number)
        except NoResultsReturnedError:
            logging.info(("**No Results were returned for Control #: %s --Description: %s",
                          control_group.control_number, control_group.search_for_description))
        finally:
            continue

        product_collection = control_group.select_record_from_control_group(product_collection)

    check_and_handle_duplicate_records(control_group.get_df_collection(), new_control_number, new_search_for_description)
    #pcg.process_control_group(control_group.get_df_collection())  # Used to write the last record to the file


if __name__ == '__main__':
    logging.basicConfig(format='%(levelname)s:%(message)s',
                        filename='C:/Users/JRytt/Documents/Brad V/AutoSelection.log',
                        level=logging.INFO)
    logging.info('Started')
    startTimer = time.perf_counter()
    read_file('C:/Users/JRytt/Documents/Brad V/Price Sample v3(wip).xlsx')

    print(f'Zero Product Counter = {zero_product_count}')
    print(f'Total Records from File = {total_records_from_file}')
    print(f'Total Attribute Errors =  {attribute_errors}')

    column_names = ["Control #",
                    "Searched For This Item",
                    "URL Used For Search",
                    "Selected Item Description",
                    "Selected Item URL",
                    "Price",
                    "Low",
                    "High"
                    ]
    product_collection = product_collection.reindex(columns=column_names)


    # Create a Pandas Excel writer using XlsxWriter as the engine.
    now = datetime.now()  # current date and time
    date_time = now.strftime("%H_%M_%S")
    writer = pd.ExcelWriter("C:/Users/JRytt/Documents/Brad V/AutoSelection_Output_" + date_time + ".xlsx",
                            engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    product_collection.to_excel(writer, sheet_name='Sheet1', index=False)

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

    # product_collection.to_excel('C:/Users/JRytt/Documents/Brad V/JobComplete_v4.xlsx', index=False)
    stopTimer = time.perf_counter()
    print(f"Program run time: {stopTimer - startTimer:0.4f} seconds")
    exit()

