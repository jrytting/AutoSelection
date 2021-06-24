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
import BingSearch

from GoogleSearch import ReCaptchaError, NoResultsReturnedError
from BingSearch import ReCaptchaError, NoResultsReturnedError

product_collection = ControlGroupCollection.ControlGroupMasterDataSet
retry_collection = pd.DataFrame()
google_search = GoogleSearch


total_records_from_file = 0
zero_product_count = 0
attribute_errors = 0
start_control_timer = 0


def check_and_handle_duplicate_records(product_collection, control_group_df, new_control_number, new_search_description):

    try:
        if not product_collection.empty:
            result = product_collection.loc[product_collection['Searched For This Item'] == new_search_description]

        if not result.empty:
            if len(result.index) > 1:  # when multiple records are returned, use data from last row in selection
                try:
                    control_group_df.control_number = new_control_number
                    control_group_df.search_for_description = new_search_description
                    control_group_df.found_description = result["Selected Item Description"][0:1].values[0]
                    control_group_df.found_price = float(result["Price"][0:1])
                    control_group_df.low_price = float(result["Low"][0:1])
                    control_group_df.high_price = float(result["High"][0:1])
                    control_group_df.source_url = result["URL Used For Search"][0:1].values[0]
                    control_group_df.found_search_url = result["Selected Item URL"][0:1].values[0]
                except Exception:
                    print('Warning -- Error checking for duplicates: result.index > 1 ')
                    e = sys.exc_info()[0]
                    print("<p>Error: %s</p>" % e)
                    exit(2)
            else:  # only 1 duplicate record found in the control_group
                try:
                    control_group_df.control_number = new_control_number
                    control_group_df.search_for_description = new_search_description
                    control_group_df.found_description = result["Selected Item Description"].values[0]
                    control_group_df.found_price = float(result["Price"])
                    control_group_df.low_price = float(result["Low"])
                    control_group_df.high_price = float(result["High"])
                    control_group_df.source_url = result["URL Used For Search"].values[0]
                    control_group_df.found_search_url = result["Selected Item URL"].values[0]
                except Exception:
                    print('Warning -- Error checking for duplicates: result.index = 1 ')
                    e = sys.exc_info()[0]
                    print("<p>Error: %s</p>" % e)
                    exit(2)

            # add this record to the control group control_group
            try:
                control_group_df.add_record_to_control_group()
            except NoResultsReturnedError:
                logging.warning("Error adding this record to control group: " + control_group_df.control_number)
                pass
            return True
    except AttributeError as e:
        logging.warning("Checking for Duplicates: AttributeError: 'NoneType' object has no attribute 'empty' \
            -- returning Skipping Duplicate Check")
        return False
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
    bs = BingSearch.DuckDuckGoSearch()
    last_control_number = ''
    source_data_file = ''
    control_group = ''
    new_control_number = ''
    new_search_for_description = ''
    use_Bing_Search= True

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
            if check_and_handle_duplicate_records(product_collection,
                                                  control_group,
                                                  new_control_number,
                                                  new_search_for_description):
                try:
                    product_collection = control_group.add_selected_control_group_record_to_collection(product_collection)
                    continue  # if duplicate record found, -- no need to look up again
                except NoResultsReturnedError as e:
                    logging.warning("Failed to Add Record to Master Collection: " + control_group.control_number +
                                    "Processing will continue to the next record")
                    continue

        if use_Bing_Search:
            url_to_search = "https://www.bing.com/shop?q=" + new_search_for_description
        else:
            url_to_search = "https://www.google.com/search?tbm=shop&q=" + new_search_for_description

        if first_pass:
            last_control_number = new_control_number
            first_pass = False

        control_group.control_number            = new_control_number
        control_group.search_for_description    = new_search_for_description
        control_group.source_url                = url_to_search
        last_control_number                     = new_control_number

        # Search the Web for this data
        try:
            if use_Bing_Search:
                bs.item_search(control_group)
            else:
                gs.item_search(control_group)
        except ReCaptchaError as e:
            logging.info("**Control Number: %s was blocked by reCaptcha Screen", control_group.control_number)
            continue
        except NoResultsReturnedError as e:
            # Build the record and commit it to the dataset
            control_group.control_number = control_group.control_number
            control_group.search_for_description = control_group.search_for_description
            control_group.found_description = "No Search Results were returned for this item"
            control_group.found_price = float("0.00")
            control_group.low_price = float("0.00")
            control_group.high_price = float("0.00")
            control_group.source_url = control_group.source_url
            control_group.found_search_url = "No Search Results were returned for this item"
            try:
                product_collection = control_group.add_selected_control_group_record_to_collection(product_collection)
            except NoResultsReturnedError as e:
                logging.warning("Error Thrown trying to add the 'No Search Results Dummy Record'")
                pass

            logging.warning("**No Results were returned for Control #: %s --Description: %s",
                          control_group.control_number,
                          control_group.search_for_description)
            continue
        #finally:
            #continue

        product_collection = control_group.select_record_from_control_group(product_collection)

    check_and_handle_duplicate_records(product_collection,
                                       control_group.get_df_collection(),
                                       new_control_number,
                                       new_search_for_description)


if __name__ == '__main__':
    logging.basicConfig(format='%(levelname)s:%(message)s',
                        filename='C:/Users/JRytt/Documents/Brad V/Product v4/WIP/AutoSelection.log',
                        level=logging.INFO)
    logging.info('Started')
    startTimer = time.perf_counter()
    read_file('C:/Users/JRytt/Documents/Brad V/Product v4/WIP/Price Sample v4(wip).xlsx')

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
    writer = pd.ExcelWriter("C:/Users/JRytt/Documents/Brad V/Product v4/WIP/AutoSelection_Output_" + date_time + ".xlsx",
                            engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    product_collection.to_excel(writer, sheet_name='Sheet1', index=False)

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

    # product_collection.to_excel('C:/Users/JRytt/Documents/Brad V/Product v4/WIP/JobComplete_v4.xlsx', index=False)
    stopTimer = time.perf_counter()
    print(f"Program run time: {stopTimer - startTimer:0.4f} seconds")
    exit()

