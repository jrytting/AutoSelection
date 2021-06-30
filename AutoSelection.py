from __future__ import with_statement

import datetime
import os

try:
    from urllib.parse import urlencode
except ImportError:
    from urllib import urlencode
try:
    from urllib.request import urlopen
except ImportError:
    from urllib2 import urlopen

from datetime import datetime
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

product_collection = pd.DataFrame() # ControlGroupCollection.ControlGroupMasterDataSet
retry_collection = pd.DataFrame()
google_search = GoogleSearch


total_records_from_file = 0
zero_product_count = 0
attribute_errors = 0
start_control_timer = 0


def check_and_handle_duplicate_records(main_collection, control_group_df, new_control_number, new_search_description):

    try:
        result = pd.DataFrame()
        if main_collection is not None and not main_collection.empty:
            result = main_collection.loc[main_collection['Searched For This Item'] == new_search_description]

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
                except Exception as e:
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
                except Exception as e:
                    print('Warning -- Error checking for duplicates: result.index = 1 ')
                    e = sys.exc_info()[0]
                    print("<p>Error: %s</p>" % e)
                    exit(2)

            # add this record to the control group control_group
            try:
                control_group_df.add_record_to_control_group()
            except NoResultsReturnedError:
                logging.warning("Error adding this record to control group: " + control_group_df.control_number)
            return True

    except AttributeError as e:
        logging.warning("Checking for Duplicates: AttributeError: 'NoneType' object has no attribute 'empty' \
            -- returning Skipping Duplicate Check")
        e = sys.exc_info()[0]
        print("<p>Error: %s</p>" % e)
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
    global retry_collection
    global google_search
    cgc = ControlGroupCollection.ControlGroupMasterDataSet()
    # product_collection = pd.DataFrame() #cgc.get_collections()

    gs = GoogleSearch.GoogleRecordSearch()
    bs = BingSearch.BingSearch()
    source_data_file = ''
    control_group = ''
    new_control_number = ''
    new_search_for_description = ''
    use_bing_search = True

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

        if use_bing_search:
            url_to_search = "https://www.bing.com/shop?q=" + new_search_for_description
            search_engine = bs
        else:
            url_to_search = "https://www.google.com/search?tbm=shop&q=" + new_search_for_description
            search_engine = gs

        control_group.control_number            = new_control_number
        control_group.search_for_description    = new_search_for_description
        control_group.source_url                = url_to_search

        if not check_and_handle_duplicate_records(product_collection,
                                                  control_group,
                                                  new_control_number,
                                                  new_search_for_description):

            perform_search(control_group, search_engine)

        product_collection = control_group.select_record_from_control_group(product_collection)

    if retry_collection.empty:
        pass
    else:
        if use_bing_search:
            use_bing_search = False
        else:
            use_bing_search = True

        for index, row in retry_collection.iterrows():
            # Create a new control group object
            control_group = ControlGroupRecords.ControlGroupDataSet()

            new_control_number = row['Control #']
            new_search_for_description = row['Searched For This Item']

            if use_bing_search:
                search_url = "https://www.bing.com/shop?q=" + new_search_for_description
                search_engine = bs
            else:
                search_url = "https://www.google.com/search?tbm=shop&q=" + new_search_for_description
                search_engine = gs

            control_group.control_number = new_control_number
            control_group.search_for_description = new_search_for_description
            control_group.source_url = search_url

            perform_search(control_group, search_engine)
            product_collection = control_group.select_record_from_control_group(product_collection)


def perform_search(control_group_data, search_engine):
    global retry_collection
    global product_collection

    # Search the Web for this data
    try:
        search_engine.item_search(control_group_data)
    except ReCaptchaError as e:
        logging.info("**Control Number: %s was blocked by reCaptcha Screen", control_group_data.control_number)
    except NoResultsReturnedError as e:
        control_group_data.found_description = "No Search Results were returned for this item"
        control_group_data.found_price = float("0.00")
        control_group_data.low_price = float("0.00")
        control_group_data.high_price = float("0.00")
        control_group_data.found_search_url = "No Search Results were returned for this item"
        try:
            retry_record = {"Control #": control_group_data.control_number,
                            "Searched For This Item": control_group_data.search_for_description}
            retry_collection = retry_collection.append(retry_record, ignore_index=True)
        except NoResultsReturnedError as e:
            logging.warning("Error Thrown trying to add the 'No Search Results Dummy Record'")
            pass

        logging.warning("**No Results were returned for Control #: %s --Description: %s",
                        control_group_data.control_number,
                        control_group_data.search_for_description)


def job_complete(path):

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    date_time = datetime.now().strftime("%H_%M_%S")
    writer = pd.ExcelWriter(path + date_time + ".xlsx", engine='xlsxwriter')

    column_names = ["Control #",
                    "Searched For This Item",
                    "URL Used For Search",
                    "Selected Item Description",
                    "Selected Item URL",
                    "Price",
                    "Low",
                    "High"
                    ]
    product_collection.reindex(columns=column_names)

    # Convert the dataframe to an XlsxWriter Excel object.
    product_collection.to_excel(writer, sheet_name='Sheet1', index=False)

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()


if __name__ == '__main__':
    job_name = 'Product v5'
    source_file_name = "Price Sample v5"
    main_path = 'C:/Users/JRytt/Documents/Brad V/'
    source_file = main_path + job_name
    xcel_file_to_create = main_path + job_name + '/WIP/AutoSelection_Output_'
    temp_file_for_csv = main_path + job_name + '/WIP/Temp_CSV.csv'
    logging_file_name = main_path + job_name + '/WIP/AutoSelection.log'

    if not os.path.exists(job_name):
        print("Creating the correct paths before processing the files")
        try:
            if not os.path.exists(main_path + job_name + '/WIP'):
                new_path = os.path.join(main_path, job_name + '/WIP')
                os.makedirs(new_path)
        except OSError as e:
            print("Caught Exception Trying to make Directories for Job")
            exit(1)

    logging.basicConfig(format='%(levelname)s:%(message)s', filename=logging_file_name, level=logging.INFO)
    logging.info('Started: ' + datetime.now().strftime("%H_%M_%S"))
    startTimer = time.perf_counter()

    read_file(source_file + '/' + source_file_name + '.xlsx')
    job_complete(xcel_file_to_create)

    print(f'Zero Product Counter = {zero_product_count}')
    print(f'Total Records from File = {total_records_from_file}')
    print(f'Total Attribute Errors =  {attribute_errors}')

    stopTimer = time.perf_counter()
    print(f"Program run time: {stopTimer - startTimer:0.4f} seconds")
    exit()

