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



# prepare the option for the chrome driver
options = webdriver.ChromeOptions()
options.add_argument('headless')

# start chrome browser
browser = webdriver.Chrome()  # options=options

product_collection = pd.DataFrame()
retry_collection = pd.DataFrame()

total_records_from_file = 0
zero_product_count = 0
attribute_errors = 0
start_control_timer = 0


class FoundDataSet:
    control_collection = pd.DataFrame()
    record = {}
    control_number = ''
    search_for_description = ''
    found_description = ''
    found_price = ''
    low_price = ''
    high_price = ''
    average_price = ''
    source_url = ''
    found_search_url = ''

    less_than_mean = 0
    greater_than_mean = 0

    def cleanup_prices(self, amount):
        amount = amount.split('.')[0]
        amount = amount.replace('$', "")
        amount = amount.replace(",", "")
        amount = amount.replace(" ", "")

        return int(amount)

    def create_record(self):
        self.record["Control #"] = self.control_number
        self.record["Searched For This Item"] = self.search_for_description
        self.record["Selected Item Description"] = self.found_description
        self.record["Price"] = self.found_price
        self.record["Low"] = self.low_price
        self.record["High"] = self.high_price
        self.record["URL Used For Search"] = self.source_url
        self.record["Selected Item URL"] = self.found_search_url

        self.add_record()

    def add_record(self):
        self.control_collection = self.control_collection.append(self.record, ignore_index=True)

    def get_df_collection(self):
        return self.control_collection

    def process_current_control_group(self):

        # Rule: Drop Highest and Lowest Amounts and then determine average price with remainder of elements
        # :return:

        if self.control_collection.empty:
            pass
        else:
            self.control_collection.sort_values(by=['Price'], inplace=True)  # sort by 'amount' column low to high
            self.control_collection = self.control_collection.iloc[1:, :]
            self.control_collection = self.control_collection.iloc[:-1, :]
            self.average_price = self.control_collection["Price"].mean()
            self.low_price = self.control_collection["Price"].min()
            self.high_price = self.control_collection["Price"].max()

            # noinspection PyCompatibility
            for index, row in self.control_collection.T.iteritems():

                if row["Price"] < self.average_price:
                    self.less_than_mean = row["Price"]

                    self.control_number = row["Control #"]
                    self.search_for_description = row["Searched For This Item"]
                    self.found_description = row["Selected Item Description"]
                    self.found_price = row["Price"]
                    self.source_url = row["URL Used For Search"]
                    self.found_search_url = row["Selected Item URL"]
                    continue
                elif row["Price"] > self.average_price:
                    self.greater_than_mean = row["Price"]

                    # df is ordered lowest to highest when we find a higher (or equal)
                    # Price we have the data we need

                if self.average_price - self.less_than_mean > self.greater_than_mean - self.average_price:
                    self.control_number = row["Control #"]
                    self.search_for_description = row["Searched For This Item"]
                    self.found_description = row["Selected Item Description"]
                    self.found_price = row["Price"]
                    self.source_url = row["URL Used For Search"]
                    self.found_search_url = row["Selected Item URL"]

                self.create_record()

# --------------------------------------  End of Class FoundDataSet  ----------------------------------------------


class ProcessControlGroup:

    global product_collection
    record = {}
    low_price = ''
    high_price = ''
    control_number = ''
    search_for_description = ''
    found_description = ''
    found_price = ''
    source_url = ''
    found_search_url = ''

    def cleanup_prices(self, amount):
        amount = amount.split('.')[0]
        amount = amount.replace('$', "")
        amount = amount.replace(",", "")
        amount = amount.replace(" ", "")
        return int(amount)

    def create_record(self):

        self.record["Control #"] = self.control_number
        self.record["Searched For This Item"] = self.search_for_description
        self.record["Selected Item Description"] = self.found_description
        self.record["Price"] = float(self.found_price)
        self.record["Low"] = float(self.low_price)
        self.record["High"] = float(self.high_price)
        self.record["URL Used For Search"] = self.source_url
        self.record["Selected Item URL"] = self.found_search_url

        self.add_record()

    def add_record(self):
        global product_collection
        stop_control_timer = time.perf_counter()

        print(str(self.control_number) + "\t\t" + self.search_for_description)
        #print(f"Time to process Control #:" + str(self.control_number) + {stop_control_timer - start_control_timer:0.4f} seconds")

        product_collection = product_collection.append(self.record, ignore_index=True)

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
        product_collection.to_csv('C:/Users/JRytt/Documents/Brad V/Temp_CSV.csv', index=False, sep='~')

    def process_control_group(self, dataframe):
        """
        Expecting a dataframe collection of records
        :param dataframe:
        :return:
        """

        df = dataframe
        if df.empty:
            pass
        else:
            try:
                df.sort_values(by=['Price'], inplace=True)  # sort by 'amount' column low to high
                df = df.iloc[1:, :]
                df = df.iloc[:-1, :]
            except TypeError:
                e = sys.exc_info()[0]
                print("<p>Error: %s</p>" % e)

            average_price = df["Price"].mean()
            self.low_price = str(df["Price"].min())
            self.high_price = str(df["Price"].max())

            # noinspection PyCompatibility
            for index, row in df.T.iteritems():

                less_than_mean = 0
                greater_than_mean = 0

                if row["Price"] < average_price:
                    less_than_mean = row["Price"]

                    self.control_number = row["Control #"]
                    self.search_for_description = row["Searched For This Item"]
                    self.found_description = row["Selected Item Description"]
                    self.found_price = str(row["Price"])
                    self.source_url = row["URL Used For Search"]
                    self.found_search_url = row["Selected Item URL"]
                    continue
                elif row["Price"] > average_price:
                    greater_than_mean = row["Price"]

                    # df is ordered lowest to highest when we find a higher (or equal)
                    # Price we have the data we need

                if average_price - less_than_mean > greater_than_mean - average_price:
                    self.control_number = row["Control #"]
                    self.search_for_description = row["Searched For This Item"]
                    self.found_description = row["Selected Item Description"]
                    self.found_price = str(row["Price"])
                    self.source_url = row["URL Used For Search"]
                    self.found_search_url = row["Selected Item URL"]

                self.create_record()
                break

# --------------------------------------  End of Class ProcessControlGroup  -----------------------------------------


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
    else:   # only 1 duplicate record found in the dataset
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


def read_file(input_file_name):
    """
    Read data from an excel spreadsheet and initiate the web search for those items
    :param input_file_name:
    """
    global start_control_timer
    global total_records_from_file
    last_control_number = ''
    pcg = ProcessControlGroup()

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

    found_dataset = FoundDataSet()
    for index, row in source_data_file.iterrows():
        start_control_timer = time.perf_counter()   # Start Timer for each Control #

        total_records_from_file += 1
        #print("Control #: " + str(row[0]) + "\t" + "Description: " + row[1])

        if duplicate_check([pcg, row[0], row[1]]):
            continue  # if duplicate search found, use that data and then continue -- no need to look up again

        url_to_search = "https://www.google.com/search?tbm=shop&q=" + row[1]
        url_to_search = make_tiny(url_to_search)

        if first_pass:
            last_control_number = row[0]
            first_pass = False

        if last_control_number == row[0]:
            found_dataset.process_current_control_group()   # When control # changes process that collection
        else:
            pcg.process_control_group(found_dataset.get_df_collection())
            found_dataset = FoundDataSet()

        found_dataset.control_number = row[0]
        found_dataset.search_for_description = row[1]
        found_dataset.source_url = url_to_search
        last_control_number = found_dataset.control_number


        item_search(found_dataset)            # Search the Web for this data

    pcg.process_control_group(found_dataset.get_df_collection())  # Used to write the last record to the file



def write_to_file(data_to_write):
    """
    Write to a given file
    :param data_to_write:
    """
    pass


def make_tiny(url):
    """
    Converts url to a 'tiny url' -- see http://tinyurl.com for additional details

    :param url:
    :return:

    try:
        request_url = ('http://tinyurl.com/api-create.php?' + urlencode({'url':url}))
        with contextlib.closing(urlopen(request_url)) as response:
            return response.read().decode('utf-8 ')
    except Exception:
        print('Warning -- Error attempting to make a tiny URL ')
        e = sys.exc_info()[0]
        print("<p>Error: %s</p>" % e)
    """

    return url #Returns the original URL if we can't convert to a tiny URL


def check_recaptcha(html_soup):
    """
    Examines the web page to see if there is a Recaptcha web page checking for robots  --I am not a robot

    :param html_soup:
    :return: True or False
    """
    if len(html_soup.find_all('div', {'id': "recaptcha"})) > 0:
        return True
    else:
        return False


def check_no_results(html_soup):
    if len(html_soup.find_all('div', {'class': "card-section"})) > 0:
        return True
    else:
        return False


def clean_the_money(amount):
    try:
        split_list = amount.split('.')
        amount = split_list[0]
        split_list = amount.split(' ')
        amount = split_list[0]          # Trim off cents
        split_list = amount.split('$')
        amount = split_list[1]          # trim off the $
        amount = amount.replace(",", "")
        amount = int(amount)
    except Exception:
        print('Warning -- Error cleaning the money: ' + amount)
        e = sys.exc_info()[0]
        print("<p>Error: %s</p>" % e)
        return 0
    return amount


def make_a_link(url, url_description=""):
    return url


def item_search(dataset):
    """ Perform a Google search for the URL provided"""

    global zero_product_count
    global record

    browser.get(dataset.source_url)

    html = browser.page_source
    html_soup = BeautifulSoup(html, 'html.parser')

    if check_recaptcha(html_soup):

        for x in range(10):
            for beep in range(10):
                winsound.Beep(440, 500)


            print("Recaptcha is blocking us.  Let's wait a few seconds and try again")
            time.sleep(30)
            print("Waking up... let's see if recaptcha is still blocking us")
            html = browser.page_source
            html_soup = BeautifulSoup(html, 'html.parser')
            if check_recaptcha(html_soup):
                continue
            else:
                break

    if check_no_results(html_soup):
        # TODO
        pass

    product_counter = 0

    for search_list in html_soup.find_all('div', {'class': "KZmu8e"}, limit=5):
        try:

            for each_product in search_list.find_all('div', {'class': "sh-np__product-title translate-content"}):
                product_counter += 1

            for product_link in search_list.find_all('a', href=True):
                # print("Product URL:https://www.google.com" + product_link['href'].lstrip())
                # Add the URL of the product found to the record set
                tiny_url = make_tiny("https://www.google.com" + product_link['href'].lstrip())
                dataset.found_search_url = make_a_link(tiny_url)

            description = each_product.get_text()

            # Add the description of the product found to the record set
            # dataset.found_description = make_a_link(tiny_url, description)
            dataset.found_description = description

            for product_price in search_list.find_all('span', {'class', "T14wmb"}):
                # Add the price of the product found to the record set
                dataset.found_price = clean_the_money(product_price.get_text())
                break

        except AttributeError:
            print("AttributeError Exception... check Control Number: " + dataset.control_number + \
                  " Description:" + dataset.search_for_description)
            product_counter = 0

        except Exception:
            e = sys.exc_info()[0]
            print("<p>Error: %s</p>" % e)
            product_counter = 0
            break

        if product_counter != 0:
            dataset.create_record()

    # -----------------2nd Attempt----------------------------------------------------------------------
    for search_list in html_soup.find_all('div', {'class': "sh-dlr__list-result"}, limit=5):
        product_counter += 1
        try:
            description = ''

            for product_link in search_list.find_all('a', href=True):
                # Add the URL of the product found to the record set
                tiny_url = make_tiny("https://www.google.com" + product_link['href'].lstrip())
                dataset.found_search_url = make_a_link(tiny_url)
                # dataset.found_search_url = "https://www.google.com" + product_link['href']
                if len(product_link['href']) > 0:
                    break

            for product_description in search_list.find_all('h3', {'class': "OzIAJc"}):
                description += product_description.get_text() + " "
                # Add the description of the product found to the record set
                # dataset.found_description = make_a_link(tiny_url, description)
                dataset.found_description = description

            for product_price in search_list.find_all('span', {'class', "QIrs8"}):
                dataset.found_price = clean_the_money(product_price.get_text())
                break
        except AttributeError:
            print("AttributeError Exception... check Control Number: " + dataset.control_number +
                  " Description:" + dataset.search_for_description)

            product_counter = 0
            break
        except Exception:
            e = sys.exc_info()[0]
            print("<p>Error: %s</p>" % e)
            product_counter = 0
            break

        if product_counter != 0:
            dataset.create_record()

    # --------------------------------------------------------------------------------------------------
    # -----------------3rd Attempt----------------------------------------------------------------------
    for search_list in html_soup.find_all('div', {'class': "sh-dgr__gr-auto sh-dgr__grid-result"}, limit=5):
        product_counter += 1
        try:
            description = ''

            for product_link in search_list.find_all('a', href=True):
                # Add the URL of the product found to the record set
                tiny_url = make_tiny("https://www.google.com" + product_link['href'].lstrip())
                dataset.found_search_url = make_a_link(tiny_url)
                if len(product_link['href']) > 0:
                    break

            for product_description in search_list.find_all('h4', {'class': "A2sOrd"}):
                description += product_description.get_text() + " "
                # Add the description of the product found to the record set
                dataset.found_description = make_a_link(tiny_url, description)
                dataset.found_description = description

            for product_price in search_list.find_all('span', {'class', "QIrs8"}):
                # Add the price of the product found to the record set
                dataset.found_price = clean_the_money(product_price.get_text())
                break
        except AttributeError:
            print("AttributeError Exception... check Control Number: " + dataset.control_number +
                  " Description:" + dataset.search_for_description)
            product_counter = 0
            break
        except Exception:
            e = sys.exc_info()[0]
            print("<p>Error: %s</p>" % e)
            product_counter = 0
            break

        if product_counter != 0:
            dataset.create_record()  # write the record

    # ---------------------------------------------------------------------------------------


if __name__ == '__main__':
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

