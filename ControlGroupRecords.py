from datetime import datetime
import winsound
import pandas as pd
import numpy as np
from sklearn.datasets import load_iris
import time
import sys
import tabulate
import ControlGroupCollection
from GoogleSearch import NoResultsReturnedError
from GoogleSearch import ReCaptchaError


class ControlGroupDataSet:
    def __init__(self):
        self.control_collection = pd.DataFrame()

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
        if self.low_price != '':
            self.record["Low"] = self.low_price                 # Only create these on the master collection
            self.record["High"] = self.high_price               # Only create these on the master collection
        self.record["URL Used For Search"] = self.source_url
        self.record["Selected Item URL"] = self.found_search_url

    def add_record_to_control_group(self):
        try:
            self.create_record()
            self.control_collection = self.control_collection.append(self.record, ignore_index=True)
            return True
        except Exception:
            print("Unable to add record to control group")
            e = sys.exc_info()[0]
            print("<p>Error: %s</p>" % e)
            raise(NoResultsReturnedError(message="Error adding record to Control Group"))

    def add_selected_control_group_record_to_collection(self, master_dataset):
        # Build the record and commit to master data set
        try:
            self.create_record()
            master_dataset = master_dataset.append(self.record, ignore_index=True)
            print("Added Record: ", self.control_number, self.search_for_description)
            return master_dataset
        except Exception:
            print("Unable to add record to control group")
            e = sys.exc_info()[0]
            print("<p>Error: %s</p>" % e)
            raise(NoResultsReturnedError(message="Control Group to Master Data Collection"))

    def get_df_collection(self):
        return self.control_collection

    def select_record_from_control_group(self, master_data):

        # Rule: Drop Highest and Lowest Amounts and then determine average price with remainder of elements
        # :return:

        if self.control_collection.empty:
            return master_data
            # raise(NoResultsReturnedError(message="Control Group Collection was empty -- This should not happen"))
        else:
            try:
                self.control_collection.sort_values(by=['Price'], inplace=True)  # sort by 'amount' column low to high
                if self.control_collection["Control #"].count() > 2:
                    self.control_collection = self.control_collection.iloc[1:, :]           # Throw out low record
                    self.control_collection = self.control_collection.iloc[:-1, :]          # throw out high record
                self.average_price = self.control_collection["Price"].mean()
                self.low_price = self.control_collection["Price"].min()
                self.high_price = self.control_collection["Price"].max()
            except Exception:
                print(self.control_collection.to_markdown())
                print("Unable to sort/get min/max value from the record from the Control Group")
                e = sys.exc_info()[0]
                print("<p>Error: %s</p>" % e)
                raise(NoResultsReturnedError(message="Control Group to Master Data Collection"))

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
                try:
                    # Build Selected/Target Recordset and add to Master Collection
                    return self.add_selected_control_group_record_to_collection(master_data)
                except Exception:
                    print("Unable to add record to Master Collection")
                    e = sys.exc_info()[0]
                    print("<p>Error: %s</p>" % e)
                    raise(NoResultsReturnedError(message="Control Group to Master Data Collection"))
# --------------------------------------  End of Class FoundDataSet  ----------------------------------------------
