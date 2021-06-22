import pandas as pd
import time
import sys


class DuplicateRecordError(Exception):
    """Exception raised duplicate record found in parent collection

    Attributes:
        salary -- input salary which caused the error
        message -- explanation of the error
    """
    def __init__(self, message="Found a matching record in the parent control_group"):
        self.message = message
        super().__init__(self.message)
# --------------------------  End of Exception Definitions --------------------------------


class ControlGroupMasterDataSet:
    def __init__(self):
        self.project_collection = pd.DataFrame()

    def get_collections(self):
        return self.project_collection

    def commit_control_group(self, control_group):
        try:
            self.project_collection.append(control_group, ignore_index=True)
            return True
        except Exception:
            print("Unable to add control group to parent control_group")
            e = sys.exc_info()[0]
            print("<p>Error: %s</p>" % e)
            exit(2)
            return False



"""
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

    def duplicate_check(control_number, source_description):
        global product_collection
        if product_collection.empty:
            return False

        result = product_collection.loc[product_collection['Searched For This Item'] == source_description]

        if result.empty:  # No Duplicate Record Found
            return False

        if len(result.index) > 1:  # when multiple records are returned, use data from last row in selection
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
        else:  # only 1 duplicate record found in the control_group
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
"""
# --------------------------------------  End of Class ControlGroupDataSet  -----------------------------------------