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

import ControlGroupRecords


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


class BingSearch:
    def __init__(self, name=''):
        self.name = name

    # prepare the option for the chrome driver
    options = webdriver.ChromeOptions()
    options.add_argument('headless')

    # start chrome browser
    browser = webdriver.Chrome()  # options=options

    def check_recaptcha(self, html_soup):
        if len(html_soup.find_all('div', {'id': "recaptcha"})) > 0:
            # raise(ReCaptchaError(message="reCaptcha Screen is blocking our search"))
            return True
        else:
            return False

    def check_no_results(self, html_soup):
        if len(html_soup.find_all('div', {'class': "card-section"})) > 0:
            raise(NoResultsReturnedError(message="No results were returned by Search"))
        else:
            return False

    def clean_the_money(self, amount):
        try:
            amount = amount.lower().translate({ord(i): None for i in 'abcdefghijklmnopqrstuvwxyz$*!@#%()+-, '})
            amount = amount.split(".")[0]
            amount = int(amount)
        except Exception as e:
            print('Warning -- Error cleaning the money: ' + amount)
            e = sys.exc_info()[0]
            print("<p>Error: %s</p>" % e)
            raise(NoResultsReturnedError(message="No results were returned by Search"))

        return amount

    def item_search(self, control_group):
        """ Perform a Google search for the URL provided"""

        self.browser.get(control_group.source_url)

        html = self.browser.page_source
        html_soup = BeautifulSoup(html, 'html.parser')

        if self.check_recaptcha(html_soup):
            for x in range(10):
                for beep in range(3):
                    winsound.Beep(440, 500)

                time.sleep(30)
                print("Waking up... let's see if recaptcha is still blocking us")
                html = self.browser.page_source
                html_soup = BeautifulSoup(html, 'html.parser')
                if self.check_recaptcha(html_soup):
                    continue
                else:
                    break

        if not self.check_no_results(html_soup):
            pass

        product_counter = 0

        for search_list in html_soup.find_all('div',
                                              {'class': "br-gOffCard br-narrowOffCard br-offHover br-vistrackitm"},
                                              limit=10):
            try:
                for product_link in search_list.find_all('a', href=True):
                    # Add the URL of the product found to the record set
                    control_group.found_search_url = product_link['href'].lstrip()
                    break

                for each_product in product_link.find_all('div', {'class': "br-offCnt"}, limit=5):
                    product_counter += 1

                for each_description in product_link.find_all('div', class_="br-offTtl b_primtxt"):
                    description = each_description.get_text()

                    # Add the description of the product found to the record set
                    control_group.found_description = description
                    break

                for product_price in each_product.find_all('div', {'class', "br-offPrice b_primtxt"}):
                    # Add the price of the product found to the record set
                    control_group.found_price = self.clean_the_money(product_price.get_text())
                    break

                if product_counter != 0:
                    control_group.add_record_to_control_group()

            except AttributeError:
                print("AttributeError Exception... check Control Number: " + control_group.control_number + \
                      " Description:" + control_group.search_for_description)
                product_counter = 0
                raise (NoResultsReturnedError(message="No results were returned by Search"))

            except Exception:
                e = sys.exc_info()[0]
                print("<p>Error: %s</p>" % e)
                product_counter = 0
                # break
                raise(NoResultsReturnedError(message="No results were returned by Search"))

        # -----------------2nd Attempt----------------------------------------------------------------------
        for search_list in html_soup.find_all('div', {'class': "br-freeads"}):
            try:

                for each_product in search_list.find_all('li', {'class': "br-item"}, limit=10):
                    product_counter += 1

                    for product_link in each_product.find_all('a', class_="ofr_lnk"):
                        # Add the URL of the product found to the record set
                        control_group.found_search_url = product_link['href'].lstrip()
                        break

                    for each_description in each_product.find_all('div', class_="br-pdItemName-onHover"):
                        description = each_description.get_text()
                        # Add the description of the product found to the record set
                        control_group.found_description = description
                        break

                    # Alternative Description Screen Location
                    for each_description in each_product.find_all('a', class_="br-voidlink br-titleClickWrap"):
                        description = each_description.get_text()
                        control_group.found_description = description

                    for product_price in each_product.find_all('div', {'class', "pd-price"}):
                        # Add the price of the product found to the record set
                        control_group.found_price = self.clean_the_money(product_price.get_text())
                        break

                    # Alternative Pricing Screen Location
                    for product_price in each_product.find_all('span', {'class', "br-focusPrice"}):
                        # Add the price of the product found to the record set
                        control_group.found_price = self.clean_the_money(product_price.get_text())
                        break

                    if product_counter != 0:
                        control_group.add_record_to_control_group()

            except AttributeError:
                print("AttributeError Exception... check Control Number: " + control_group.control_number + \
                      " Description:" + control_group.search_for_description)
                product_counter = 0
                raise (NoResultsReturnedError(message="No results were returned by Search"))

            except Exception:
                e = sys.exc_info()[0]
                print("<p>Error: %s</p>" % e)
                product_counter = 0
                # break
                raise(NoResultsReturnedError(message="No results were returned by Search"))

        #  Take this out when adding additional searches....
        if product_counter == 0:
            raise (NoResultsReturnedError(message="No results were returned by Search"))
        # ---------------------------------------------------------------------------------------
