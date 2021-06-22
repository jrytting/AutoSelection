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


class GoogleRecordSearch:
    def __init__(self, name=''):
        self.name = name

    # prepare the option for the chrome driver
    options = webdriver.ChromeOptions()
    options.add_argument('headless')

    # start chrome browser
    browser = webdriver.Chrome()  # options=options

    def check_recaptcha(self, html_soup):
        """
        Examines the web page to see if there is a Recaptcha web page checking for robots  --I am not a robot

        :param html_soup:
        :return: True or False
        """
        if len(html_soup.find_all('div', {'id': "recaptcha"})) > 0:
            raise(ReCaptchaError(message="reCaptcha Screen is blocking our search"))
        else:
            return False

    def check_no_results(self, html_soup):
        if len(html_soup.find_all('div', {'class': "card-section"})) > 0:
            raise(NoResultsReturnedError(message="No results were returned by Search"))
        else:
            return False

    def clean_the_money(self, amount):
        try:
            split_list = amount.split('.')
            amount = split_list[0]
            split_list = amount.split(' ')
            amount = split_list[0]  # Trim off cents
            split_list = amount.split('$')
            amount = split_list[1]  # trim off the $
            amount = amount.replace(",", "")
            amount = int(amount)
        except Exception:
            print('Warning -- Error cleaning the money: ' + amount)
            e = sys.exc_info()[0]
            print("<p>Error: %s</p>" % e)
            exit(2)

        return amount

    def item_search(self, control_group):
        """ Perform a Google search for the URL provided"""

        self.browser.get(control_group.source_url)

        html = self.browser.page_source
        html_soup = BeautifulSoup(html, 'html.parser')

        if self.check_recaptcha(html_soup):

            for beep in range(10):
                winsound.Beep(440, 500)

            time.sleep(5)
            """
            time.sleep(30)
            print("Waking up... let's see if recaptcha is still blocking us")
            html = self.browser.page_source
            html_soup = BeautifulSoup(html, 'html.parser')
            if self.check_recaptcha(html_soup):
                continue
            else:
                break
            """

        if self.check_no_results(html_soup):
            pass

        product_counter = 0

        for search_list in html_soup.find_all('div', {'class': "KZmu8e"}, limit=5):
            try:

                for each_product in search_list.find_all('div', {'class': "sh-np__product-title translate-content"}):
                    product_counter += 1

                for product_link in search_list.find_all('a', href=True):
                    # Add the URL of the product found to the record set
                    control_group.found_search_url = "https://www.google.com" + product_link['href'].lstrip()

                description = each_product.get_text()

                # Add the description of the product found to the record set
                control_group.found_description = description

                for product_price in search_list.find_all('span', {'class', "T14wmb"}):
                    # Add the price of the product found to the record set
                    control_group.found_price = self.clean_the_money(product_price.get_text())
                    break

            except AttributeError:
                print("AttributeError Exception... check Control Number: " + control_group.control_number + \
                      " Description:" + control_group.search_for_description)
                product_counter = 0

            except Exception:
                e = sys.exc_info()[0]
                print("<p>Error: %s</p>" % e)
                product_counter = 0
                break

            if product_counter != 0:
                control_group.add_record_to_control_group()

        # -----------------2nd Attempt----------------------------------------------------------------------
        for search_list in html_soup.find_all('div', {'class': "sh-dlr__list-result"}, limit=5):
            product_counter += 1
            try:
                description = ''

                for product_link in search_list.find_all('a', href=True):
                    # Add the URL of the product found to the record set
                    control_group.found_search_url ="https://www.google.com" + product_link['href'].lstrip()

                    # control_group.found_search_url = "https://www.google.com" + product_link['href']
                    if len(product_link['href']) > 0:
                        break

                for product_description in search_list.find_all('h3', {'class': "OzIAJc"}):
                    description += product_description.get_text() + " "
                    # Add the description of the product found to the record set
                    # control_group.found_description = make_a_link(tiny_url, description)
                    control_group.found_description = description

                for product_price in search_list.find_all('span', {'class', "QIrs8"}):
                    control_group.found_price = self.clean_the_money(product_price.get_text())
                    break
            except AttributeError:
                print("AttributeError Exception... check Control Number: " + control_group.control_number +
                      " Description:" + control_group.search_for_description)

                product_counter = 0
                break
            except Exception:
                e = sys.exc_info()[0]
                print("<p>Error: %s</p>" % e)
                product_counter = 0
                break

            if product_counter != 0:
                control_group.add_record_to_control_group()


        # --------------------------------------------------------------------------------------------------
        # -----------------3rd Attempt----------------------------------------------------------------------
        for search_list in html_soup.find_all('div', {'class': "sh-dgr__gr-auto sh-dgr__grid-result"}, limit=5):
            product_counter += 1
            try:
                description = ''

                for product_link in search_list.find_all('a', href=True):
                    # Add the URL of the product found to the record set
                    control_group.found_search_url ="https://www.google.com" + product_link['href'].lstrip()
                    if len(product_link['href']) > 0:
                        break

                for product_description in search_list.find_all('h4', {'class': "A2sOrd"}):
                    description += product_description.get_text() + " "
                    # Add the description of the product found to the record set
                    control_group.found_description = description

                for product_price in search_list.find_all('span', {'class', "QIrs8"}):
                    # Add the price of the product found to the record set
                    control_group.found_price = self.clean_the_money(product_price.get_text())
                    break
            except AttributeError:
                print("AttributeError Exception... check Control Number: " + control_group.control_number +
                      " Description:" + control_group.search_for_description)
                product_counter = 0
                break
            except Exception:
                e = sys.exc_info()[0]
                print("<p>Error: %s</p>" % e)
                product_counter = 0
                break

            if product_counter != 0:
                control_group.add_record_to_control_group()

        # ---------------------------------------------------------------------------------------