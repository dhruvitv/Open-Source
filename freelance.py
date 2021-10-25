import requests
import datetime
from time import sleep
import pandas as pd
import numpy as np
import xlrd
import xlsxwriter
import os
from win32com.client import Dispatch
from bs4 import BeautifulSoup
import re
from selenium import webdriver

URL = "https://www.nseindia.com/option-chain?symbolCode=-10006&symbol=NIFTY&symbol=NIFTY&instrument=-&date=-&segmentLink=17&symbolCount=2&segmentLink=17"
USER_AGENT = {"User-Agent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.141 Safari/537.36'}
TIMEOUT = 10  # The amount of time the script should wait after requesting data from the server before giving up
WAIT_INTERVAL = 1  # The number of minutes to wait between each download

INTERESTED_PRICE = 0


def main():
    global INTERESTED_PRICE
    
    INTERESTED_PRICE = input('Please enter the strike price you are interested in (number only):\n')
    INTERESTED_PRICE = int(INTERESTED_PRICE.replace(',', ''))
    last_exec_time = datetime.datetime.now()
    close_workbook()

    driver = make_new_driver(0, 30)
    download_selenium(driver, INTERESTED_PRICE)
    # download(INTERESTED_PRICE)
    open_workbook()
    print(f'Next download will be executed at or around {(last_exec_time + datetime.timedelta(minutes=WAIT_INTERVAL)).strftime("%I:%M:%S")}')

    alerted = False
    while True:
        # Check if it's time to download again
        try:
            if datetime.datetime.now() - last_exec_time >= datetime.timedelta(minutes=WAIT_INTERVAL):
                close_workbook()
                last_exec_time = datetime.datetime.now()
                download_selenium(driver, INTERESTED_PRICE)
                print(f'Next download will be executed at or around {(last_exec_time + datetime.timedelta(minutes=WAIT_INTERVAL)).strftime("%I:%M:%S")}')
                alerted = False
                open_workbook()
                print('=============================Done===================================')
            else:  # Otherwise, sleep for a bit
                sleep(5)
        except Exception as e:
            print(e)
            print("\nRe-running it again because some error occured!!!!")


def make_new_driver(headless_flag=0, page_load_timeout=30):
    options = webdriver.ChromeOptions()
    if headless_flag == 1:
        #options.add_argument('headless')  # Allows Chrome to start in headless mode
        options.add_argument('--window-size=1920x1080')  # Since "maximized" in headless mode means nothing
    options.add_argument('-start-maximized')
    
    #options.add_argument('headless')
    options.add_argument('--log-level=3')
    driver = webdriver.Chrome('chromedriver.exe', options=options)
    driver.set_page_load_timeout(page_load_timeout)
    return driver

def close_workbook():
    try:
        print('Closing workbook')
        xl = Dispatch('Excel.Application')
        wb = xl.Workbooks.Open(os.path.join(os.getcwd(), 'scrape_output.xlsx'))

        try:
            wb.Save()
            wb.close()
        except PermissionError:
            print('Permission error received. Waiting and reattempting')
            sleep(1)
            wb.close()

    except:
        return


def open_workbook():
    print('Opening workbook')
    xl = Dispatch('Excel.Application')
    xl.Workbooks.Open(os.path.join(os.getcwd(), 'scrape_output.xlsx'))
    xl.Visible = True

# Responsible for returning a blank row except for having the HH:MM time right in the middle
def get_timing_row(row_length, current_price):
	row = ['']*int(row_length/2) # int will always round down if it's a decimal
    row.append(datetime.datetime.today().strftime('%d-%m-%Y'))
	row.append(datetime.datetime.now().strftime('%I:%M'))
	row.append(str(current_price))
	row.extend(['']*(row_length - int(row_length/2)-2))
	return row

def func(x):
        try:
            return float(x)
        except ValueError:
            return x


def download_selenium(driver, interested_price):
    print(f'Downloading new data at {datetime.datetime.now().strftime("%I:%M:%S")}')
    
    sleep(3)
    driver.delete_all_cookies()
    driver.get(URL)
    driver.minimize_window()
    sleep(3)
    driver.delete_all_cookies()
    driver.get(URL)
    sleep(3)  # Wait for the underlying text index dollar value to exist
    underlying_index_text = driver.find_element_by_css_selector('span#equity_underlyingVal').text
    print(underlying_index_text)
    underlying_index_text = underlying_index_text.replace(',', '')
    print(underlying_index_text)
    current_price = re.findall(r'\d+\.\d{2}', underlying_index_text)[0]
    print(current_price)
    sleep(1)

    # Use Pandas to read charts in from the page
    df = pd.read_html(driver.page_source)[0]
    # Drop first and last columns because they are the columns with little chart icons which will alwyas be NaN
    df.drop(df.columns[-1], axis=1, inplace=True)
    df.drop(df.columns[0], axis=1, inplace=True)
    #print(df.to_string())

    # Iterate over columns and renamed Unnamed* to Unnamed
    for i, col in enumerate(df.columns.levels):
        columns = np.where(col.str.contains('Unnamed'), 'Unnamed', col)
        #print(columns)
        df.columns.set_levels(columns, level=i, inplace=True)
        break

    #print(df.to_string())


    # Get the index of the row we care about
    index = df.index[df['Unnamed']['Strike Price'] == interested_price].values[0]
    interested_rows = df[index:index + 11]

    # Dump all the data into a 2d list
    new_rows = []
    for index, row in interested_rows.iterrows():
        new_rows.append(
            [func(item) for item in row.values])  # iterrows returns arrays, which need to be iterated through as well
    
    new_rows.insert(0, get_timing_row(len(new_rows[0]), current_price))
    #print(new_rows)

    # Read the existing Excel data
    try:
        worksheet = pd.read_excel('scrape_output.xlsx' , header = None  , engine='openpyxl')
        
        worksheet = worksheet.append(new_rows)
        worksheet = worksheet.apply(pd.to_numeric, errors='ignore')
        worksheet.to_excel('scrape_output.xlsx' , index = False , header = None)
        #wb = xlrd.open_workbook('scrape_output.xlsx')
        #sheet = wb.sheet_by_name('Sheet1')
        
        # Get the data
        #previous_rows = []
        
        
        #num_cols = sheet.ncols
        
        
        # Iterate over every row
        #for row in range(0, num_rows):
        #   row_data = []
        #    for col in range(0, num_cols):
        #        # Iterate over every column
        #        row_data.append(sheet.cell(row, col).value)
        #    previous_rows.append(row_data)

        # This is where we would add in the extra data requested 1/23

        # Build a new list of the rows and extend the new rows onto it so we are ready to write
        #rows_to_write = previous_rows
        #rows_to_write.extend(new_rows)
    except FileNotFoundError:  # If we get this error, this is our first time running and we can consider no past data
        print('Output file does not yet exist')
        new_rows.insert(0, [''] * len(new_rows[0]))
        initial_excel = pd.DataFrame(new_rows)
        initial_excel = initial_excel.apply(pd.to_numeric, errors='ignore' , downcast = "float")
        initial_excel.to_excel('scrape_output.xlsx' , index = False , header = None)
        # Add a blank row to the beginning per Govind's request
        #rows_to_write.insert(0, [''] * len(rows_to_write[0]))
    return
    
if __name__ == "__main__":
    main()
