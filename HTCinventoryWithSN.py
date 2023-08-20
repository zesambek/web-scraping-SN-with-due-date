# Import the necessary libraries

import sys


# Selenium is a web automation framework that allows you to interact with web browsers.
from selenium import webdriver

from openpyxl import load_workbook

# The `pandas` module is a powerful library for data analysis and manipulation.
import pandas as pd

# By provides a way to find elements on a web page by their id, name, class, etc.
from selenium.webdriver.common.by import By

#NoSuchElementException provides a way to handle exceptions that occur when an element cannot be found on a web page.
from selenium.common.exceptions import NoSuchElementException

# WebDriverWait provides a way to wait for an element to appear on a web page before continuing execution.
from selenium.webdriver.support.ui import WebDriverWait

# EC provides a set of conditions that can be used to wait for an element to appear on a web page.
from selenium.webdriver.support import expected_conditions as EC

# Import the regular expression module
import re

#provides the user password and name
from loginPage import get_user_credentials,login_to_ethiopian_airlines

#To select an option from the dropdown list using the Select class
from selenium.webdriver.support.ui import Select

# Define the search_serial_numbers function
def search_serial_numbers(driver, file_path,sheet_name):

    # Parse the first sheet in the file into a Pandas DataFrame
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    #for serial_number in serial_numbers:
    for index, row in df.iterrows():
        # Go to the inventory search page
        driver.get("http://etmxi.ethiopianairlines.com/maintenix/web/inventory/InventorySearchByType.jsp")

        # Wait for the element to be present on the page
        wait = WebDriverWait(driver, 100000)
        element = wait.until(EC.presence_of_element_located((By.ID, "idGBox1")))

        # Find the dropdown list by its ID
        dropdown = Select(driver.find_element(By.ID, "aSearchType"))

        # Select an option by its value
        dropdown.select_by_value("SERIAL")

        # Wait for the page to load
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.ID, "idBySerial_SerialNoLabel")))

        #find the input box to enter the serial number
        serial_number_element = browser.find_element(By.ID,"idBySerial_SerialNo").find_elements(By.TAG_NAME, "input")[0]

        # Enter the serial number
        serial_number_element.send_keys(row['Serial No / Batch No'])
        driver.find_element(By.ID, "idButtonSearch").click()

        # Wait for the page to load
        WebDriverWait(driver, 1000).until(EC.presence_of_element_located((By.ID, "idTableSearchResults")))

        #Find the table by ID
        table  = driver.find_element(By.ID, "idTableSearchResults")

        # Get the table data as a Pandas DataFrame
        SN_df = pd.read_html(table.get_attribute("outerHTML"),header=0)[0]

        # Select the rows where the OEM Part No column is equal to the specific part number
        SN_filtered_df = SN_df[(SN_df["OEM Part No"] == row["OEM Part No"]) & (SN_df["Installed On"]==row["Installed On"]) & (SN_df["Part Name"]==row["Part Name"])]
              
        # Check if the filtered DataFrame is empty
        if  SN_filtered_df.empty:

            df.at[index, 'Due Date'] = "no matching rows found"
        else: 

            # Get the content of the `Condition` column
            SN_condition_content = SN_filtered_df["Condition"].iloc[0]

            # Check the value of the `Condition` column
            if SN_condition_content == "INSRV":

                # Get the value of the `Serial No` column
                serial_no = SN_filtered_df.iloc[0]["Serial No / Batch No"]

                # Use the `index` attribute to get the row number of the first occurrence of the value
                row_number = SN_filtered_df.index[SN_filtered_df['Serial No / Batch No'] == serial_no].tolist()[0]

                # Increment the row number by 1
                row_number += 2
                xpath = "//*[@id='idTableSearchResults']/tbody/tr[position() = " + str(row_number) + "]/td[6]/a"

                # Find the anchor element by its link text
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))

                # Find the element using XPath
                anchor = driver.find_element(By.XPATH, xpath)

                # click the anchor element
                anchor.click()

                # Wait for the document to be loaded before continuing
                WebDriverWait(driver, 1000).until(EC.presence_of_element_located((By.ID, "Open_link")))
                driver.find_element(By.ID, "Open_link").click()

                # Wait for the document to be loaded before continuing
                WebDriverWait(driver, 1000).until(EC.presence_of_element_located((By.ID, "OpenTasks_link")))
                driver.find_element(By.ID, "OpenTasks_link").click()

                # Wait for the document to be loaded before continuing
                WebDriverWait(driver, 1000).until(EC.presence_of_element_located((By.ID, "idTableOpenTasks")))

                #Find the table by ID
                Due_date_table  = driver.find_element(By.ID, "idTableOpenTasks")
                # Get the table data as a Pandas DataFrame
                task_df = pd.read_html(Due_date_table.get_attribute("outerHTML"),header=0)[0]

                #task_filtered_df = task_df[(task_df["Due"].notnull()) & (task_df["Originator"]=="MPD")]
                task_filtered_df = task_df[task_df["Due"].notnull()]

                # Check if the filtered DataFrame is empty
                if not task_filtered_df.empty:
                    # Get the content of the `DUE` column
                    due_date_element = task_filtered_df["Due"].iloc[1]
                    print(due_date_element)
                    # Extract date from text using regular expression
                    date_regex = r"\d{2}-[A-Z]{3}-\d{4}"
                    due_date_match = re.search(date_regex, str(due_date_element))
                    if due_date_match:
                        print(due_date_match.group())
                        df.at[index, 'Due Date'] = due_date_match.group()
                    else:
                        print("No date found")
                        df.at[index, 'Due Date'] = "No date found"
                else:
                    print("No task.")
                    df.at[index, 'Due Date'] = "the part do not have  tasks"
            elif SN_condition_content == "REPREQ":
                df.at[index, 'Due Date'] = "OFF WING"
                print("the part needs to be replaced.")
            else:
                print("unknown condition.")
                df.at[index, 'Due Date'] = "UNKNOWN CONDITION"
    #Write the processed DataFrame back to the Excel file
    with pd.ExcelWriter(file_path, engine='openpyxl',if_sheet_exists='replace', mode='a') as writer:
       df.to_excel(writer, index=False, sheet_name=sheet_name)
       

def write_due_date(driver, file_path,selected_sheet_name=None):

    # Read the Excel file and return an object that contains all the sheets in the file
    excel_file = pd.ExcelFile(file_path)

    # Get the names of all the sheets in the file
    sheet_names = excel_file.sheet_names

    # If a specific sheet_name is provided, only process that sheet
    if selected_sheet_name is not None:
        sheet_names = [selected_sheet_name]
    print("sams")
    for sheet_name in sheet_names:
        # Perform the search on the DataFrame
        search_serial_numbers(driver,file_path, sheet_name)

    browser.close()





