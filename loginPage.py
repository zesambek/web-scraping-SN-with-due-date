# Import the necessary libraries

# Selenium is a web automation framework that allows you to interact with web browsers.
from selenium import webdriver

# getpass provides a way to get the user's password without displaying it on the screen.
import getpass

# Keys provides a way to simulate keyboard input.
from selenium.webdriver.common.keys import Keys

# By provides a way to find elements on a web page by their id, name, class, etc.
from selenium.webdriver.common.by import By

# NoSuchElementException provides a way to handle exceptions that occur when an element cannot be found on a web page.
from selenium.common.exceptions import NoSuchElementException

# WebDriverWait provides a way to wait for an element to appear on a web page before continuing execution.
from selenium.webdriver.support.ui import WebDriverWait

# EC provides a set of conditions that can be used to wait for an element to appear on a web page.
from selenium.webdriver.support import expected_conditions as EC

import tkinter as tk
from tkinter import ttk


# Define a function to get the user credentials
def get_user_credentials():
    """Gets the username and password from the user."""

    # Get the username from the user.
    username = input("Enter your username: ")

    # Get the password from the user.
    password = getpass.getpass(prompt ="Enter your password: ")

    # Return the username and password.
    return username, password


# Define a function to log in to Ethiopian Airlines
def login_to_ethiopian_airlines(username, password, username_entry, password_entry, error_text, login_button):
    """Logs in to Ethiopian Airlines."""

    # Create an instance of the browser using the ChromeDriver.
    browser = webdriver.Chrome()
    # Navigate to the website's login page.
    browser.get("http://etmxi.ethiopianairlines.com/maintenix/common/security/login.jsp")

    # Wait for the login page to load completely.
    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, "j_username")))
    # Enter the username and password.
    browser.find_element(By.ID,"j_username").send_keys(username)
    browser.find_element(By.ID,"j_password").send_keys(password)
    # Enter the username and password.
    username_entry.delete(0, tk.END)
    username_entry.insert(0, username)
    password_entry.delete(0, tk.END)
    password_entry.insert(0, password)

    # Click the login button.
    browser.find_element(By.ID,"idButtonLogin").click()

    # Wait for the home page to load completely.
    WebDriverWait(browser, 10).until(lambda driver: driver.execute_script("return document.readyState === 'complete'"))

    if "Incorrect username and/or password" not in browser.page_source:
        print("Login successful!")
        error_text.config(state=tk.NORMAL)
        error_text.delete("1.0", tk.END)
        error_text.insert(tk.END, "Login successful!")
        error_text.config(state=tk.DISABLED)
        return browser
    else:
        # Check if the page source contains the error message.
        print("Incorrect username or password. Please try again.")
        error_text.config(state=tk.NORMAL)
        error_text.delete("1.0", tk.END)
        error_text.insert(tk.END, "Incorrect username or password. Please try again.")
        error_text.config(state=tk.DISABLED)
         # Enable the login button after a delay to ensure that it visually releases
        login_button.after(100, login_button.config, {"state": tk.NORMAL})
        # Return None to indicate an unsuccessful login.
        return None

