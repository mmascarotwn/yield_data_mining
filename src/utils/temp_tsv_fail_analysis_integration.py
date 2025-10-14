## test structure
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pyvirtualdisplay import Display

from tkinter import Tk
from tkinter.filedialog import askopenfilename

import pandas as pd


## Inputs declaration (to be changed if needed by the user)
link_web_page = 'http://mamweb.tatw.micron.com/MAMWeb/bin/MAMWeb.pl?APP=ASMTB&ACTION=REPORT&REPORTID=Status&XSLT=CLIENT&ID=368K-05:792368K&MATYPE=80&FORMAT=HTML&ORIGIN=/MAMWeb/site/ASMTB'

hbm3e_12h_id_code = "//span[text()='HBM3E (12H)']"
hbm3e_8h_id_code = "//span[text()='HBM3E (8H)']"
hbm4_12h_id_code = "//span[text()='HBM4 (12H)']"

dpm_lot_type_code = "//div[@role='option' and contains(., '6D_APTD_DPM')]"
qlc_lot_type_code = "//div[@role='option' and contains(., '6D_APTD_QLC')]"
rlc_lot_type_code = "//div[@role='option' and contains(., '6D_APTD_RLC')]"

c2pp_data_step_code = "//button[contains(text(), 'C2PP')]" # ATYMS gives all data avaialbility in raw data mode. No need to select others.

data_date_code = 'path[aria-label*="Month 202510, C2PP"]' # Change this date to be automatically detected later on


## Driver path and windows timing
web_driver = r"C:\Users\mmascaro\Documents\VisualStudio Projects\venv_test_project\chromedriver.exe"
timeout = 10
sleep = 0

## Setup to minimise the web page on screen (WIP)
chrome_options = Options()
chrome_options.add_argument("--headless=new") # Run in headless mode
chrome_options.add_argument("--disable-gpu") # Recommended for headless
chrome_options.add_argument("--no-sandbox") # Often required for headless mode, especially in Linux environments
chrome_options.add_argument("--disable-dev-shm-usage") # Often required for headless mode, especially in Linux environments


## Function to search for the web driver on the local PC
def chrome_web_driver(web_driver_path = None):
    """
    Text here
    """
    # Step 1: Text here
    s = Service(web_driver)
    driver = webdriver.Chrome(service = s)  # Optional argument, if not specified will search path.

    if not driver:
        print("Error: No driver specified/found.")
        return

    return driver

# Function to open the excel 'Open file selection' dialog to choose Excel file
def open_excel_dialogue(input = None):
    Tk().withdraw()  # Hide the root window
    excel_file = askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])

    return excel_file

# Function to read the first row and extract 'wafer_id_in_mam'
def load_wafer_id_excel(excel_file = None):
    df = pd.read_excel(excel_file, sheet_name='hbm_test_yield', engine='openpyxl')
    wafer_id = df.loc[0, 'wafer_id_in_mam']
    # print(wafer_id)

    if not excel_file:
        print("Error: No excel_file specified/found.")
        return

    return wafer_id

## Function to generate a dynamic URL
def generate_dynamic_url(wafer_id):
    """
    Replace the static lot ID in the URL with a dynamic input.

    Parameters:
    wafer_id (str): The lot ID to insert into the URL.

    Returns:
    str: The modified URL with the new lot ID.
    """
    base_url = f"http://mamweb.tatw.micron.com/MAMWeb/bin/MAMWeb.pl?APP=ASMTB&ACTION=REPORT&REPORTID=Status&XSLT=CLIENT&ID={wafer_id}&MATYPE=80&FORMAT=HTML&ORIGIN=/MAMWeb/site/ASMTB"
    print(wafer_id)
    print(base_url)
    
    if not wafer_id:
        print("Error: No lot_id specified/found.")
        return

    return base_url

## Function to open web page
def open_wfr_status_report_website(driver = None, link_web_page = None):
    """
    Text here
    """
    # Step 1: Text here
    driver.get(link_web_page)
    time.sleep(45) # Let the user actually see something!

    if not link_web_page :
        print("Error: No web page link found/provided.")
        return

    return 

## Function to hit enter onto the search bar (unused so far and untested)
def press_enter_in_search_bar(driver = None, timeout = None):
    """
    Locate the search bar element and simulate pressing the ENTER key.

    Parameters:
    driver (webdriver): The Selenium WebDriver instance.
    timeout (int): Maximum time to wait for the element to be present.
    """

    search_input = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, "//input[@type='text']")))
    search_input.send_keys(Keys.ENTER)
    print("ENTER key pressed in the search bar.")

    return

## Function to read the Fab Lot number
def extract_fab_lot_number(driver = None, timeout = None):
    """
    Extract the FAB LOT NUMBER from the webpage.

    Parameters:
    driver (webdriver): The Selenium WebDriver instance.
    timeout (int): Maximum time to wait for the element to be present.

    Returns:
    str: The extracted FAB LOT NUMBER, or None if not found.
    """
    try:
        # Wait for the element containing 'FAB LOT NUMBER'
        fab_lot_element = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, "//td[contains(text(), 'FAB LOT NUMBER')]/following-sibling::td"))
        )
        full_text = fab_lot_element.text.strip()
        fab_lot_number = full_text.split('.')[0]  # Get only the part before the first '.'

        print(f"FAB LOT NUMBER found: {fab_lot_number}")
        return fab_lot_number
    except Exception as e:
        print(f"Error extracting FAB LOT NUMBER: {e}")
        return None

## Function to navigate MAMASMTB(Assembly TAIWAN) and hit the search bar by entering a LOT#
def submit_lot_status_report(driver = None, timeout = None, lot_id = None):
    """
    Go to the Lot Status Report page, enter the Lot ID, and click the GO button.

    Parameters:
    driver (webdriver): Selenium WebDriver instance.
    lot_id (str): The Lot ID to submit.
    timeout (int): Timeout for waiting on elements.
    """
    try:
        # Step 1: Navigate to the Lot Status Report page
        driver.get("http://mamweb.tatw.micron.com/MAMWeb/site/ASMTB/")

        # Step 2: Wait for the input box and enter the Lot ID
        input_box = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.NAME, "ID"))
        )
        input_box.clear()
        input_box.send_keys(lot_id)

        # Step 3: Click the GO button
        go_button = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.NAME, "btnGo"))
        )
        go_button.click()

        print(f"Submitted Lot ID '{lot_id}' successfully.")
    except Exception as e:
        print(f"Error submitting Lot Status Report: {e}")
    
    return

## Function to open the selectable window for Product Group of interest
def select_product_window(driver = None, timeout = None):
    """
    Look for the "Prodcut Group" input search box on the webpage, click on it to expand drop down menu
    """
    wait = WebDriverWait(driver, timeout)
    product_group_span = wait.until(EC.element_to_be_clickable((By.XPATH, 
                    "//ng-select[contains(@class, 'ng-select')]//span[contains(text(), 'HBM3E (12H)')]/ancestor::ng-select"
                ))
            )
    product_group_span.click()
    time.sleep(2)

    if not timeout:
        print("Error: No timeout specified.")
        return

    return 

## Function to select the HBM product of interest within that window
def select_hbm_product(driver = None, timeout = None, product = None):
    """
    Click on a dropdown menu option, select "HBM4 (12H)"
    """
    option = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, product)))
    option.click()
    time.sleep(2)

    if not timeout:
        print("Error: No timeout specified.")
        return

    return 

## Function to open the expandable 'More Options' voice
def expand_menu_window(driver = None, timeout = None):
    """
    Locate the "More Options" menu
    """
    more_options_button = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'More Options')]")))
    driver.execute_script("arguments[0].scrollIntoView(true);", more_options_button)
    more_options_button.click()
    time.sleep(2)

    if not timeout:
        print("Error: No timeout specified.")
        return

    return 

## Function to select "Asm Eng Lot Descript"
def open_asm_eng_window(driver = None, timeout = None):
    """
    Wait for the section that contains the label "Asm Eng Lot Descript" and click on it
    """
    section = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//span[text()='Asm Eng Lot Descript']/ancestor::div[contains(@class, 'text-start')]")))
    dropdown_container = section.find_element(By.XPATH, ".//div[contains(@class, 'ng-select-container')]")
    dropdown_container.click()
    time.sleep(2)

    if not timeout:
        print("Error: No timeout specified.")
        return

    return 

## Function to unselect all voices from "Asm Eng Lot Descript"
def unselect_all(driver = None, timeout = None):
    """
    Click "Unselect All"
    """
    unselect_all_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[@title='Unselect All' and contains(text(), 'Unselect All')]")))
    unselect_all_button.click()
    time.sleep(2)

    if not timeout:
        print("Error: No timeout specified.")
        return

    return 

## Function to select QLC lots
def select_QLC(driver = None, timeout = None, lot_type = None):
    """
    Select the QLC lots
    """
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "ng-dropdown-panel-items")))
    options = driver.find_elements(By.XPATH, lot_type)
    for option in options:
        # Scroll into view and click each matching option
        driver.execute_script("arguments[0].scrollIntoView(true);", option)
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable(option)).click()
        time.sleep(1)
    time.sleep(2)

    if not timeout:
        print("Error: No timeout specified.")
        return

    return 

## Function to click "Filter" button and load the data on screen
def select_filter_button(driver = None, timeout = None):
    """
    Click "Filter" button data to load the data on ATYMS graph
    """
    filter_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn-warning') and contains(text(), 'Filter')]")))
    driver.execute_script("arguments[0].scrollIntoView(true);", filter_button)
    filter_button.click()
    time.sleep(2)

    if not timeout:
        print("Error: No timeout specified.")
        return

    return 

## Function to select C2PP data filtering on the plot
def select_C2PP_filter(driver = None, timeout = None, data_step = None):
    """
    Click the C2PP data only (actually ATYMS it gives all the other data as well)
    """
    c2pp_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, data_step)))
    driver.execute_script("arguments[0].scrollIntoView(true);", c2pp_button)
    c2pp_button.click()
    time.sleep(3)

    if not timeout:
        print("Error: No timeout specified.")
        return

    return 

## Function to select the last month of data available
def select_date_filter(driver = None, timeout = None, data_date = None):
    """
    Click the latest month data set available (fixed to '202510') # TO-DO: make it depending on current month
    TO-DO: 
        1) implement a method to aoutomatically select the latest data available by month on the ATYMS website;
    """
    data_point = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, data_date)))
    driver.execute_script("arguments[0].scrollIntoView(true);", data_point)
    data_point.click()
    time.sleep(3)

    if not timeout:
        print("Error: No timeout specified.")
        return

    return 

## Function to download the data (.xlsx)
def download_data(driver = None, timeout = None):
    """
    Click "Download raw data" button
    """
    download_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[@title='Download FID and Wafer level raw data.']")))
    driver.execute_script("arguments[0].scrollIntoView(true);", download_button)
    download_button.click()
    time.sleep(3)

    if not timeout:
        print("Error: No timeout specified.")
        return

    return 

## Function to quit the process afterwards
def quit_script(driver = None, timeout = None):
    """
    Wait and Quit script
    """
    time.sleep(10) # Let the user actually see something!
    driver.quit()

    if not timeout:
        print("Error: No timeout specified.")
        return

    return 


###
### Functions grouping
###
def lot_search(a=None):
    select_excel_file = open_excel_dialogue()
    read_wafer_id = load_wafer_id_excel(select_excel_file)
    generale_url = generate_dynamic_url(read_wafer_id)
    driver = chrome_web_driver(web_driver)
    dynamic_link_web_page = generate_dynamic_url(generale_url)
    wafer_MAM_web_page = open_wfr_status_report_website(driver, dynamic_link_web_page)
    read_fab_lot_number = extract_fab_lot_number(driver, timeout)
    search_lot_MAM = submit_lot_status_report(driver, timeout, read_fab_lot_number)
    # hit_search_bar = press_enter_in_search_bar(driver, timeout)
    quit = quit_script(driver, timeout)

def hbm4_dpm_data(a=None):
    # Step 1: Text here
    driver = chrome_web_driver(web_driver)
    atyms_web_page = open_atyms_website(driver, link_web_page)
    open_hbm_product_window = select_product_window(driver, timeout)
    select_hbm = select_hbm_product(driver, timeout, hbm4_12h_id_code) # option
    expand_window = expand_menu_window(driver, timeout)
    open_asm_eng = open_asm_eng_window(driver, timeout)
    unselect_asm_eng = unselect_all(driver, timeout)
    select_QLC_lots = select_QLC(driver, timeout, dpm_lot_type_code) # option
    display_data = select_filter_button(driver, timeout)
    select_C2PP_data = select_C2PP_filter(driver, timeout, c2pp_data_step_code) 
    select_date = select_date_filter(driver, timeout, data_date_code) # option
    click_download = download_data(driver, timeout)
    quit = quit_script(driver, timeout)

def hbm4_rlc_data(a=None):
    # Step 1: Text here
    driver = chrome_web_driver(web_driver)
    atyms_web_page = open_atyms_website(driver, link_web_page)
    open_hbm_product_window = select_product_window(driver, timeout)
    select_hbm = select_hbm_product(driver, timeout, hbm4_12h_id_code) # option
    expand_window = expand_menu_window(driver, timeout)
    open_asm_eng = open_asm_eng_window(driver, timeout)
    unselect_asm_eng = unselect_all(driver, timeout)
    select_QLC_lots = select_QLC(driver, timeout, rlc_lot_type_code) # option
    display_data = select_filter_button(driver, timeout)
    select_C2PP_data = select_C2PP_filter(driver, timeout, c2pp_data_step_code) 
    select_date = select_date_filter(driver, timeout, data_date_code) # option
    click_download = download_data(driver, timeout)
    quit = quit_script(driver, timeout)


###
### test the functions
###
lot_search()
# hbm4_dpm_data()
# hbm4_rlc_data()