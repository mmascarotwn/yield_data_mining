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


## Inputs declaration
web_driver = r"C:\Users\mmascaro\Documents\VisualStudio Projects\venv_test_project\chromedriver.exe"
# link_web_page = 'http://mamweb.tatw.micron.com/MAMWeb/site/ASMTB/'
link_web_page = 'https://o2-instance.si-ose-mfgt1-vip.sing.micron.com/ATYMS/#/?c.Category=Apps&c.Module=AT-Yield-Report&c.View=HBMYield&c.Tab=HBM-Test-Report'
# Lot_ID = '788479K.0BL'
Product_ID = 'HBM4 (12H)'


## Setup to minimise the web page on screen (WIP)
chrome_options = Options()
chrome_options.add_argument("--headless=new") # Run in headless mode
chrome_options.add_argument("--disable-gpu") # Recommended for headless
chrome_options.add_argument("--no-sandbox") # Often required for headless mode, especially in Linux environments
chrome_options.add_argument("--disable-dev-shm-usage") # Often required for headless mode, especially in Linux environments
# display = Display(visible=0, size=(800, 600))
# display.start()


## 1. Get web driver on local PC
s = Service(web_driver)
driver = webdriver.Chrome(service = s)  # Optional argument, if not specified will search path.


## 2. Get the web page link to explore
driver.get(link_web_page)
time.sleep(45) # Let the user actually see something!


## 3. Look for the "Prodcut Group" input search box on the webpage, click on it to expand drop down menu
timeout=10
wait = WebDriverWait(driver, timeout)
product_group_span = wait.until(EC.element_to_be_clickable((By.XPATH, 
                "//ng-select[contains(@class, 'ng-select')]//span[contains(text(), 'HBM3E (12H)')]/ancestor::ng-select"
            ))
        )
product_group_span.click()
time.sleep(1)


## 4. Click on a dropdown menu option, select "HBM4 (12H)"
option = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='HBM4 (12H)']")))
option.click()
time.sleep(1)


## 5. Locate the "More Options" menu
more_options_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'More Options')]")))
driver.execute_script("arguments[0].scrollIntoView(true);", more_options_button)
more_options_button.click()
time.sleep(1)


## 6. Wait for the section that contains the label "Asm Eng Lot Descript" and click on it
section = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//span[text()='Asm Eng Lot Descript']/ancestor::div[contains(@class, 'text-start')]")))
dropdown_container = section.find_element(By.XPATH, ".//div[contains(@class, 'ng-select-container')]")
dropdown_container.click()


## 7. Click "Unselect All"
unselect_all_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//button[@title='Unselect All' and contains(text(), 'Unselect All')]"))
)
unselect_all_button.click()

## 8. Select the QLC lots
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CLASS_NAME, "ng-dropdown-panel-items"))
)
options = driver.find_elements(By.XPATH, "//div[@role='option' and contains(., '6D_APTD_QLC')]")
for option in options:
    # Scroll into view and click each matching option
    driver.execute_script("arguments[0].scrollIntoView(true);", option)
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable(option)).click()

## 9. Click "Filter" button data to load the data on ATYMS graph
filter_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn-warning') and contains(text(), 'Filter')]")))
driver.execute_script("arguments[0].scrollIntoView(true);", filter_button)
filter_button.click()


## 10. Click the C2PP data only
# Wait for the button containing 'C2PP' to be clickable
c2pp_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'C2PP')]")))
# Scroll into view and click C2PP data
driver.execute_script("arguments[0].scrollIntoView(true);", c2pp_button)
c2pp_button.click()


## 11. Click the latest month data set available (fixed to '202510') # TO-DO: make it depending on current month
data_point = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'path[aria-label*="Month 202510, C2PP"]')))
driver.execute_script("arguments[0].scrollIntoView(true);", data_point)
data_point.click()


## 12. Click "Download raw data" button
download_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[@title='Download FID and Wafer level raw data.']")))
driver.execute_script("arguments[0].scrollIntoView(true);", download_button)
download_button.click()


## 13. Wait and Quit script
time.sleep(10) # Let the user actually see something!
driver.quit()
# display.stop()
