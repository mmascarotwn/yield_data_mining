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

## Look for the input search box on the webpage, fill in the infor & click the search button
# search_box = driver.find_element(By.CSS_SELECTOR, "ng-value-label ng-star-inserted")
# search_box = driver.find_element("name", "ID")
# product_group_span = driver.find_element(By.XPATH, "//select[text()='HBM3E (12H)']")

timeout=10
wait = WebDriverWait(driver, timeout)
# product_group_span = driver.find_element(By.XPATH, "//span[@class='ng-value-label ng-star-inserted']/div[text()='HBM3E (12H)']")
product_group_span = wait.until(EC.element_to_be_clickable((By.XPATH, 
                "//ng-select[contains(@class, 'ng-select')]//span[contains(text(), 'HBM3E (12H)')]/ancestor::ng-select"
            ))
        )

product_group_span.click()
time.sleep(10)

WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//div[@class='dropdown-menu show']/a")))

click_option = driver.find_element(By.XPATH, "//div[@class='dropdown-menu show']/a[text()='HBM4 (12H)']")
click_option.click()

# print(product_group_span.text)
time.sleep(10) # Let the user actually see something!

# search_box.send_keys(Product_ID)
# search_box.send_keys(Keys.RETURN) # hit return after you enter search text

## Copy element in the loaded page
# time.sleep(3) # Let the user actually see something!
# td_element = driver.find_element(By.CSS_SELECTOR, "tr.ReportRow1")
# print("Wafer ID is: ", td_element.text)

driver.quit()
# display.stop()
