import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from pyvirtualdisplay import Display


chrome_options = Options()
chrome_options.add_argument("--headless=new") # Run in headless mode
chrome_options.add_argument("--disable-gpu") # Recommended for headless
chrome_options.add_argument("--no-sandbox") # Often required for headless mode, especially in Linux environments
chrome_options.add_argument("--disable-dev-shm-usage") # Often required for headless mode, especially in Linux environments

# display = Display(visible=0, size=(800, 600))
# display.start()

s = Service(r"C:\Users\mmascaro\Documents\VisualStudio Projects\venv_test_project\chromedriver.exe")
driver = webdriver.Chrome(service = s)  # Optional argument, if not specified will search path.
driver.get('http://mamweb.tatw.micron.com/MAMWeb/site/ASMTB/')
# time.sleep(3) # Let the user actually see something!
search_box = driver.find_element("name", "ID")
search_box.send_keys("788479K.0BL")
search_box.send_keys(Keys.RETURN) # hit return after you enter search text
# time.sleep(3) # Let the user actually see something!
td_element = driver.find_element(By.CSS_SELECTOR, "tr.ReportRow1")
print(td_element.text)

driver.quit()
# display.stop()
