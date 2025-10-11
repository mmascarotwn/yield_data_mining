"""
Selenium function to locate and click the 'HBM3E (12H)' dropdown menu.
This module provides a minimal implementation for interacting with the dropdown.
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time


def click_hbm3e_dropdown(driver, timeout=10):
    """
    Locate and click the 'HBM3E (12H)' dropdown menu.
    
    Args:
        driver: Selenium WebDriver instance
        timeout (int): Maximum time to wait for elements (default: 10 seconds)
    
    Returns:
        bool: True if successful, False otherwise
    
    Raises:
        TimeoutException: If elements are not found within timeout
        NoSuchElementException: If drsopdown elements are not found
    """
    try:
        wait = WebDriverWait(driver, timeout)
        
        # Method 1: Try to click the dropdown by finding the ng-select container
        # Look for the ng-select container that contains HBM3E (12H)
        dropdown_container = wait.until(
            EC.element_to_be_clickable((
                By.XPATH, 
                "//ng-select[contains(@class, 'ng-select')]//span[contains(text(), 'HBM3E (12H)')]/ancestor::ng-select"
            ))
        )
        
        # Click to open/interact with dropdown
        dropdown_container.click()
        print("Successfully clicked the HBM3E (12H) dropdown container")
        
        # Wait a moment for dropdown to open
        time.sleep(0.5)
        
        return True
        
    except TimeoutException:
        print("Timeout: Could not find HBM3E (12H) dropdown container")
        return False
    except Exception as e:
        print(f"Error clicking HBM3E dropdown: {str(e)}")
        return False


def click_hbm3e_dropdown_alternative(driver, timeout=10):
    """
    Alternative method to locate and click the 'HBM3E (12H)' dropdown menu.
    Uses more specific selectors based on the HTML structure.
    
    Args:
        driver: Selenium WebDriver instance
        timeout (int): Maximum time to wait for elements (default: 10 seconds)
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        wait = WebDriverWait(driver, timeout)
        
        # Method 2: Click on the arrow or input area to open dropdown
        # Look for the ng-select container with the specific structure
        dropdown_arrow = wait.until(
            EC.element_to_be_clickable((
                By.XPATH, 
                "//ng-select//span[@class='ng-arrow-wrapper']//span[@class='ng-arrow']"
            ))
        )
        
        dropdown_arrow.click()
        print("Successfully clicked the dropdown arrow")
        
        # Wait for dropdown panel to appear
        time.sleep(0.5)
        
        # Now click on the HBM3E (12H) option if it's not already selected
        hbm3e_option = wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//ng-dropdown-panel//div[@role='option']//span[contains(text(), 'HBM3E (12H)')]"
            ))
        )
        
        hbm3e_option.click()
        print("Successfully selected HBM3E (12H) option")
        
        return True
        
    except TimeoutException:
        print("Timeout: Could not find HBM3E (12H) dropdown elements")
        return False
    except Exception as e:
        print(f"Error in alternative method: {str(e)}")
        return False


def select_hbm3e_from_dropdown(driver, timeout=10):
    """
    Comprehensive function to handle the HBM3E (12H) dropdown selection.
    Tries multiple approaches to ensure successful interaction.
    
    Args:
        driver: Selenium WebDriver instance
        timeout (int): Maximum time to wait for elements (default: 10 seconds)
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        wait = WebDriverWait(driver, timeout)
        
        # First, try to find if HBM3E (12H) is already selected
        try:
            current_selection = driver.find_element(
                By.XPATH, 
                "//ng-select//span[@class='ng-value-label ng-star-inserted' and contains(text(), 'HBM3E (12H)')]"
            )
            if current_selection.is_displayed():
                print("HBM3E (12H) is already selected")
                return True
        except NoSuchElementException:
            print("HBM3E (12H) is not currently selected, proceeding to select it")
        
        # Try to click the dropdown to open it
        # Look for the ng-select container
        dropdown_container = wait.until(
            EC.presence_of_element_located((
                By.XPATH, 
                "//ng-select[contains(@class, 'ng-select')]"
            ))
        )
        
        # Click on the dropdown to open it
        ActionChains(driver).click(dropdown_container).perform()
        time.sleep(1)
        
        # Wait for dropdown panel to be visible
        dropdown_panel = wait.until(
            EC.visibility_of_element_located((
                By.XPATH,
                "//ng-dropdown-panel[@role='listbox']"
            ))
        )
        
        # Find and click the HBM3E (12H) option
        hbm3e_option = wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//ng-dropdown-panel//div[@role='option']//span[@class='ng-option-label ng-star-inserted' and contains(text(), 'HBM3E (12H)')]"
            ))
        )
        
        hbm3e_option.click()
        print("Successfully selected HBM3E (12H) from dropdown")
        
        # Wait a moment for selection to take effect
        time.sleep(0.5)
        
        return True
        
    except TimeoutException:
        print("Timeout: Could not complete HBM3E (12H) selection")
        return False
    except Exception as e:
        print(f"Error selecting HBM3E from dropdown: {str(e)}")
        return False


def demo_usage():
    """
    Demo function showing how to use the HBM3E dropdown functions.
    """
    # Initialize WebDriver (Chrome example)
    driver = webdriver.Chrome()
    
    try:
        # Navigate to your target page
        # driver.get("your_target_url_here")
        
        # Try to select HBM3E (12H) from dropdown
        success = select_hbm3e_from_dropdown(driver)
        
        if success:
            print("HBM3E (12H) dropdown interaction completed successfully")
        else:
            print("Failed to interact with HBM3E (12H) dropdown")
            
    except Exception as e:
        print(f"Demo error: {str(e)}")
    finally:
        # Don't forget to close the driver
        driver.quit()


if __name__ == "__main__":
    # Uncomment the line below to run the demo
    # demo_usage()
    print("HBM3E dropdown module loaded. Use select_hbm3e_from_dropdown() function.")
