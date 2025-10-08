#!/usr/bin/env python3
"""
Web Scraper Module

This module provides functionality to scrape data from websites and add the scraped values
as new columns to Excel files. It includes web interaction capabilities such as filling forms,
clicking buttons, and extracting specific data from web pages.

Features:
- Configurable target websites and data extraction points
- Interactive web page functionality (form filling, button clicking)
- Excel file integration for adding scraped data as new columns
- Configurable delay and retry mechanisms
- Error handling and logging for web scraping operations
- Support for different web drivers (Chrome, Firefox, etc.)
- Data validation and cleaning before Excel integration

Placeholder Structure:
- Website navigation and interaction
- Data extraction from web pages
- Excel file integration
- Configuration management
"""

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import os
import time
import logging
from pathlib import Path
from typing import Optional, List, Dict, Tuple, Any
from datetime import datetime
import json

# Web scraping libraries
try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    from selenium.webdriver.firefox.options import Options as FirefoxOptions
    from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
    SELENIUM_AVAILABLE = True
except ImportError:
    SELENIUM_AVAILABLE = False
    
try:
    import requests
    from bs4 import BeautifulSoup
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False

# Configuration - Web scraping settings
DEFAULT_TIMEOUT = 10  # seconds
DEFAULT_DELAY = 1     # seconds between actions
MAX_RETRIES = 3       # maximum retry attempts
HEADLESS_MODE = False # run browser in headless mode

# Configuration - Excel integration
TARGET_SHEET_NAME = 'Sheet1'  # Default sheet to add scraped data
SCRAPED_DATA_PREFIX = 'web_'  # Prefix for scraped column names

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class WebScrapingConfig:
    """
    Configuration class for web scraping operations.
    """
    def __init__(self):
        self.websites: Dict[str, Dict] = {}
        self.scraping_rules: Dict[str, Dict] = {}
        self.excel_mapping: Dict[str, str] = {}
        
    def add_website(self, name: str, url: str, description: str = ""):
        """
        Add a website configuration.
        
        Args:
            name: Unique name for the website
            url: URL to scrape
            description: Optional description
        """
        self.websites[name] = {
            'url': url,
            'description': description,
            'added_date': datetime.now().isoformat()
        }
        
    def add_scraping_rule(self, rule_name: str, website_name: str, 
                         selector: str, action_type: str = 'extract',
                         **kwargs):
        """
        Add a scraping rule for data extraction.
        
        Args:
            rule_name: Unique name for the rule
            website_name: Name of the website (must exist in websites)
            selector: CSS selector or XPath for the element
            action_type: Type of action ('extract', 'click', 'fill', 'wait')
            **kwargs: Additional parameters for the action
        """
        self.scraping_rules[rule_name] = {
            'website': website_name,
            'selector': selector,
            'action_type': action_type,
            'parameters': kwargs
        }
        
    def save_config(self, config_path: str):
        """Save configuration to JSON file."""
        config_data = {
            'websites': self.websites,
            'scraping_rules': self.scraping_rules,
            'excel_mapping': self.excel_mapping
        }
        with open(config_path, 'w') as f:
            json.dump(config_data, f, indent=2)
            
    def load_config(self, config_path: str):
        """Load configuration from JSON file."""
        try:
            with open(config_path, 'r') as f:
                config_data = json.load(f)
                self.websites = config_data.get('websites', {})
                self.scraping_rules = config_data.get('scraping_rules', {})
                self.excel_mapping = config_data.get('excel_mapping', {})
        except FileNotFoundError:
            logger.warning(f"Config file not found: {config_path}")
        except json.JSONDecodeError as e:
            logger.error(f"Invalid config file format: {e}")


class WebScraper:
    """
    Main web scraper class with Excel integration capabilities.
    """
    
    def __init__(self):
        self.driver: Optional[webdriver.Chrome] = None
        self.config = WebScrapingConfig()
        self.scraped_data: Dict[str, Any] = {}
        self.excel_file_path: Optional[str] = None
        self.excel_data: Optional[Dict[str, pd.DataFrame]] = None
        
    def check_dependencies(self) -> bool:
        """
        Check if required dependencies are available.
        
        Returns:
            True if all dependencies are available, False otherwise
        """
        missing_deps = []
        if not SELENIUM_AVAILABLE:
            missing_deps.append("selenium")
        if not REQUESTS_AVAILABLE:
            missing_deps.append("requests and beautifulsoup4")
            
        if missing_deps:
            logger.error(f"Missing dependencies: {', '.join(missing_deps)}")
            messagebox.showerror("Missing Dependencies", 
                f"Please install the following packages:\n"
                f"pip install {' '.join(missing_deps.replace('and ', '').split())}")
            return False
        return True
        
    def setup_driver(self, browser: str = 'chrome', headless: bool = HEADLESS_MODE) -> bool:
        """
        Setup web driver for browser automation.
        
        Args:
            browser: Browser type ('chrome', 'firefox')
            headless: Run in headless mode
            
        Returns:
            True if driver setup successful, False otherwise
        """
        try:
            if not SELENIUM_AVAILABLE:
                logger.error("Selenium not available. Cannot setup web driver.")
                return False
                
            if browser.lower() == 'chrome':
                options = ChromeOptions()
                if headless:
                    options.add_argument('--headless')
                options.add_argument('--no-sandbox')
                options.add_argument('--disable-dev-shm-usage')
                self.driver = webdriver.Chrome(options=options)
            elif browser.lower() == 'firefox':
                options = FirefoxOptions()
                if headless:
                    options.add_argument('--headless')
                self.driver = webdriver.Firefox(options=options)
            else:
                logger.error(f"Unsupported browser: {browser}")
                return False
                
            self.driver.set_page_load_timeout(DEFAULT_TIMEOUT)
            logger.info(f"Web driver setup successful: {browser}")
            return True
            
        except WebDriverException as e:
            logger.error(f"Failed to setup web driver: {e}")
            messagebox.showerror("Driver Setup Error", 
                f"Failed to setup web driver. Please ensure {browser} is installed.\n"
                f"Error: {str(e)}")
            return False
            
    def navigate_to_url(self, url: str) -> bool:
        """
        Navigate to a specific URL.
        
        Args:
            url: URL to navigate to
            
        Returns:
            True if navigation successful, False otherwise
        """
        try:
            if not self.driver:
                logger.error("Web driver not initialized")
                return False
                
            logger.info(f"Navigating to: {url}")
            self.driver.get(url)
            time.sleep(DEFAULT_DELAY)
            return True
            
        except Exception as e:
            logger.error(f"Failed to navigate to {url}: {e}")
            return False
            
    def fill_form_field(self, selector: str, value: str, 
                       selector_type: str = 'css') -> bool:
        """
        Fill a form field with a value.
        
        Args:
            selector: CSS selector or XPath for the element
            value: Value to enter
            selector_type: 'css' or 'xpath'
            
        Returns:
            True if successful, False otherwise
        """
        try:
            if not self.driver:
                logger.error("Web driver not initialized")
                return False
                
            wait = WebDriverWait(self.driver, DEFAULT_TIMEOUT)
            
            if selector_type == 'css':
                element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
            elif selector_type == 'xpath':
                element = wait.until(EC.presence_of_element_located((By.XPATH, selector)))
            else:
                logger.error(f"Invalid selector type: {selector_type}")
                return False
                
            element.clear()
            element.send_keys(value)
            time.sleep(DEFAULT_DELAY)
            
            logger.info(f"Filled form field {selector} with value: {value}")
            return True
            
        except TimeoutException:
            logger.error(f"Timeout waiting for element: {selector}")
            return False
        except Exception as e:
            logger.error(f"Failed to fill form field {selector}: {e}")
            return False
            
    def click_element(self, selector: str, selector_type: str = 'css') -> bool:
        """
        Click an element on the page.
        
        Args:
            selector: CSS selector or XPath for the element
            selector_type: 'css' or 'xpath'
            
        Returns:
            True if successful, False otherwise
        """
        try:
            if not self.driver:
                logger.error("Web driver not initialized")
                return False
                
            wait = WebDriverWait(self.driver, DEFAULT_TIMEOUT)
            
            if selector_type == 'css':
                element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
            elif selector_type == 'xpath':
                element = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
            else:
                logger.error(f"Invalid selector type: {selector_type}")
                return False
                
            element.click()
            time.sleep(DEFAULT_DELAY)
            
            logger.info(f"Clicked element: {selector}")
            return True
            
        except TimeoutException:
            logger.error(f"Timeout waiting for clickable element: {selector}")
            return False
        except Exception as e:
            logger.error(f"Failed to click element {selector}: {e}")
            return False
            
    def extract_text(self, selector: str, selector_type: str = 'css') -> Optional[str]:
        """
        Extract text from an element.
        
        Args:
            selector: CSS selector or XPath for the element
            selector_type: 'css' or 'xpath'
            
        Returns:
            Extracted text or None if failed
        """
        try:
            if not self.driver:
                logger.error("Web driver not initialized")
                return None
                
            wait = WebDriverWait(self.driver, DEFAULT_TIMEOUT)
            
            if selector_type == 'css':
                element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
            elif selector_type == 'xpath':
                element = wait.until(EC.presence_of_element_located((By.XPATH, selector)))
            else:
                logger.error(f"Invalid selector type: {selector_type}")
                return None
                
            text = element.text.strip()
            logger.info(f"Extracted text from {selector}: {text}")
            return text
            
        except TimeoutException:
            logger.error(f"Timeout waiting for element: {selector}")
            return None
        except Exception as e:
            logger.error(f"Failed to extract text from {selector}: {e}")
            return None
            
    def extract_attribute(self, selector: str, attribute: str, 
                         selector_type: str = 'css') -> Optional[str]:
        """
        Extract an attribute value from an element.
        
        Args:
            selector: CSS selector or XPath for the element
            attribute: Attribute name to extract
            selector_type: 'css' or 'xpath'
            
        Returns:
            Attribute value or None if failed
        """
        try:
            if not self.driver:
                logger.error("Web driver not initialized")
                return None
                
            wait = WebDriverWait(self.driver, DEFAULT_TIMEOUT)
            
            if selector_type == 'css':
                element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
            elif selector_type == 'xpath':
                element = wait.until(EC.presence_of_element_located((By.XPATH, selector)))
            else:
                logger.error(f"Invalid selector type: {selector_type}")
                return None
                
            value = element.get_attribute(attribute)
            logger.info(f"Extracted {attribute} from {selector}: {value}")
            return value
            
        except TimeoutException:
            logger.error(f"Timeout waiting for element: {selector}")
            return None
        except Exception as e:
            logger.error(f"Failed to extract {attribute} from {selector}: {e}")
            return None
            
    def wait_for_element(self, selector: str, timeout: int = DEFAULT_TIMEOUT,
                        selector_type: str = 'css') -> bool:
        """
        Wait for an element to be present on the page.
        
        Args:
            selector: CSS selector or XPath for the element
            timeout: Maximum time to wait in seconds
            selector_type: 'css' or 'xpath'
            
        Returns:
            True if element found, False if timeout
        """
        try:
            if not self.driver:
                logger.error("Web driver not initialized")
                return False
                
            wait = WebDriverWait(self.driver, timeout)
            
            if selector_type == 'css':
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
            elif selector_type == 'xpath':
                wait.until(EC.presence_of_element_located((By.XPATH, selector)))
            else:
                logger.error(f"Invalid selector type: {selector_type}")
                return False
                
            logger.info(f"Element found: {selector}")
            return True
            
        except TimeoutException:
            logger.error(f"Timeout waiting for element: {selector}")
            return False
        except Exception as e:
            logger.error(f"Error waiting for element {selector}: {e}")
            return False
            
    def scrape_website_data(self, website_name: str, scraping_rules: List[str]) -> Dict[str, Any]:
        """
        Execute multiple scraping rules on a website.
        
        Args:
            website_name: Name of the website to scrape
            scraping_rules: List of rule names to execute
            
        Returns:
            Dictionary with scraped data
        """
        results = {}
        
        try:
            if website_name not in self.config.websites:
                logger.error(f"Website not configured: {website_name}")
                return results
                
            url = self.config.websites[website_name]['url']
            
            if not self.navigate_to_url(url):
                return results
                
            for rule_name in scraping_rules:
                if rule_name not in self.config.scraping_rules:
                    logger.warning(f"Scraping rule not found: {rule_name}")
                    continue
                    
                rule = self.config.scraping_rules[rule_name]
                selector = rule['selector']
                action_type = rule['action_type']
                params = rule.get('parameters', {})
                
                if action_type == 'extract':
                    value = self.extract_text(selector, params.get('selector_type', 'css'))
                    if value:
                        results[rule_name] = value
                elif action_type == 'extract_attribute':
                    attribute = params.get('attribute', 'value')
                    value = self.extract_attribute(selector, attribute, params.get('selector_type', 'css'))
                    if value:
                        results[rule_name] = value
                elif action_type == 'fill':
                    fill_value = params.get('value', '')
                    self.fill_form_field(selector, fill_value, params.get('selector_type', 'css'))
                elif action_type == 'click':
                    self.click_element(selector, params.get('selector_type', 'css'))
                elif action_type == 'wait':
                    timeout = params.get('timeout', DEFAULT_TIMEOUT)
                    self.wait_for_element(selector, timeout, params.get('selector_type', 'css'))
                    
            logger.info(f"Scraped data from {website_name}: {results}")
            return results
            
        except Exception as e:
            logger.error(f"Error scraping website {website_name}: {e}")
            return results
            
    def select_excel_file(self) -> Optional[str]:
        """
        Open file dialog to select Excel file for data integration.
        
        Returns:
            Path to selected Excel file or None if cancelled
        """
        root = tk.Tk()
        root.withdraw()
        
        try:
            messagebox.showinfo("File Selection", "Please select the Excel file to add scraped data to")
            excel_file = filedialog.askopenfilename(
                title="Select Excel File for Web Scraping Integration",
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            )
            
            if not excel_file:
                messagebox.showwarning("Warning", "No file selected")
                return None
                
            self.excel_file_path = excel_file
            logger.info(f"Selected Excel file: {excel_file}")
            return excel_file
            
        except Exception as e:
            messagebox.showerror("Error", f"Error selecting file: {str(e)}")
            return None
        finally:
            root.destroy()
            
    def load_excel_file(self) -> bool:
        """
        Load the selected Excel file.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            if not self.excel_file_path:
                logger.error("No Excel file selected")
                return False
                
            # Load all sheets
            self.excel_data = pd.read_excel(self.excel_file_path, sheet_name=None)
            logger.info(f"Loaded Excel file with sheets: {list(self.excel_data.keys())}")
            return True
            
        except Exception as e:
            logger.error(f"Error loading Excel file: {e}")
            messagebox.showerror("Error", f"Error loading Excel file: {str(e)}")
            return False
            
    def add_scraped_data_to_excel(self, scraped_data: Dict[str, Any], 
                                 target_sheet: str = TARGET_SHEET_NAME) -> bool:
        """
        Add scraped data as new columns to the Excel file.
        
        Args:
            scraped_data: Dictionary with scraped data
            target_sheet: Sheet name to add data to
            
        Returns:
            True if successful, False otherwise
        """
        try:
            if not self.excel_data or target_sheet not in self.excel_data:
                logger.error(f"Target sheet not found: {target_sheet}")
                return False
                
            df = self.excel_data[target_sheet]
            
            # Add scraped data as new columns
            for key, value in scraped_data.items():
                column_name = f"{SCRAPED_DATA_PREFIX}{key}"
                df[column_name] = value  # This will add the same value to all rows
                logger.info(f"Added column '{column_name}' with value: {value}")
                
            self.excel_data[target_sheet] = df
            return True
            
        except Exception as e:
            logger.error(f"Error adding scraped data to Excel: {e}")
            return False
            
    def save_excel_file(self, output_path: Optional[str] = None) -> bool:
        """
        Save the Excel file with scraped data.
        
        Args:
            output_path: Optional output path. If None, creates new file.
            
        Returns:
            True if successful, False otherwise
        """
        try:
            if not self.excel_data:
                logger.error("No Excel data to save")
                return False
                
            if output_path is None:
                input_path = Path(self.excel_file_path)
                output_path = str(input_path.with_name(f"{input_path.stem}_with_scraped_data{input_path.suffix}"))
                
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for sheet_name, df in self.excel_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
            logger.info(f"Saved Excel file with scraped data: {output_path}")
            messagebox.showinfo("Success", f"Excel file saved with scraped data:\n{output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error saving Excel file: {e}")
            messagebox.showerror("Error", f"Error saving Excel file: {str(e)}")
            return False
            
    def cleanup(self):
        """Clean up resources (close driver, etc.)."""
        if self.driver:
            try:
                self.driver.quit()
                logger.info("Web driver closed")
            except Exception as e:
                logger.error(f"Error closing web driver: {e}")
                
    def run_complete_scraping_process(self) -> bool:
        """
        Run the complete web scraping and Excel integration process.
        This is a placeholder for the full workflow.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            # Check dependencies
            if not self.check_dependencies():
                return False
                
            # Setup web driver
            if not self.setup_driver():
                return False
                
            # Select Excel file
            if not self.select_excel_file():
                return False
                
            # Load Excel file
            if not self.load_excel_file():
                return False
                
            # TODO: Add actual scraping logic here
            # This is where users would configure websites and scraping rules
            # For now, we'll show a placeholder message
            
            messagebox.showinfo("Placeholder", 
                "Web scraper is ready for configuration.\n\n"
                "Next steps:\n"
                "1. Configure websites using config.add_website()\n"
                "2. Define scraping rules using config.add_scraping_rule()\n"
                "3. Execute scraping using scrape_website_data()\n"
                "4. Integrate data using add_scraped_data_to_excel()")
                
            return True
            
        except Exception as e:
            logger.error(f"Error in scraping process: {e}")
            messagebox.showerror("Error", f"Error in scraping process: {str(e)}")
            return False
        finally:
            self.cleanup()


def create_sample_config() -> WebScrapingConfig:
    """
    Create a sample configuration for demonstration purposes.
    
    Returns:
        Configured WebScrapingConfig instance
    """
    config = WebScrapingConfig()
    
    # Sample website configurations
    config.add_website(
        name="example_site",
        url="https://example.com",
        description="Example website for demonstration"
    )
    
    # Sample scraping rules
    config.add_scraping_rule(
        rule_name="extract_title",
        website_name="example_site",
        selector="h1",
        action_type="extract"
    )
    
    config.add_scraping_rule(
        rule_name="fill_search",
        website_name="example_site",
        selector="input[name='search']",
        action_type="fill",
        value="sample search term"
    )
    
    config.add_scraping_rule(
        rule_name="click_submit",
        website_name="example_site",
        selector="button[type='submit']",
        action_type="click"
    )
    
    return config


def run_web_scraper():
    """
    Convenience function to run the web scraper.
    """
    scraper = WebScraper()
    return scraper.run_complete_scraping_process()


if __name__ == "__main__":
    print("🌐 Starting Web Scraper...")
    print("📋 This module provides web scraping capabilities for Excel integration")
    print("🔧 Configure websites and scraping rules before use")
    print("📁 Will integrate scraped data into selected Excel files")
    
    # Create sample configuration
    sample_config = create_sample_config()
    print(f"📝 Sample configuration created with {len(sample_config.websites)} websites")
    
    success = run_web_scraper()
    
    if success:
        print("✅ Web scraper setup completed successfully!")
    else:
        print("❌ Web scraper setup failed or was cancelled")
