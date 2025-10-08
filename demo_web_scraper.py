#!/usr/bin/env python3
"""
Web Scraper Demo Script

This script demonstrates how to use the web_scraper module for scraping
website data and integrating it with Excel files.

Usage:
    python demo_web_scraper.py

Features demonstrated:
- Setting up web scraper configuration
- Defining websites and scraping rules
- Executing web scraping operations
- Integrating scraped data with Excel files
"""

import sys
import os
from pathlib import Path

# Add the parent directory to the path so we can import our modules
current_dir = Path(__file__).parent
parent_dir = current_dir.parent.parent
sys.path.insert(0, str(parent_dir))

from src.utils.web_scraper import WebScraper, WebScrapingConfig, create_sample_config
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def demo_basic_configuration():
    """Demonstrate basic web scraper configuration."""
    print("\n=== Demo: Basic Configuration ===")
    
    # Create a new configuration
    config = WebScrapingConfig()
    
    # Add websites
    config.add_website(
        name="httpbin",
        url="https://httpbin.org/html",
        description="HTTPBin HTML test page"
    )
    
    config.add_website(
        name="quotes",
        url="http://quotes.toscrape.com/",
        description="Quotes to scrape - testing website"
    )
    
    # Add scraping rules
    config.add_scraping_rule(
        rule_name="extract_page_title",
        website_name="httpbin",
        selector="h1",
        action_type="extract"
    )
    
    config.add_scraping_rule(
        rule_name="extract_first_quote",
        website_name="quotes",
        selector=".quote .text",
        action_type="extract"
    )
    
    config.add_scraping_rule(
        rule_name="extract_first_author",
        website_name="quotes",
        selector=".quote .author",
        action_type="extract"
    )
    
    print(f"✅ Configuration created with {len(config.websites)} websites")
    print(f"✅ Configuration has {len(config.scraping_rules)} scraping rules")
    
    return config


def demo_scraping_workflow():
    """Demonstrate the complete scraping workflow."""
    print("\n=== Demo: Complete Scraping Workflow ===")
    
    try:
        # Create scraper instance
        scraper = WebScraper()
        
        # Check dependencies
        if not scraper.check_dependencies():
            print("❌ Dependencies not available. Please install: pip install selenium beautifulsoup4 requests")
            return False
            
        # Load configuration
        config = demo_basic_configuration()
        scraper.config = config
        
        # Setup web driver (this might fail if Chrome/Firefox not installed)
        print("🔧 Setting up web driver...")
        if not scraper.setup_driver('chrome', headless=True):
            print("⚠️  Chrome driver setup failed. This is expected if Chrome is not installed.")
            print("   In production, ensure Chrome/Firefox and respective drivers are installed.")
            return False
            
        # Demonstrate scraping (placeholder - would need actual browser)
        print("🌐 Web driver setup successful!")
        print("📋 In production, you would now:")
        print("   1. Navigate to websites")
        print("   2. Execute scraping rules")
        print("   3. Extract data")
        print("   4. Integrate with Excel files")
        
        return True
        
    except Exception as e:
        logger.error(f"Demo error: {e}")
        return False
    finally:
        # Cleanup
        if 'scraper' in locals():
            scraper.cleanup()


def demo_configuration_save_load():
    """Demonstrate saving and loading configuration."""
    print("\n=== Demo: Configuration Save/Load ===")
    
    # Create configuration
    config = demo_basic_configuration()
    
    # Save configuration
    config_path = "demo_config.json"
    config.save_config(config_path)
    print(f"✅ Configuration saved to {config_path}")
    
    # Load configuration
    new_config = WebScrapingConfig()
    new_config.load_config(config_path)
    print(f"✅ Configuration loaded with {len(new_config.websites)} websites")
    
    # Clean up
    if os.path.exists(config_path):
        os.remove(config_path)
        print("🧹 Cleaned up demo config file")
    
    return True


def demo_excel_integration_placeholder():
    """Demonstrate Excel integration (placeholder)."""
    print("\n=== Demo: Excel Integration (Placeholder) ===")
    
    # This would be the typical workflow:
    sample_scraped_data = {
        "page_title": "Example Page Title",
        "quote_text": "The way to get started is to quit talking and begin doing.",
        "quote_author": "Walt Disney",
        "scrape_timestamp": "2025-10-09 10:30:00"
    }
    
    print("📊 Sample scraped data:")
    for key, value in sample_scraped_data.items():
        print(f"   {key}: {value}")
    
    print("\n📝 In production workflow:")
    print("   1. Select Excel file using scraper.select_excel_file()")
    print("   2. Load Excel data using scraper.load_excel_file()")
    print("   3. Add scraped data using scraper.add_scraped_data_to_excel()")
    print("   4. Save updated file using scraper.save_excel_file()")
    
    return True


def main():
    """Run all demonstrations."""
    print("🌐 Web Scraper Demo")
    print("=" * 50)
    
    # Run demonstrations
    demos = [
        ("Basic Configuration", demo_basic_configuration),
        ("Configuration Save/Load", demo_configuration_save_load),
        ("Excel Integration", demo_excel_integration_placeholder),
        ("Complete Workflow", demo_scraping_workflow),
    ]
    
    results = []
    for demo_name, demo_func in demos:
        try:
            print(f"\n🚀 Running: {demo_name}")
            result = demo_func()
            results.append((demo_name, result))
            status = "✅ PASSED" if result else "❌ FAILED"
            print(f"   {status}")
        except Exception as e:
            logger.error(f"Demo '{demo_name}' failed: {e}")
            results.append((demo_name, False))
            print(f"   ❌ FAILED: {e}")
    
    # Summary
    print("\n" + "=" * 50)
    print("📋 Demo Summary:")
    passed = sum(1 for _, result in results if result)
    total = len(results)
    
    for demo_name, result in results:
        status = "✅" if result else "❌"
        print(f"   {status} {demo_name}")
    
    print(f"\n🎯 Results: {passed}/{total} demos passed")
    
    if passed == total:
        print("🎉 All demos completed successfully!")
        print("\n📚 Next steps:")
        print("   1. Install web scraping dependencies: pip install selenium beautifulsoup4 requests")
        print("   2. Install browser drivers (ChromeDriver or GeckoDriver)")
        print("   3. Configure your websites and scraping rules")
        print("   4. Test with your actual Excel files")
    else:
        print("⚠️  Some demos failed - this is expected if dependencies are not installed")


if __name__ == "__main__":
    main()
