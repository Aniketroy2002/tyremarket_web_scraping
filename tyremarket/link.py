from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time

# Set up the Selenium WebDriver without using chromedriver.exe
options = webdriver.ChromeOptions()
options.add_argument("--headless")  # Run in headless mode (optional)
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Define the URL of the website
url = "https://www.tyremarket.com/Firestone-Car-Tyres"
driver.get(url)
time.sleep(5)  # Wait for the page to load

# Find all product links
product_elements = driver.find_elements(By.CSS_SELECTOR, ".productlist.tm-productlis a")
product_links = [element.get_attribute("href") for element in product_elements if element.get_attribute("href")]

# Close the driver
driver.quit()

# Save links to an Excel file
df = pd.DataFrame(product_links, columns=["Links"])
df.to_excel("input_links.xlsx", index=False)

print("Scraping completed. Links saved.")