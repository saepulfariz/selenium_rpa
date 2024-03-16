from selenium.webdriver.chrome.service import Service
from selenium import webdriver
import openpyxl
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
# chrome_options.add_experimental_option("detach", False)


# service = Service(executable_path='E:/dev/python/RPA-Challenge/chromedriver-win64/chromedriver.exe')
service = Service(executable_path='./chromedriver-win64/chromedriver.exe')
driver = webdriver.Chrome(service=service, options=chrome_options)
driver.maximize_window()

# Open RPA website #
rpa_website = driver.get("http://www.rpachallenge.com/")

# paksa close
# driver.close();