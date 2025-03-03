import csv
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# CSV file setup
output_file = "results.csv"
headers = ["Roll Number", "Name", "SPI", "CPI"]

# Write headers to the CSV
with open(output_file, "w", newline="") as file:
    writer = csv.writer(file)
    writer.writerow(headers)

rNoYear = input("Enter your year of joining::")
rNoBranch = input("Enter your branch code (e.g., CSE = 1)::")
rollNoMin = input("Enter the First Roll Number in your branch::")
rollNoMax = input("Enter the Last Roll Number in your branch::")

rollNumberBaseStr = f"{rNoYear[-2:]}010{rNoBranch}"  # e.g., 23 + 0 + 1 + 0 = 23010

# Set up the WebDriver using ChromeDriverManager
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

for i in range(int(rollNoMin), int(rollNoMax) + 1):
    try:
        # Open the website
        driver.get(f"http://14.139.202.226/{rNoYear[2:]}batchresult/")

        # Wait for the page to load
        time.sleep(2)

        # Locate the roll number input field
        roll_number_field = driver.find_element(By.NAME, "enX")
        roll_number = rollNumberBaseStr + f"{i:03d}"
        print(f"Processing Roll Number: {roll_number}")
        roll_number_field.send_keys(roll_number)

        # Locate the submit button and click it
        submit_button = driver.find_element(By.XPATH, "//input[@type='submit']")
        submit_button.click()

        # Wait for the new page to load
        time.sleep(3)

        # Determine the number of semesters (divs) available
        divs = driver.find_elements(By.XPATH, '//*[@id="divToPrint"]/div')
        latest_sem_div_index = len(divs)  # Get the last semester index

        # Dynamically create XPath for the latest semester
        spi_xpath = f'//*[@id="divToPrint"]/div[{latest_sem_div_index}]/center/table/tbody/tr[1]/td/b[1]'
        cpi_xpath = f'//*[@id="divToPrint"]/div[{latest_sem_div_index}]/center/table/tbody/tr[1]/td/b[2]'

        # Extract values
        name = driver.find_element(By.XPATH, '//*[@id="divToPrint"]/table[1]/tbody/tr[1]/td[2]').text[2:]
        spi = driver.find_element(By.XPATH, spi_xpath).text[2:]  # Adjust slicing if necessary
        cpi = driver.find_element(By.XPATH, cpi_xpath).text[2:]

        # Append the results to the CSV file
        with open(output_file, "a", newline="") as file:
            writer = csv.writer(file)
            writer.writerow([roll_number, name, spi, cpi])

        print(f"Processed Roll Number {roll_number}: {name}, SPI: {spi}, CPI: {cpi}")

    except Exception as e:
        print(f"An error occurred for Roll Number {roll_number}: {e}")

driver.quit()
