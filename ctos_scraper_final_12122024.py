import os
import json
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from datetime import datetime

counter = 0
flag = False

# File path and sheet details
file_path = r"C:\Users\User\Desktop\Analytics\Database\Raw Data.xlsx"
sheet_name = "Extract"
company_column = "Company Name"
nature_column = "Nature of Business"

# Checkpoint directory and file
checkpoint_dir = r"C:\Users\User\Desktop\Analytics\Checkpoint"
os.makedirs(checkpoint_dir, exist_ok=True)
checkpoint_file = os.path.join(checkpoint_dir, "checkpoint.json")

# Driver path
driver_path = r"C:\chromedriver\chromedriver.exe"

# Verify required files and paths
if not os.path.exists(file_path):
    print(f"Error: File not found: {file_path}")
    exit()

if not os.path.exists(driver_path):
    print(f"Error: ChromeDriver not found: {driver_path}")
    exit()

# Load Excel data
try:
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df[nature_column] = df[nature_column].astype(str)
except Exception as e:
    print(f"Error loading Excel file: {e}")
    exit()

# Set up Chrome options (Remove this part of the code to run with GUI)
chrome_options = Options()
chrome_options.add_argument("--headless")  # Enable headless mode
chrome_options.add_argument("--disable-gpu")  # Disable GPU acceleration (recommended for headless mode)
chrome_options.add_argument("--window-size=1920,1080")  # Set window size (recommended for headless mode)
chrome_options.add_argument("--no-sandbox")  # Bypass OS security model (useful in some environments)
chrome_options.add_argument("--disable-dev-shm-usage")  # Overcome limited resource issues in some systems

# Selenium WebDriver setup
service = Service(driver_path)
driver = webdriver.Chrome(service=service) #options=chrome_options)

# Navigate to CTOS search page
try:
    driver.get("https://businessreport.ctoscredit.com.my/oneoffreport/home")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "searchKey")))
except Exception as e:
    print(f"Error accessing CTOS website: {e}")
    driver.quit()
    exit()

# Utility functions
def trim_company_name(company_name):
    return company_name.split(" Sdn")[0]

def load_checkpoint():
    if os.path.exists(checkpoint_file):
        with open(checkpoint_file, "r") as file:
            return json.load(file)
    return {"last_index": -1}

def save_checkpoint(last_index):
    with open(checkpoint_file, "w") as file:
        json.dump({"last_index": last_index}, file)

def initialize_log_file():
    timestamp = datetime.now().strftime("%d%m%y_%H%M%S")
    log_file_path = fr"C:\Users\User\Desktop\Analytics\Logs\Log_{timestamp}.log"
    os.makedirs(os.path.dirname(log_file_path), exist_ok=True)
    return log_file_path

def log_search_results(company_name, results_count, results_details):
    global counter
    try:
        with open(log_file_path, mode="a", encoding="utf-8") as log_file:
            #Based on the log file skips all the criteria listed below *can add more if you need
            skip_criteria1 = "Error"
            skip_criteria2 = "no such element"
            null_data = "Error: Message:" #Company not Found
            log_file.write("---------------------------------------------------\n")
            log_file.write(f"Company: {company_name} | Results: {results_count}\n")
            log_file.write("---------------------------------------------------\n")
            log_file.write(f"{results_details}\n")
            log_file.write("---------------------------------------------------\n")
            
            #Progession showcase
            if skip_criteria1 and skip_criteria2 in results_details:
                print(f"    {company_name} Company's Nature of Business is Not Found on CTOS website...")
            else:
                print(f"[+] Progression: {counter} [{company_name}]")
                if null_data in results_details:
                    print(f"    {company_name} Company is Not Found on CTOS Website...")
                counter = counter + 1

    except Exception as e:
        print(f"Error writing to log file: {e}")

# Initialize log file
log_file_path = initialize_log_file()

# Process batch of companies (CTOS website only allow 500 search per session)
checkpoint = load_checkpoint()
start_index = checkpoint["last_index"] + 1
end_index = min(start_index + 500, len(df))
#Check Completion Status
if start_index >= len(df):
    print("All companies have been processed.")
    driver.quit()
    exit()

print(f"Processing companies from index {start_index} to {end_index - 1}.")

def get_nature_of_business(company_name):
    try:
        trimmed_name = trim_company_name(company_name)

        # Perform search
        search_bar = driver.find_element(By.ID, "searchKey")
        search_bar.clear()
        search_bar.send_keys(trimmed_name)
        time.sleep(1)
        search_bar.send_keys(Keys.RETURN)

        # Wait for results to load
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "tbl_scroll")))
        rows = driver.find_elements(By.CLASS_NAME, "mat-row.ng-star-inserted")

        # Log results
        if rows:
            results_details = "\n".join(row.text.strip() for row in rows)
            log_search_results(trimmed_name, len(rows), results_details)

            rows[0].find_element(By.TAG_NAME, "a").click()
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "table.table-striped.ng-star-inserted"))
            )

            time.sleep(1)
            return driver.find_element(By.CSS_SELECTOR, "tr:nth-child(3) .ng-star-inserted").text
        else:
            log_search_results(trimmed_name, 0, "No results found")
            return None
    except (TimeoutException, NoSuchElementException) as e:
        log_search_results(trimmed_name, 0, f"Error: {e}")
        return None

try:
    for index in range(start_index, end_index):
        company_name = df.loc[index, company_column]
        nature = get_nature_of_business(company_name)
        df.at[index, nature_column] = nature
except Exception as e:
    print(f"Error processing companies: {e}")

save_checkpoint(end_index - 1)

# Save results
timestamp = datetime.now().strftime("%H%M%S_%d%m%y")
output_file_path = rf"C:\Users\User\Desktop\Analytics\Output\Database_Updated_{timestamp}.xlsx" #Save Excel File

try:
    df.to_excel(output_file_path, sheet_name=sheet_name, index=False)
    print(f"Results saved to {output_file_path}") 
except Exception as e:
    print(f"Error saving results: {e}")

driver.quit()
