import pandas as pd
import os
import glob
import threading
import requests
import time
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import sys
# Add the full path to "scripts" based on current script location
current_dir = os.path.dirname(os.path.abspath(__file__))
scripts_dir = os.path.join(current_dir, "scripts")
if scripts_dir not in sys.path:
    sys.path.append(scripts_dir)
from functions_drill import *
sys.path.append(os.path.relpath('config/'))
import secrets_local
# --- Constants & Configurations ---
INPUT_GES_RULE = "./input/GES Rules"
OUTPUT_DIR = "Output"
BUFFER_WRITE_INTERVAL = 10
global row_index
YELLOW = '\033[33m'
RESET = '\033[0m'
n = input(f"{YELLOW}Please close all Chrome browsers and processes before continuing, since this program uses multithreading to expedite the process of retrieving loan rules.\n\nPress any key to continue, once you have closed Chrome.\n\nThis program also uses multithreading to increase speed since drilling through fulfillment configuration windows takes a while.  But the number of threads that is best matched to your local environment is based on how reliable and fast your internet is.  Please choose one of the options below:\n\n\t1 (1 thread: Slow internet)\n\t2 (2 threads: Faster internet)\n\t3 (3 threads: Fastest internet)\n\n\tChoice: {RESET}")
# Prompt for environment
instance = ""
env = input(f"Run against (1) Sandbox or (2) Production? Enter 1 or 2: {RESET}").strip()
if env == "1":
    
    api_base_url = secrets_local.alma_base_url_sandbox
    instance = "sandbox"
elif env == "2":
    api_base_url = secrets_local.alma_base_url_prod
    instance = "prod"
else:
    print("Invalid selection. Exiting.")
    exit(1)
n = int(n)
if n in (1,2,3):
    N = n
else:
    print("\n\nInvalid choice.   Try again")
    sys.exit()
row_index_lock = threading.Lock()
row_index = 0
order_of_loan_policy_columns = []
order_of_request_policy_columns = []
def load_excel_file(directory):
    files = glob.glob(os.path.join(directory, "*.xlsx"))
    if not files:
        print(f"No Excel files found in {directory}. Exiting.")
        exit()
    return pd.read_excel(files[0], dtype="str", engine="openpyxl")
ges_rule_data = load_excel_file(INPUT_GES_RULE)
def init_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver
def worker_thread(thread_id, ges_rule_data_sorted):
    driver = init_driver()
    buffer = []
    current_parameter_Value = None
    def navigate_to_ges():
        driver.get(api_base_url)
        login(driver, secrets_local.username, secrets_local.password)
        time.sleep(30)
        try:
            modal = driver.find_element(By.XPATH, "//div[@id='onetrust-close-btn-container']//button")
            print("GDPR modal detected. Attempting to close it.")
            modal = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@id='onetrust-close-btn-container']//button"))
            )
            modal.click()
            print("GDPR modal closed.")
        except TimeoutException:
            print("No GDPR modal/ detected.")
        except:
            print("No GDPR modal")
        # AGS Service Availability Rules
        driver.get(api_base_url + "ng/page;u=%2Fful%2Faction%2FpageAction.do%3FxmlFileName%3Dgeneral_electronic_services.general_electronic_services.xml&almaConfiguration%3Dtrue&pageViewMode%3DEdit&pageBean.currentUrl%3DxmlFileName%253Dgeneral_electronic_services.general_electronic_services.xml%2526almaConfiguration%253Dtrue%2526pageViewMode%253DEdit%2526resetPaginationContext%253Dtrue%2526showBackButton%253Dfalse&pageBean.navigationBackUrl%3D..%252Faction%252Fhome.do&resetPaginationContext%3Dtrue&showBackButton%3Dfalse&pageBean.securityHashToken%3D-457533158646046076;ng=true")
        
        time.sleep(15)
        # Edit second General Electronic Service
        more_actions = safe_find_element(driver, By.ID, "input_gesList_1")
        more_actions.click()
        edit_menu = safe_find_element(driver, By.ID, "ROW_ACTION_gesList_1_c.ui.table.btn.edit")
        edit_menu.click()
        
        # Navigate to Service Availability Rules tab
        sa_rules = safe_find_element(driver, By.ID, "A_NAV_LINK_general_electronic_servicesrules_span")
        sa_rules.click()
        
    navigate_to_ges()
    
    while True:
        
        global row_index
        with row_index_lock:
            if row_index >= len(ges_rule_data_sorted):
                break
            
            row = ges_rule_data_sorted.iloc[row_index]
            row_index += 1
        
        title = row["Title"]
        year = row["Year"]
        parameter_Name_Value = row["Parameter_Name"]
        parameter_Operator_Value = row["Parameter_Operator"]
        parameter_Value = row["Parameter_Value"]
        
        parameter_Date = 'rft.date'
        parameter_Operator_Greaterthan = '>='
        
        if len(title) > 50:
            title = title[:50]
        # if len(title) > 40:
        #     title = f"{title[:40]} (OpenURL)"
            
        
        def add_rule():
            nonlocal current_parameter_Value
            try:
                time.sleep(60)                
                # Add New Rule
                add_rule = safe_find_element(driver, By.ID, "ADD_HIDERADIO_up_rules_uiconfiguration_rulesadd")
                add_rule.click()
                
                # Rule Name and Description
                time.sleep(2)
                name_input = safe_find_element(driver, By.ID, "pageBeanrulename")
                name_input.send_keys(title)
                
                description_input = safe_find_element(driver, By.ID, "pageBeanruledescription")
                description_input.send_keys(f"{parameter_Name_Value}={parameter_Value}")
                
                # Add New Parameter
                add_parameter = safe_find_element(driver, By.ID, "widgetId_Right_rulesaddParameter")
                add_parameter.click()
                # ISSN RULE
                parameter_name = safe_find_element(driver, By.ID, "pageBeanparameterName")
                parameter_name.send_keys(parameter_Name_Value)
                time.sleep(2)
                parameter_name_select = safe_find_element(driver, By.XPATH, "//li[@role='option']")
                parameter_name_select.click()
                time.sleep(2)
                parameter_operator = safe_find_element(driver, By.ID, "pageBeanoperator")
                parameter_operator.click()
                time.sleep(2)
                parameter_operator = safe_find_element(driver, By.ID, "pageBeanoperator")
                parameter_operator.send_keys(parameter_Operator_Value)
                time.sleep(2)
                parameter_operator_inputv = safe_find_element(driver, By.XPATH, f"//li[@title='{parameter_Operator_Value}']")
                parameter_operator_inputv.click()
                
                time.sleep(2)    
                parameter_value_input = safe_find_element(driver, By.ID, "pageBeanparameterValue")
                parameter_value_input.send_keys(parameter_Value)
                
                add_parameter = safe_find_element(driver, By.ID, "rulesaddParameter")
                add_parameter.click()
                # END ISSN RULE
                # Year Rule
                time.sleep(4)
                add_parameter = safe_find_element(driver, By.ID, "widgetId_Right_rulesaddParameter")
                add_parameter.click()
                
                parameter_name = safe_find_element(driver, By.ID, "pageBeanparameterName")
                parameter_name.send_keys(parameter_Date)
                time.sleep(2)
                parameter_name_select = safe_find_element(driver, By.XPATH, "//li[@role='option']")
                parameter_name_select.click()
                
                time.sleep(2)
                parameter_operator = safe_find_element(driver, By.ID, "pageBeanoperator")
                parameter_operator.click()
                time.sleep(2)
                parameter_operator = safe_find_element(driver, By.ID, "pageBeanoperator")
                parameter_operator.send_keys(parameter_Operator_Value)
                time.sleep(2)
                parameter_operator_inputv = safe_find_element(driver, By.XPATH, f"//li[@title='{parameter_Operator_Greaterthan}']")
                parameter_operator_inputv.click()
                
                time.sleep(4)
                parameter_value_input = safe_find_element(driver, By.ID, "pageBeanparameterValue")
                parameter_value_input.send_keys(year)
                
                add_parameter = safe_find_element(driver, By.ID, "rulesaddParameter")
                add_parameter.click()
                # END Year Rule
                
                # Update Output Parameters to True
                time.sleep(3)
                is_display = safe_find_element(driver, By.ID, "pageBeanoutputParameter")
                is_display.click()
                display = safe_find_element(driver, By.XPATH, "//li[@title='True']")
                display.click()
                
                # Save New rule
                time.sleep(2)
                save_rule = safe_find_element(driver, By.ID, "PAGE_BUTTONS_cbuttonsave")
                save_rule.click()
                current_parameter_Value = parameter_Value
                print(driver.page_source)
            except Exception as e:
                print(f"Thread={thread_id} Error adding new GES Rule: {e} and {driver.page_source}")
        if parameter_Value != current_parameter_Value:
            add_rule()
                
    time.sleep(20)
    driver.quit()
    
def main():
    global row_index
    ges_rule_data_sorted = ges_rule_data.sort_values(by=["Parameter_Value", "Parameter_Name", "Parameter_Operator", "Year", "Title"])
    if os.path.exists(OUTPUT_DIR) and os.listdir(OUTPUT_DIR):
        
        row_index, start_thread_id = retrieve_current_row_index(OUTPUT_DIR)
        print("previous analysis in progress.  start at row " + str(row_index) + " in the ges sheet\n")
    else:
        start_thread_id = 0
        
    threads = []
    
    for i in range(N):
        thread_id = start_thread_id + i
        print("thread id" + str(thread_id) + "\n")
        t = threading.Thread(target=worker_thread, args=(thread_id, ges_rule_data_sorted))
        t.start()
        threads.append(t)

    for t in threads:
        t.join()

    print("All threads complete.")

if __name__ == "__main__":
    main()
    