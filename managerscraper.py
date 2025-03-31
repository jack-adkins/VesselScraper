import pandas as pd
import numpy as np
import time
from selenium import webdriver 
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import time

# Function to introduce a word in a string after a specific word contained in the original string
def introduce_string(original_string, specific_word, string_to_introduce):
    index = original_string.find(specific_word)
    if index != -1:
        modified_string = original_string[:index + len(specific_word)] + string_to_introduce + original_string[index + len(specific_word):]
        return modified_string

# Function to remove a word in a string after a specific word contained in the original string
def remove_string(original_string, string_to_remove):
    index = original_string.find(string_to_remove)
    if index != -1:
        modified_string = original_string[:index] + original_string[index + len(string_to_remove):]
        return modified_string

# Function to remove thousand separator ","
def remove_comma(number):
    if ',' in number:
        number_without_comma = number.replace(",", "")
    else:
        number_without_comma = number
    return number_without_comma

# Function to login to VesselFinder with premium account
def login_vesselfinder(driver, username, password):
    driver.get("https://www.vesselfinder.com/users/login")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "email"))).send_keys(username)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "password"))).send_keys(password)
    login_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//button[@type="submit"]')))
    login_button.click()
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "account-menu")))  # Ensure login is successful

# Function to check manager of the ship
def check_manager(driver):
    try:
        manager_info = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//td[contains(text(), "Manager:")]/following-sibling::td'))).text
        return manager_info in ['LOMAR SHIPPING LTD']  # Update with actual manager names
    except:
        return False

# Continue with the initialization and login process
chrome_options = Options()
chrome_options.add_argument("--headless")
driver = webdriver.Chrome(options=chrome_options)
login_vesselfinder(driver, 'jack@calcarea.com', '1731Morada!')

# Setting up the scraping process
ship_type = 'Bulk Carrier'  # Define the ship type
ship_type_code = ship_type_dict[ship_type]
min_dwt = 5000
driver.get(f'https://www.vesselfinder.com/vessels?{ship_type_code}')
DWT_sort_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/main/div/table/thead/tr/th[4]')))
DWT_sort_button.click()

# Starting the scraping loop
pages_number = int(remove_comma(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/main/div/nav[1]/div[2]/span')).text.split('/')[1][1:])))
initial_pages_number = pages_number
page_counter = 0
total_pages_counter = 0

try:
    for page in range(initial_pages_number+1)[1:]:
        page_counter += 1
        total_pages_counter += 1
        print(total_pages_counter)
        element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/main/div/table/tbody/tr[1]/td[1]/a")))
        current_page_url = driver.current_url
        first_item_dwt = float(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/main/div/table/tbody/tr[1]/td[4]')).text))

        if first_item_dwt >= min_dwt:
            for i in range(1, 21):  # Assuming 20 ships per page
                dwt_value = float(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f'/html/body/div[1]/div/main/div/table/tbody/tr[{i}]/td[4]')).text))
                if dwt_value >= min_dwt:
                    driver.find_element(By.XPATH, f'/html/body/div[1]/div/main/div/table/tbody/tr[{i}]/td[1]/a').click()
                    if check_manager(driver):  # Check manager here
                        imo_number = int(driver.current_url.split('/')[-1])
                        if imo_number not in imo_number_list:
                            imo_number_list.append(imo_number)
                            dwt_list.append(dwt_value)
                    driver.back()
                else:
                    break
        else:
            break

        # Handle pagination and new tabs
        if page_counter % items_switch == 0 and page_counter != 200:
            driver.switch_to.new_window('tab')
            driver.switch_to.window(driver.window_handles[0])
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
            login_vesselfinder(driver, 'jack@calcarea.com', '1731Morada!')  # Re-login after switching tabs

        if page_counter != 200:
            if page_counter == 1:
                next_page_url = introduce_string(current_page_url, '/vessels?', 'page=' + str(page_counter+1) + '&')
            else:
                next_page_url = remove_string(current_page_url, 'page=' + str(page_counter) + '&')
                next_page_url = introduce_string(next_page_url, '/vessels?', 'page=' + str(page_counter+1) + '&')
            driver.get(next_page_url)
except TimeoutException:
    print("Timeout occurred. The element was not clickable within the specified time frame.")

# Data frame creation and file saving
imo_df = pd.DataFrame({'IMO #': imo_number_list, 'DWT': dwt_list})
date_string = datetime.now().strftime("%Y%m%d")
imo_df.to_excel(f'C:\\Users\\jacks\\OneDrive\\Documents\\CalcareaCode\\imo_list_manager_{ship_type}_{date_string}.xlsx')

driver.quit()
