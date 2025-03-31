import pandas as pd
import numpy as np
import time
from selenium import webdriver 
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import time


# function to introduce a word in a string after a specific word contained in the original string
def introduce_string(original_string, specific_word, string_to_introduce):
    # Find the index of the specific word
    index = original_string.find(specific_word)
    if index != -1:
        # Insert the string to introduce at the found index
        modified_string = original_string[:index + len(specific_word)] + string_to_introduce + original_string[index + len(specific_word):]
        return modified_string


# function to remove a word in a string after a specific word contained in the original string
def remove_string(original_string, string_to_remove):
    # Find the index of the specific word
    index = original_string.find(string_to_remove)
    if index != -1:
        # Insert the string to introduce at the found index
        modified_string = original_string[:index] + original_string[index + len(string_to_remove):]
        return modified_string


# Function to remove thousand separator ,
def remove_comma(number):
    if ',' in number:
        number_without_comma = number.replace(",", "")
    else:
        number_without_comma = number
    return number_without_comma



# List 1
# List of type ship types options
# 'All Cargo Vessels' 
# 'Bulk Carrier'
# 'General Cargo'
# 'Container Ship'
# 'Reefer'
# 'Ro-Ro'
# 'Vehicles Carrier'
# 'Cement Carrier'
# 'All Tankers'
# 'Passenger/Cruise'


# defining objects we will use
dwt_list = []
imo_number_list = []

# Creating dictionary of type of ships and their respective link code
ship_type_dict = {'All Cargo Vessels' : 'type=4' , 'Bulk Carrier' : 'type=401', 'General Cargo' : 'type=402', 'Container Ship' : 'type=403', 'Reefer' : 'type=404', 'Ro-Ro' : 'type=405', 
'Vehicles Carrier' : 'type=406', 'Cement Carrier' : 'type=407', 'All Tankers' : 'type=6', 'Passenger/Cruise' : 'type=3'}

#Select type of ship
ship_type = 'Bulk Carrier' # Select type of ship here to create IMO list (exact as listed in list 1)
ship_type_code = ship_type_dict[ship_type]

# Minimum DWT that we are interested in
min_dwt = 5000

# Configure Chrome options for headless mode
chrome_options = Options()
chrome_options.add_argument("--headless")
# Initialize Chrome WebDriver
driver = webdriver.Chrome()
# Maximize the window
driver.maximize_window()


driver.get(f'https://www.vesselfinder.com/login')
username = 'jack@calcarea.com'
password = '1731Morada!'

print("AFTER DRIVER")
# Wait for the login page to load
wait = WebDriverWait(driver, 10)

iframes = driver.find_elements(By.TAG_NAME, 'iframe')
if iframes:
    driver.switch_to.frame(iframes[0])

# Find the username field and send the username
username_field = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div/form/div[1]/label/input')))
username_field.send_keys(username)

# Find the password field using XPath and send the password
password_field = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/form/div[2]/label/input")
password_field.send_keys(password)

# Find the login button and click it to log in
login_button = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/form/div[4]/input[3]")
login_button.click()

wait = WebDriverWait(driver, 10)

print("AFTER LOGIN")
driver.get(f'https://www.vesselfinder.com/vessels?{ship_type_code}')

# Sorting ships by DWT (in decreasing order)
DWT_sort_button = driver.find_element('xpath', '/html/body/div[1]/div/main/div/table/thead/tr/th[4]')
DWT_sort_button.click()

pages_number = int(remove_comma(driver.find_element('xpath', '/html/body/div[1]/div/main/div/nav[1]/div[2]/span').text.split('/')[1][1:])) # Getting number of pages to iterate on this later
initial_pages_number = pages_number
page_counter = 0
total_pages_counter = 0

try:
    # Scrapping information
    for page in range(initial_pages_number+1)[1:]:
        page_counter = page_counter + 1
        page_cp = page
        total_pages_counter += 1
        #print(page_counter)
        print(total_pages_counter)
        # Wait for an element to become clickable within 10 seconds
        element = WebDriverWait(driver, 0).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/main/div/table/tbody/tr[1]/td[1]/a")))  
        current_page_url = driver.current_url  # Getting current page url in case we need later
        first_item_dwt = float(driver.find_element('xpath', '/html/body/div[1]/div/main/div/table/tbody/tr[1]/td[4]').text) # getting dwt value of ship
        
        if first_item_dwt >= min_dwt:
            for i in range(21)[1:]: # (each page contains 20 ships)
                dwt_value = float(driver.find_element('xpath', f'/html/body/div[1]/div/main/div/table/tbody/tr[{i}]/td[4]').text) # getting dwt value of ship

                if dwt_value >= min_dwt:
                    imo_number = int(driver.find_element('xpath', f'/html/body/div[1]/div/main/div/table/tbody/tr[{i}]/td[1]/a').get_attribute('href')[-7:]) #getting IMO number of ship (7 digits)
                    i_cp = i # i checkpoint
                    
                    if imo_number not in imo_number_list:
                        imo_number_list.append(imo_number)
                        dwt_list.append(dwt_value)
                
                else:
                    break
            
        # This part is to avoid the code to get stuck
            items_switch = 1 # this parameter adjust how many items will be scrapped before switching to a new tab
            if page_counter % items_switch == 0 and page_counter != 200:
                # Open & Switch to new tab
                driver.switch_to.new_window('tab')
                # Closing previous tab
                driver.switch_to.window(driver.window_handles[0])
                driver.close()
                # Switching to new tab again
                driver.switch_to.window(driver.window_handles[0])            

            # Method writing page in link
            if page_counter != 200:
                if page_counter == 1:
                    next_page_url = introduce_string(current_page_url, '/vessels?', 'page=' + str(page_counter+1) + '&')
                    driver.get(next_page_url)
                else:
                    # removing current page string on link
                    next_page_url = remove_string(current_page_url, 'page=' + str(page_counter) + '&')
                    #introducing next page string on link
                    next_page_url = introduce_string(next_page_url, '/vessels?', 'page=' + str(page_counter+1) + '&')
                    driver.get(next_page_url)

            else: # This part is because the webpage just shows the first 200 pages, so we have to handle this
                last_ship_imo =  imo_number_list[-1]
                last_ship_dwt = dwt_list[-1]
                
                # Searching for last ship in page 200
                # Clearing box
                max_dwt_box = driver.find_element('xpath', '/html/body/div[1]/div/main/div/form/div[1]/div[2]/div[3]/input[2]')
                max_dwt_box.clear()
                # Writing last ship dwt searched
                max_dwt_box = driver.find_element('xpath', '/html/body/div[1]/div/main/div/form/div[1]/div[2]/div[3]/input[2]')
                max_dwt_box.send_keys(last_ship_dwt)
                search_button = driver.find_element('xpath', '/html/body/div[1]/div/main/div/form/div[2]/button[1]')
                search_button.click()

                # Sorting in decreasing order
                DWT_sort_button = driver.find_element('xpath', '/html/body/div[1]/div/main/div/table/thead/tr/th[4]')
                DWT_sort_button.click()
                
                # Getting new pages number
                pages_number = int(remove_comma(driver.find_element('xpath', '/html/body/div[1]/div/main/div/nav[1]/div[2]/span').text.split('/')[1][1:]))
                page_counter = 0 # Resetting page counter

                current_url = driver.current_url
                # This part is to avoid the code to get stuck
                items_switch = 1 # this parameter adjust how many items will be scrapped before switching to a new tab
                if page_counter % items_switch == 0:
                    # Open & Switch to new tab
                    driver.switch_to.new_window('tab')
                    # Closing previous tab
                    driver.switch_to.window(driver.window_handles[0])
                    driver.close()
                    # Switching to new tab again
                    driver.switch_to.window(driver.window_handles[0])
                    # Opening last page in new tab
                    driver.get(current_url)

        else: 
            break
        
except TimeoutException:
        # If the timeout occurs, take appropriate action (e.g., print an error message)
        print("Timeout occurred. The element was not clickable within the specified time frame.")



# Data Frame
imo_df = pd.DataFrame({'IMO #' : imo_number_list, 'DWT' : dwt_list})

# Get the current date and time
current_date_time = datetime.now()
# Convert the date and time to a string
date_string = current_date_time.strftime("%Y%m%d")

# Saving file
imo_df.to_excel(f'C:\\Users\\jacks\\OneDrive\\Documents\\CalcareaCode\\imo_list_{ship_type}_{date_string}.xlsx')