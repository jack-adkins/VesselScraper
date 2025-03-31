import pandas as pd
import numpy as np
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
import time
import shutil
from datetime import datetime
import random





def find_duplicate_indices(input_list):
    value_indices = {}
    duplicate_indices = []

    for index, value in enumerate(input_list):
        if value in value_indices:
            duplicate_indices.append(index)
        else:
            value_indices[value] = index
    
    return duplicate_indices





ship_type = ''
focus = 'GreatEastern'
current_date_time = datetime.now()
# date_string = current_date_time.strftime("%Y%m%d")
date_string = '20240611'
#  Reading imo list we want
# file_name = 'imo_list_{focus}_{date_string}.xlsx' #Introduce here the name of the file we want to read as a string
fleet_imo = pd.read_excel(f'C:\\Users\\jacks\\OneDrive\\Documents\\CalcareaCode\\imo_list_{focus}_{date_string}.xlsx')  




status_list = []
port_origin_list = []
port_destination_list = []
country_origin_list = []
country_destination_list = []
imo_number_list = []
vessel_name_list = []
flag_list = []
speed_list = []
ship_type_list = []
dwt_list = []
length_list = []
beam_list = []
year_list = []
draft_list = []
teu_list = []
ship_owner_list = []
size_category_list = []

# Set Chrome options to run in headless mode
chrome_options = Options()
chrome_options.add_argument('--headless')  # This line makes Chrome run in headless mode
driver = webdriver.Chrome()
driver.maximize_window()
#driver = webdriver.Chrome(options=chrome_options)

items_switch = 1

print('RIGHT BEFORE LOOP')

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

for index, name in enumerate(list(fleet_imo['Name'])):
    print("START OF LOOP")
    item = index + 1
    print(item) # This is to know in which item is the scrapping program

    driver.get(f'https://www.vesselfinder.com/pro/map')

    wait = WebDriverWait(driver, 10)

    iframes = driver.find_elements(By.TAG_NAME, 'iframe')
    if iframes:
        driver.switch_to.frame(iframes[0])

    # Find the username field and send the username
    search_bar = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div/div[1]/div[3]/div[1]/div[1]/input')))
    search_bar.send_keys(name)

    driver.get(f'https://www.vesselfinder.com/pro/map#vessel-details?imo={ship_imo}&mmsi=0')

    wait = WebDriverWait(driver, 1)
    
    try:
        status = ''
        try:
            status = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[1]/td[2]/span').text
            status = 'Not in Service'
            print("NOT IN SERVICE")
        except:
            # If the element is not found, 'status' will remain an empty string
            print("Element not found. Status remains empty.")
        # except Exception as e:
        #     # Catch any other exceptions that might occur
        #     print(f"An error occurred: {str(e)}")
                                                                 
        if 'Not in Service' not in status:
            print("IS THIS THE ISSUE????")
            try:
                # Wait for staple info to become visible within 10 seconds
                # wait = WebDriverWait(driver, 1)
                element = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div/div[4]/div/div/div[1]/div/div[2]/h1')))
                print("ON THE RIGHT PAGE")
                # Scraping Origin
                try:
                    origin = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div/div[4]/div/div/div[3]/div[2]/div[3]/div[2]/div[1]/span')))
                    # origin = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[3]/div[2]/div[3]/div[2]/div[1]/span').text
                #except NoSuchElementException as e:
                #    origin = driver.find_element('xpath', '/html/body/div[1]/div/main/div/section[2]/div/div[2]/div/div[3]/div[2]/div[1]').text
                except NoSuchElementException as e:
                    origin = ''
                except:
                    origin = ''
                if not isinstance(origin, str):
                    origin = origin.text  # Extract text if it's a WebElement
                port_origin = origin.split(',')[0]
                try:
                    country_origin = origin.split(',')[1][1:]
                except IndexError as e:
                    country_origin = ""
                except:
                    country_origin = ""
                print("ORIGIN", port_origin)
                # Scraping destination
                try:
                    destination = driver.find_element('xpath',"/html/body/div/div/div[4]/div/div/div[3]/div[2]/div[1]/div[2]/div[1]/span").text 
                    port_destination = destination.split(',')[0]
                    try:
                        country_destination = destination.split(',')[1][1:]
                    except IndexError as e:
                        country_destination = ''
                    except:
                        country_destination = ''
                # except NoSuchElementException as e:
                #     try:
                #         port_destination = driver.find_element('xpath',"/html/body/div/div/div[4]/div/div/div[3]/div[2]/div[1]/div[2]/div[1]/span/text()[1]").text
                #         country_destination = driver.find_element('xpath', "/html/body/div/div/div[4]/div/div/div[3]/div[2]/div[1]/div[2]/div[1]/span/text()[2]").text
                #         country_destination = country_destination.split(',')[1][1:]
                #     except:
                #         destination = ''
                #         port_destination = ''
                #         country_destination = ''
                except:
                    destination = ''
                    port_destination = ''
                    country_destination = ''


                print("DESTINATION", port_destination)

                # Scrapping vessel name
                try:
                    vessel_name = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[1]/div/div[2]/h1').text
                except:
                    vessel_name = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[2]/td[2]').text 
                print("NAME", vessel_name)
                # Scrapping ship type
                # ship_type = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[3]/td[2]').text

                # Scrapping flag
                # flag = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[4]/td[2]').text
                flag = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[4]/td[2]'))).text
                print("FLAG", flag)

                # Scrapping speed
                speed = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[5]/div/div[2]/div/div[1]/div[1]/table/tbody/tr[3]/td[2]').text
                speed = speed.split(' /')[0]
                print("Speed", speed)
                if (('-' not in speed) and (speed != '')):
                    speed = float(speed)


                # Scrapping DWT
                dwt = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[7]/td[2]').text
                dwt = int(dwt)

                # Scrapping length
                length = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[8]/td[2]').text
                length = int(length)

                # Scrapping beam
                beam = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[9]/td[2]').text
                beam = int(beam)

                # Scrapping year
                year = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[11]/td[2]').text
                year = int(year)

                #scraping draft
                draft = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[10]/td[2]').text

                if (draft == '-'):
                    draft = ''
                else:
                    draft = float(draft)

                teu = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[15]/td[2]').text
                if (teu == '-'):
                    teu = ''
                else:
                    teu = int(teu)
                
                status_list.append('In Service')
                port_origin_list.append(port_origin)
                port_destination_list.append(port_destination)
                country_origin_list.append(country_origin)
                country_destination_list.append(country_destination)
                imo_number_list.append(ship_imo)
                vessel_name_list.append(vessel_name)
                ship_type_list.append(ship_type)
                flag_list.append(flag)
                speed_list.append(speed)
                dwt_list.append(dwt)
                length_list.append(length)
                beam_list.append(beam)
                year_list.append(year)
                draft_list.append(draft)
                teu_list.append(teu)

                # This part is to avoid the code to get stuck
                items_switch = 1 # this parameter adjust how many items will be scrapped before switching to a new tab
                if item % items_switch == 0:
                    # Open & Switch to new tab
                    driver.switch_to.new_window('tab')
                    # Closing previous tab
                    driver.switch_to.window(driver.window_handles[0])
                    driver.close()
                    # Switching to new tab again
                    driver.switch_to.window(driver.window_handles[0])

            except TimeoutException or NoSuchElementException:
                # If the timeout occurs, take appropriate action (e.g., print an error message)
                print("Timeout occurred. The element was not clickable within the specified time frame.")
                status_list.append('NA')
                port_origin_list.append('NA')
                port_destination_list.append('NA')
                country_origin_list.append('NA')
                country_destination_list.append('NA')
                imo_number_list.append(ship_imo)
                vessel_name_list.append('NA')
                ship_type_list.append('NA')
                flag_list.append('NA')
                speed_list.append('NA')
                dwt_list.append('NA')
                length_list.append('NA')
                beam_list.append('NA')
                year_list.append('NA')
                draft_list.append('NA')
                teu_list.append('NA')

                # This part is to avoid the code to get stuck
                if item % items_switch == 0:
                    # Switch and open new tab
                    driver.switch_to.new_window('tab')
                    # Closing previous tab
                    driver.switch_to.window(driver.window_handles[0])
                    driver.close()
                    # Switching to new tab again
                    driver.switch_to.window(driver.window_handles[0])

        else: # is ship not in service
            status_list.append(status)
            port_origin_list.append('NS')
            port_destination_list.append('NS')
            country_origin_list.append('NS')
            country_destination_list.append('NS')
            imo_number_list.append(ship_imo)
            vessel_name_list.append('NS')
            ship_type_list.append('NS')
            flag_list.append('NS')
            speed_list.append('NS')
            dwt_list.append('NS')
            length_list.append('NS')
            beam_list.append('NS')
            year_list.append('NS')
            draft_list.append('NS')
            teu_list.append('NS')

        # This part is to avoid the code to get stuck
            if item % items_switch == 0:
                # Switch and open new tab
                driver.switch_to.new_window('tab')
                # Closing previous tab
                driver.switch_to.window(driver.window_handles[0])
                driver.close()
                # Switching to new tab again
                driver.switch_to.window(driver.window_handles[0])
    except TimeoutException or NoSuchElementException:
                # If the timeout occurs, take appropriate action (e.g., print an error message)
                print("Timeout occurred. The element was not clickable within the specified time frame.")
                status_list.append('NA')
                port_origin_list.append('NA')
                port_destination_list.append('NA')
                country_origin_list.append('NA')
                country_destination_list.append('NA')
                imo_number_list.append(ship_imo)
                vessel_name_list.append('NA')
                ship_type_list.append('NA')
                flag_list.append('NA')
                speed_list.append('NA')
                dwt_list.append('NA')
                length_list.append('NA')
                beam_list.append('NA')
                year_list.append('NA')
                draft_list.append('NA')
                teu_list.append('NA')

                # This part is to avoid the code to get stuck
                if item % items_switch == 0:
                    # Switch and open new tab
                    driver.switch_to.new_window('tab')
                    # Closing previous tab
                    driver.switch_to.window(driver.window_handles[0])
                    driver.close()
                    # Switching to new tab again
                    driver.switch_to.window(driver.window_handles[0])
    print("END OF LOOP")
print("COMPLETED LOOPING!!!!!!!")
fleet_imo['Origin Port'] = port_origin_list
fleet_imo['Origin Country'] = country_origin_list
fleet_imo['Destination Port'] = port_destination_list
fleet_imo['Destination Country'] = country_destination_list



# Ship owner feature
ship_owner_list = []
# Configure Chrome options for headless mode
chrome_options = Options()
chrome_options.add_argument("--headless")
# Initialize Chrome WebDriver
driver = webdriver.Chrome()
# Maximize the window
driver.maximize_window()

for ship_imo in list(fleet_imo['IMO #']):
    driver.get(f'https://www.balticshipping.com/vessel/imo/{ship_imo}')
    try:
    # Wait for an element to become clickable within 10 seconds
        element = WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div[2]/article/section/div[1]/div/div/table/tbody/tr[1]/td")))
                                         
        try:
            for i in range(17)[1:]: #the owner features is usually in the number 16
                try:
                    ship_owner_box = driver.find_element("xpath",f'/html/body/div[5]/div[2]/article/section/div[1]/div/div[3]/table/tbody/tr[{i}]/th')
                except NoSuchElementException:
                    ship_owner_box = driver.find_element("xpath",f'/html/body/div[5]/div[2]/article/section/div[1]/div/div/table/tbody/tr[{i}]/th')
                if ship_owner_box.text == 'Owner':
                    try:
                        ship_owner = driver.find_element("xpath",f'/html/body/div[5]/div[2]/article/section/div[1]/div/div[3]/table/tbody/tr[{i}]/td').text
                    except NoSuchElementException:
                        ship_owner = driver.find_element("xpath",f'/html/body/div[5]/div[2]/article/section/div[1]/div/div/table/tbody/tr[{i}]/td').text
                else:
                    ship_owner = 'Not Found'
        
        except NoSuchElementException:
                ship_owner = 'NA'
        ship_owner_list.append(ship_owner)
       

    except TimeoutException or NoSuchElementException:
        ship_owner = 'Not Found'
        ship_owner_list.append(ship_owner)


#Categorizing into ship type by size feature
size_category_list = categorize_by_size(ship_type_list, dwt_list, length_list, draft_list, teu_list, size_category_list)


# European Union feature
EU_countries = ['Austria', 'Belgium', 'Bulgaria', 'Croatia', 'Cyprus', 'Czechia', 'Denmark', 'Estonia', 'Finland', 'France', 'Germany', 'Greece', 'Hungary', 'Ireland', 'Italy', 'Latvia', 'Lithuania'
                , 'Luxembourg', 'Malta', 'Netherlands', 'Poland', 'Portugal', 'Romania', 'Slovakia', 'Slovenia', 'Spain', 'Sweden']

fleet_imo['Origin EU?'] = np.NaN
fleet_imo['Destination EU?'] = np.NaN

for i in range(len(fleet_imo)):
    if fleet_imo['Origin Country'][i] in EU_countries:
        fleet_imo['Origin EU?'][i] = 'Yes'
    else:
        fleet_imo['Origin EU?'][i] = 'No'

    if fleet_imo['Destination Country'][i] in EU_countries:
        fleet_imo['Destination EU?'][i] = 'Yes'
    else:
        fleet_imo['Destination EU?'][i] = 'No'



fleet_imo['Vessel Name'] = vessel_name_list
fleet_imo['Status'] = status_list
fleet_imo['Origin Port - Destination Port'] = fleet_imo['Origin Port'] + ' - ' + fleet_imo['Destination Port']
fleet_imo['Origin Country - Destination Country'] = fleet_imo['Origin Country'] + ' - ' + fleet_imo['Destination Country']
fleet_imo['Flag'] = flag_list
fleet_imo['Speed'] = speed_list
fleet_imo['Ship type'] = ship_type_list
fleet_imo['Size Category'] = size_category_list
fleet_imo['DWT'] = dwt_list
fleet_imo['Length'] = length_list
fleet_imo['Beam'] = beam_list
fleet_imo['Year'] = year_list
fleet_imo['Draft'] = draft_list
fleet_imo['TEU'] = teu_list
fleet_imo['Ship Owner'] = ship_owner_list

colums_order = ['IMO #', 'Vessel Name', 'Origin Port', 'Origin Country', 'Origin EU?', 'Destination Port', 'Destination Country', 'Destination EU?', 'Origin Port - Destination Port', 'Origin Country - Destination Country', 'Country Route', 
                'Flag', 'Speed', 'Ship type', 'Size Category', 'DWT', 'Length', 'Beam', 'Draft', 'TEU', 'Year', 'Ship Owner']
# Columns reordered
fleet_imo = fleet_imo.reindex(columns=colums_order)



# Saving file
# Get the current date and time
current_date_time = datetime.now()
# Convert the date and time to a string
date_string = current_date_time.strftime("%Y%m%d")
fleet_imo.to_excel(f'C:\\Users\\jacks\\OneDrive\\Documents\\CalcareaCode\\fleet_list_{focus}_{date_string}.xlsx')