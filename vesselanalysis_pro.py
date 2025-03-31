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



def categorize_by_size(ship_type_list, dwt_list, length_list, draft_list, teu_list, size_category_list):
    count = 0
    for stype in ship_type_list:
        if ('carrier' in stype.lower()):
            if (dwt_list[count] < 20000):
                size_category_list.append('Mini') 
            elif(dwt_list[count] <= 35000 and length_list[count] <= 150):
                size_category_list.append('Handysize') 
            elif(dwt_list[count] < 50000):
                size_category_list.append('Handymax') 
            elif(dwt_list[count] <= 61000):
                size_category_list.append('Supramax') 
            elif(dwt_list[count] <= 80000 and length_list[count] <= 290):
                size_category_list.append('Panamax') 
            elif(dwt_list[count] > 180000 and length_list[count] <= 300):
                size_category_list.append('Newcastlemax')
            elif(length_list[count] <= 226):
                size_category_list.append('Seawaymax')
            elif(dwt_list[count] > 80000):
                size_category_list.append('Capesize')
            else:
                size_category_list.append('NA')
        elif('tanker' in stype.lower()):
            if (dwt_list[count] <= 10000):
                size_category_list.append('General Purpose')
            elif(dwt_list[count] <= 50000):
                size_category_list.append('Handysize')
            elif(dwt_list[count] <= 80000):
                size_category_list.append('Panamax')
            elif(dwt_list[count] <= 120000):
                size_category_list.append('Aframax')
            elif(dwt_list[count] <= 200000):
                size_category_list.append('Suezmax')
            elif(dwt_list[count] <= 320000):
                size_category_list.append('VLCC') 
            else:
                size_category_list.append('ULCC')
        elif('container' in stype.lower()):
            if (teu_list[count] < 1000):
                size_category_list.append('Feeder')
            elif(teu_list[count] <= 3000):
                size_category_list.append('Feedermax')
            elif(teu_list[count] <= 5100):
                size_category_list.append('Panamax')
            elif(teu_list[count] <= 10000):
                size_category_list.append('Post-Panamax')
            elif(teu_list[count] <= 14500):
                size_category_list.append('New Panamax')
            else:
                size_category_list.append('ULCV')
        else:
            size_category_list.append('NA')
        count += 1
    return size_category_list







ship_type = ''
focus = 'BergeBulk'
current_date_time = datetime.now()
# date_string = current_date_time.strftime("%Y%m%d")
date_string = '20240628'
#  Reading imo list we want
# file_name = 'imo_list_{focus}_{date_string}.xlsx' #Introduce here the name of the file we want to read as a string
fleet_imo = pd.read_excel(f'C:\\Users\\jacks\\OneDrive\\Documents\\CalcareaCode\\IMO_Lists\\imo_list_{focus}_{date_string}.xlsx')
voyages = pd.read_excel(f'C:\\Users\\jacks\\OneDrive\\Documents\\CalcareaCode\\IMO_Lists\\imo_list_{focus}_{date_string}.xlsx')
history = pd.DataFrame(columns=['Port', 'Country', 'Visits', 'EU?'])  


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
ship_manager_list = []
size_category_list = []
port_calls1 = []
port_country_calls1 = []
port_calls2 = []
port_country_calls2 = []
port_calls3 = []
port_country_calls3 = []
port_calls4 = []
port_country_calls4 = []
port_calls5 = []
port_country_calls5 = []
EU_counter = []
history_name = []
history_count = []

# Set Chrome options to run in headless mode
chrome_options = Options()
chrome_options.add_argument('--headless')  # This line makes Chrome run in headless mode
driver = webdriver.Chrome()
driver.maximize_window()
#driver = webdriver.Chrome(options=chrome_options)
items_switch = 1





#LOGIN 
print('RIGHT BEFORE LOGIN')

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

elem = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div/div/div[1]/nav/ul/li[1]/a')))








#MAIN SCRAPING LOOP
for index, ship_imo in enumerate(list(fleet_imo['IMO #'])):
    print("START OF LOOP")
    item = index + 1
    print(item) # This is to know in which item is the scraping program

    driver.get(f'https://www.vesselfinder.com/pro/map#vessel-details?imo={ship_imo}&mmsi=0')

    wait = WebDriverWait(driver, 1)
    try:
        status = ''
        #Filtering out ships that are not in service
        try:
            elem = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[1]/td[1]')))
            status = elem.text
            print('STATUS 1:', status)
            if (status == 'Status'):
                status = 'Not in Service'
                print("NOT IN SERVICE")
        except:
            # If the element is not found, 'status' will remain an empty string
            print("Element 1 not found. Status remains empty.")
            try:
                elem = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div/div/div[4]/div/div/div[4]/div/table/tbody/tr[1]/td[1]')))
                status = elem.text
                print('STATUS 2:', status)
                if (status == 'Status'):
                    status = 'Not in Service'
                    print("NOT IN SERVICE")
            except:
                print("Element 2 not found. Status remains empty.")

        #Collecting the data
        if 'Not in Service' not in status:
            try:
                # Wait for staple info to become visible on the page
                element = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div/div[4]/div/div/div[1]/div/div[2]/h1')))
                print("ON THE RIGHT PAGE")

                # Scraping Origin
                try:
                    origin = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div/div[4]/div/div/div[3]/div[2]/div[3]/div[2]/div[1]/span')))
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
                except:
                    destination = ''
                    port_destination = ''
                    country_destination = ''


                print("DESTINATION", port_destination)

                # Scraping vessel name
                try:
                    vessel_name = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[1]/div/div[2]/h1').text
                except:
                    vessel_name = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[2]/td[2]').text 
                print("NAME", vessel_name)

                # Scraping ship type
                try:
                    ship_type = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[3]/td[2]').text
                except:
                    ship_type = '-'

                # Scraping flag
                # flag = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[4]/td[2]').text
                flag = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[4]/td[2]'))).text
                print("FLAG", flag)

                # Scraping speed
                speed = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[5]/div/div[2]/div/div[1]/div[1]/table/tbody/tr[3]/td[2]').text
                speed = speed.split(' /')[0]
                print("Speed", speed)
                if (('-' not in speed) and (speed != '')):
                    speed = float(speed)


                # Scraping DWT
                dwt = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[7]/td[2]').text
                try:
                    dwt = int(dwt)
                except:
                    dwt = None

                # Scraping length
                length = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[8]/td[2]').text
                try:
                    length = int(length)
                except:
                    length = None

                # Scraping beam
                beam = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[9]/td[2]').text
                try:
                    beam = int(beam)
                except:
                    beam = None

                # Scraping year
                year = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[11]/td[2]').text
                try:
                    year = int(year)
                except:
                    year = None

                #scraping draft
                draft = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[10]/td[2]').text

                if (draft == '-'):
                    draft = ''
                else:
                    try:
                        draft = float(draft)
                    except:
                        draft = None

                #Scraping TEU
                teu = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[15]/td[2]').text
                if (teu == '-'):
                    teu = ''
                else:
                    try:
                        teu = int(teu)
                    except:
                        teu = None

                #Scraping manager and owner
                ship_owner = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[21]/td[2]').text
                print('OWNER:', ship_owner)

                ship_manager = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[6]/div/table/tbody/tr[25]/td[2]').text
                print('MANAGER:', ship_manager)

                #GETTING UP TO THE LAST 5 PORT CALLS SCRAPING
                try:
                    full_port = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[4]/div[2]/div/div[1]/a').text
                    port1 = full_port.split(',')[0]
                    try:
                        country_call1 = full_port.split(',')[1][1:]
                    except:
                        country_call1 = ''
                except:
                    port1 = ''
                    country_call1 = ''

                try:
                    full_port = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[4]/div[2]/div/div[2]/a').text
                    port2 = full_port.split(',')[0]
                    try:
                        country_call2 = full_port.split(',')[1][1:]
                    except:
                        country_call2 = ''
                except:
                    port2 = ''
                    country_call2 = ''

                try:
                    full_port = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[4]/div[2]/div/div[3]/a').text
                    port3 = full_port.split(',')[0]
                    try:
                        country_call3 = full_port.split(',')[1][1:]
                    except:
                        country_call3 = ''
                except:
                    port3 = ''
                    country_call3 = ''
                
                try:
                    full_port = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[4]/div[2]/div/div[4]/a').text
                    port4 = full_port.split(',')[0]
                    try:
                        country_call4 = full_port.split(',')[1][1:]
                    except:
                        country_call4 = ''
                except:
                    port4 = ''
                    country_call4 = ''

                try:
                    full_port = driver.find_element('xpath', '/html/body/div/div/div[4]/div/div/div[4]/div[2]/div/div[5]/a').text
                    port5 = full_port.split(',')[0]
                    try:
                        country_call5 = full_port.split(',')[1][1:]
                    except:
                        country_call5 = ''
                except:
                    port5 = ''
                    country_call5 = ''
                print('PORT5:', port5)


                # FOR GETTING PORT CALLS ACROSS A FLEET FROM 2023
                for i in range(1,7):
                    try:
                        lookup_xpath1 = f'/html/body/div/div/div[4]/div/div/div[5]/div/div[2]/div/div[2]/div[1]/div[2]/div/div/div[{i}]/a'
                        lookup_xpath2 = f'/html/body/div/div/div[4]/div/div/div[5]/div/div[2]/div/div[2]/div[1]/div[2]/div/div/div[{i}]/div'
                        lookup_xpathflag = f'/html/body/div/div/div[4]/div/div/div[5]/div/div[2]/div/div[2]/div[1]/div[2]/div/div/div[{i}]/a/span'
                        grabber = wait.until(EC.visibility_of_element_located((By.XPATH, lookup_xpath1)))
                        port_name = grabber.text
                        grabber = wait.until(EC.visibility_of_element_located((By.XPATH, lookup_xpath2)))
                        port_visits = grabber.text
                        grabber = wait.until(EC.visibility_of_element_located((By.XPATH, lookup_xpathflag)))
                        port_flag = grabber.get_attribute('title')
                        try:
                            port_visits = int(port_visits)
                        except:
                            port_visits = 0
                    except:
                        port_name = ''
                        port_visits = 0

                    if port_name:  # Ensure the port name is not empty
                        if port_name in history['Port'].values:
                            history.loc[history['Port'] == port_name, 'Visits'] += port_visits
                        else:
                            new_entry = pd.DataFrame([[port_name, port_flag, port_visits]], columns=['Port', 'Country', 'Visits'])
                            history = pd.concat([history, new_entry], ignore_index=True)

                
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
                ship_owner_list.append(ship_owner)
                ship_manager_list.append(ship_manager)
                port_calls1.append(port1)
                port_country_calls1.append(country_call1)
                port_calls2.append(port2)
                port_country_calls2.append(country_call2)
                port_calls3.append(port3)
                port_country_calls3.append(country_call3)
                port_calls4.append(port4)
                port_country_calls4.append(country_call4)
                port_calls5.append(port5)
                port_country_calls5.append(country_call5)

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
                ship_owner_list.append('NA')
                ship_manager_list.append('NA')
                port_calls1.append('NA')
                port_country_calls1.append('NA')
                port_calls2.append('NA')
                port_country_calls2.append('NA')
                port_calls3.append('NA')
                port_country_calls3.append('NA')
                port_calls4.append('NA')
                port_country_calls4.append('NA')
                port_calls5.append('NA')
                port_country_calls5.append('NA')

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
            ship_owner_list.append('NS')
            ship_manager_list.append('NS')
            port_calls1.append('NS')
            port_country_calls1.append('NS')
            port_calls2.append('NS')
            port_country_calls2.append('NS')
            port_calls3.append('NS')
            port_country_calls3.append('NS')
            port_calls4.append('NS')
            port_country_calls4.append('NS')
            port_calls5.append('NS')
            port_country_calls5.append('NS')


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
                ship_owner_list.append('NA')
                ship_manager_list.append('NA')
                port_calls1.append('NA')
                port_country_calls1.append('NA')
                port_calls2.append('NA')
                port_country_calls2.append('NA')
                port_calls3.append('NA')
                port_country_calls3.append('NA')
                port_calls4.append('NA')
                port_country_calls4.append('NA')
                port_calls5.append('NA')
                port_country_calls5.append('NA')

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
fleet_imo['Ship Manager'] = ship_manager_list

colums_order = ['IMO #', 'Vessel Name', 'Origin Port', 'Origin Country', 'Origin EU?', 'Destination Port', 'Destination Country', 'Destination EU?', 'Origin Port - Destination Port', 'Origin Country - Destination Country', 'Country Route', 
                'Flag', 'Speed', 'Ship type', 'Size Category', 'DWT', 'Length', 'Beam', 'Draft', 'TEU', 'Year', 'Ship Owner', 'Ship Manager']
# Columns reordered
fleet_imo = fleet_imo.reindex(columns=colums_order)

print("GOT HERE")
#Creating the Ports Calls Sheet
voyages['Port 1'] = port_calls1
voyages['Port 2'] = port_calls2
voyages['Port 3'] = port_calls3
voyages['Port 4'] = port_calls4
voyages['Port 5'] = port_calls5
voyages['Country 1'] = port_country_calls1
voyages['Country 2'] = port_country_calls2
voyages['Country 3'] = port_country_calls3
voyages['Country 4'] = port_country_calls4
voyages['Country 5'] = port_country_calls5
voyages['EU Visits'] = 0
voyages['Vessel Name'] = vessel_name_list
voyages['Ship Owner'] = ship_owner_list
voyages['Ship Manager'] = ship_manager_list

# In the EU? Checks
for i in range(len(voyages)):
    if voyages['Country 1'][i] in EU_countries:
        voyages['EU Visits'][i] += 1
    if voyages['Country 2'][i] in EU_countries:
        voyages['EU Visits'][i] += 1
    if voyages['Country 3'][i] in EU_countries:
        voyages['EU Visits'][i] += 1
    if voyages['Country 4'][i] in EU_countries:
        voyages['EU Visits'][i] += 1
    if voyages['Country 5'][i] in EU_countries:
        voyages['EU Visits'][i] += 1
for i in range(len(history)):
    if history['Country'][i] in EU_countries:
        history['EU?'][i] = 'Yes'
    else:
        history['EU?'][i] = 'No'
    

print("ALSO GOT HERE")
columns_order_2 = ['Port 1', 'Country 1', 'Port 2', 'Country 2', 'Port 3', 'Country 3', 'Port 4', 'Country 4', 'Port 5', 'Country 5', 'EU Visits', 'Vessel Name', 'Ship Owner', 'Ship Manager']

print("AND HERE")
voyages_df = voyages.reindex(columns=columns_order_2)

# Get the current date and time
current_date_time = datetime.now()
# Convert the date and time to a string
date_string = current_date_time.strftime("%Y%m%d")

#Saving file
voyages_df.to_excel(f'C:\\Users\\jacks\\OneDrive\\Documents\\CalcareaCode\\Fleet_Lists\\voyages_{focus}_{date_string}.xlsx')

# Saving file
fleet_imo.to_excel(f'C:\\Users\\jacks\\OneDrive\\Documents\\CalcareaCode\\Fleet_Lists\\fleet_list_{focus}_{date_string}.xlsx')

#Saving file
history.to_excel(f'C:\\Users\\jacks\\OneDrive\\Documents\\CalcareaCode\\Fleet_Lists\\port_calls_{focus}_{date_string}.xlsx')