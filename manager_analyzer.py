import pandas as pd
import numpy as np
from selenium import webdriver 
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import WebDriverException, InvalidSessionIdException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import shutil
from datetime import datetime
import random
import winsound
import threading

timer_cancel_flag = False

def is_driver_active(driver):
    try:
        driver.current_url  # Access the current URL to check if the driver is active
        return True
    except (WebDriverException):
        return False
    except:
        return False

def play_sound():
    # Replace with the desired sound file or use winsound.Beep on Windows
    winsound.Beep(1000, 500)  # Frequency 1000 Hz, Duration 500 ms

def timer(cancel_flag):
    time.sleep(20)  # Wait for 20 seconds
    if not cancel_flag.is_set():
        play_sound()

ship_type = 'Bulk Carrier_80k-180k'
date_string = '20240628'
#  Reading imo list we want
fleet_imo = pd.read_excel(f'C:\\Users\\jacks\\OneDrive\\Documents\\CalcareaCode\\IMO_Lists\\imo_list_{ship_type}_{date_string}.xlsx')  

imo_number_list = []
ship_owner_list = []
mmsi_list = []

EU_countries = ['Austria', 'Belgium', 'Bulgaria', 'Croatia', 'Cyprus', 'Czechia', 'Denmark', 'Estonia', 'Finland', 'France', 'Germany', 'Greece', 'Hungary', 'Ireland', 'Italy', 'Latvia', 'Lithuania'
                , 'Luxembourg', 'Malta', 'Netherlands', 'Poland', 'Portugal', 'Romania', 'Slovakia', 'Slovenia', 'Spain', 'Sweden']

# Ship owner feature
# Configure Chrome options for headless mode
chrome_options = Options()
chrome_options.add_argument("--headless")
driver = webdriver.Chrome()
driver.maximize_window()
driver.set_page_load_timeout(20)



owner_data = {}
try:
    for index, ship_imo in enumerate(list(fleet_imo['IMO #'])):
        if not is_driver_active(driver):
            print("Driver not open, opening now...")
            driver = webdriver.Chrome()
            driver.maximize_window()
            driver.set_page_load_timeout(20)
        
        cancel_flag = threading.Event()
        t = threading.Thread(target=timer, args=(cancel_flag,))
        t.start()

        try:
            item = index + 1
            print(item)
            try:
                driver.get(f'https://www.balticshipping.com/vessel/imo/{ship_imo}')
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, 'h1')))
                print("Page loaded succesfully")
            except TimeoutException:
                print("Page load timed out, retrying...")
                continue  # or use driver.refresh() to retry loading the page
            except NoSuchElementException:
                continue
            except WebDriverException as e:
                print(f"Error loading page: {e}")
                continue
            except:
                continue
            try:
                element = driver.find_element("xpath",f'/html/body/div[5]/div[2]/article/div/div/h1').text
            except:
                continue
            print("Before XPath Searching")                               
            try:
                for i in range(12,19): #the owner features is usually in the number 16
                    print("Attempt:", i)
                    try:
                        ship_owner_box = driver.find_element("xpath",f'/html/body/div[5]/div[2]/article/section/div[1]/div/div[3]/table/tbody/tr[{i}]/th').text
                    except NoSuchElementException:
                        ship_owner_box = driver.find_element("xpath",f'/html/body/div[5]/div[2]/article/section/div[1]/div/div/table/tbody/tr[{i}]/th').text

                    if ship_owner_box == 'Owner' or ship_owner_box == 'Owner ':
                        try:
                            ship_owner = driver.find_element("xpath",f'/html/body/div[5]/div[2]/article/section/div[1]/div/div[3]/table/tbody/tr[{i}]/td').text
                            break
                        except NoSuchElementException:
                            ship_owner = driver.find_element("xpath",f'/html/body/div[5]/div[2]/article/section/div[1]/div/div/table/tbody/tr[{i}]/td').text
                    else:
                        ship_owner = 'Not Found'       
            except NoSuchElementException or TimeoutException:
                    ship_owner = 'N/A'
            print("After XPath Searching")
            # ship_owner_list.append(ship_owner)
        
        
            try:
                mmsi = driver.find_element("xpath",f'/html/body/div[5]/div[2]/article/section/div[1]/div/div[3]/table/tbody/tr[2]/td').text
                mmsi_list.append(mmsi)
            except:
                mmsi = ''
                mmsi_list.append(mmsi)
            print(mmsi)

            try:
                driver.get(f'https://www.myshiptracking.com/vessels/mmsi-{mmsi}')
            except:
                continue
            try:
                page_check = driver.find_element('xpath','/html/body/div[9]/div[2]/div[3]/div/div[6]/div/div/div/div[1]/div/div[1]/h1').text
                print(page_check)
            except NoSuchElementException:
                continue
            
            # TRAVELED TO COUNTRIES
            for i in range(1,7):
                try:
                    # lookup_xpath1 = f'/html/body/div[9]/div[2]/div[3]/div/div[7]/div/div[6]/div/div[2]/table/tbody/tr[{i}]/td[2]/a'
                    lookup_xpath2 = f'/html/body/div[9]/div[2]/div[3]/div/div[7]/div/div[6]/div/div[2]/table/tbody/tr[{i}]/td[3]'
                    lookup_xpathflag = f'/html/body/div[9]/div[2]/div[3]/div/div[7]/div/div[6]/div/div[2]/table/tbody/tr[{i}]/td[2]/a/img'
                    # grabber = wait.until(EC.visibility_of_element_located((By.XPATH, lookup_xpath1)))
                    # grabber = driver.find_element('xpath', lookup_xpath1)
                    # port_name = grabber.text
                    # grabber = wait.until(EC.visibility_of_element_located((By.XPATH, lookup_xpath2)))
                    grabber = driver.find_element('xpath', lookup_xpath2)
                    port_visits = grabber.text
                    # grabber = wait.until(EC.visibility_of_element_located((By.XPATH, lookup_xpathflag)))
                    grabber = driver.find_element('xpath', lookup_xpathflag)
                    port_flag = grabber.get_attribute('title')
                    if port_flag.startswith(" "):
                        port_flag = port_flag[1:]
                    
                    try:
                        port_visits = int(port_visits)
                    except ValueError:
                        port_visits = 0
                    
                    #Dataframe Building
                    if ship_owner not in owner_data:
                        owner_data[ship_owner] = {}
                        owner_data[ship_owner]['EU'] = 0     
                    
                    if port_flag in owner_data[ship_owner]:
                        owner_data[ship_owner][port_flag] += port_visits
                    else:
                        owner_data[ship_owner][port_flag] = port_visits
                    
                    if port_flag in EU_countries:
                        owner_data[ship_owner]['EU'] += port_visits
                    
                    print("FULL DATA")
                except (NoSuchElementException, TimeoutException) as e:
                    print(f"Encountered error for ship IMO {ship_imo}, skipping")
                    continue 

        except KeyboardInterrupt:
            print("Keyboard interrupt detected, next ship...")
            continue

        cancel_flag.set()

except KeyboardInterrupt:
    print("Interrupted by user, exiting...")

except Exception as e:
    print(f"Encountered a critical error: {e}. Exiting the loop now...")

finally:
    print("LOOP DONE!!!")
    driver.quit()
    # owner_df = pd.DataFrame([{'Owner': owner, 'Country': country, 'Visits': visits}
    #                          for owner, countries in owner_data.items()
    #                          for country, visits in countries.items()])


    owner_df = pd.DataFrame.from_dict(owner_data, orient='index')
    owner_df = owner_df.reset_index().rename(columns={'index':'Owner'})

    ship_type = 'BulkCarrier_80k180k'
    current_date_time = datetime.now()
    date_string = current_date_time.strftime("%Y%m%d")
    owner_df.to_excel(f'C:\\Users\\jacks\\OneDrive\\Documents\\CalcareaCode\\Fleet_Lists\\ownermanager_data_{ship_type}_{date_string}.xlsx')

    print("EXPORTED!!!")
    winsound.Beep(1000, duration=1000)
