from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import math

import logging

import openpyxl
from openpyxl import Workbook
from datetime import datetime

# The number of cents is the number of weeks from the date on the clearance tag until the next markdown. .03 means three weeks.

def log_clearance_item_to_excel(key, store, value):
    filename = 'clearance_items.xlsx'
    try:
        wb = openpyxl.load_workbook(filename)
        ws = wb.active
    except FileNotFoundError:
        wb = Workbook()
        ws = wb.active
        ws.append(['Store', 'Date Found', 'Price', 'MSRP', 'Key'])  # Header row

    ws.append([store, datetime.now(), value[0], value[1], key])
    wb.save(filename)

def setup_logger():
    # Create a logger
    logger = logging.getLogger('debug_logger')
    logger.setLevel(logging.DEBUG)  # Set the logging level

    # Create a file handler
    file_handler = logging.FileHandler('logfile.txt')
    file_handler.setLevel(logging.DEBUG)

    # Create a formatter
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    # Add the formatter to the file handler
    file_handler.setFormatter(formatter)

    # Add the file handler to the logger
    logger.addHandler(file_handler)

    return logger

def setup_clearance_logger():
    # Create a logger
    logger = logging.getLogger('clearance_logger')
    logger.setLevel(logging.INFO)  # Set the logging level

    # Create a file handler
    file_handler = logging.FileHandler('scav.txt')
    file_handler.setLevel(logging.INFO)

    # Create a formatter
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    # Add the formatter to the file handler
    file_handler.setFormatter(formatter)

    # Add the file handler to the logger
    logger.addHandler(file_handler)

    return logger

def wait_random_time(low, high):
    import random
    import time
    time.sleep(random.uniform(low, high))

def extract_clearance_links(html):
    soup = BeautifulSoup(html, 'html.parser')

    # Find all 'See In-Store Clearance Price' elements
    clearance_elements = soup.find_all('span', string='See In-Store Clearance Price')

    # Find the next 'a' element after each 'See In-Store Clearance Price' element and get its href
    links = [element.find_next('a').get('href') for element in clearance_elements]

    return links

def find_clearance():
    wait_random_time(2, 4)
    try:
        elements = driver.find_elements(By.XPATH, '//span[contains(text(), "See In-Store Clearance Price")]')
        elements_dict = {}

        for element in elements:
            element.click()

            parent_element = element.find_element(By.XPATH, '..')
            parent_element = parent_element.find_element(By.XPATH, '..')
            parent_element = parent_element.find_element(By.XPATH, '..')
            parent_element = parent_element.find_element(By.XPATH, '..')
            parent_element = parent_element.find_element(By.XPATH, '..')

            # Find the element with href that starts with "/p/DEWALT-3" within the parent element
            href_element = parent_element.find_element(By.XPATH, './/a[starts-with(@href, "/p/")]')

            # Grab the href attribute
            href = href_element.get_attribute('href')
            try: 
                discount_price_element = driver.find_element(By.CLASS_NAME, 'clearance-price')
                discount_price = discount_price_element.text
            except:
                discount_price = "N/A"
            
            try:
                msrp_element = driver.find_element(By.CLASS_NAME, 'sui-line-through')
                msrp = msrp_element.text
            except:
                msrp = "N/A"
            
            elements_dict[href] = (discount_price, msrp)
        
        return elements_dict
    except:
        pass

def next_page():
    scroll_down_counter = 0
    scroll_up_counter = 0
    scroll_0_0_counter = 0
    retry_counter = 0
    while True:
        try:
            next_page_link = driver.find_element(By.XPATH, '//a[@aria-label="Next"]')
            next_page_link.click()
            break
        except:
            scroll_down_counter += 1
            scroll_up_counter += 1
            scroll_0_0_counter += 1
            retry_counter += 1
            pass
        
        if scroll_down_counter % 2 == 0:
            driver.execute_script("window.scrollBy(0, 800);")  
            wait_random_time(1, 3)
        
        if scroll_up_counter == 1 or scroll_up_counter % 5 == 0:
            driver.execute_script("window.scrollBy(0, -500);")  
            wait_random_time(1, 3)
        
        if scroll_0_0_counter % 7 == 0:
            driver.execute_script("window.scrollBy(0, -200000);")
            wait_random_time(1, 3)
            driver.execute_script("window.scrollBy(0, -200000);")
            wait_random_time(1, 3)

        if retry_counter > 15:
            break
    
    wait_random_time(2, 3)
    # Scroll to top of page
    driver.execute_script("window.scrollTo(0, 0)")

search_terms = {
    'dewalt': 7, 
    'makita': 7, 
    'milwaukee': 7, 
    'bosch': 5, 
    'ridgid': 4, 
    'fiskars': 1,
    'storage shed': 2, 
    'water heater': 2, 
    'cleaning supplies': 3,
    'lawn mower': 2,
    'ladders': 1,
    'air compressor': 1,
    'grills': 2,
    'patio furniture': 2,
    'ceiling fan': 2,
    'light fixtures': 2,
    'bathroom vanity': 2,
    'painting': 2,
    'paint sprayers': 3,
    'garage epoxy': 2,
    'flooring': 2,
    'tile': 2,
    'carpet': 2,
    'electronics': 2,
    'appliances': 2,
    'glue': 2,
    'grills': 2,
    'patio furniture': 2,
    'griddle': 2,
    'pest control': 2,
    'weed killer': 2,
    'garage': 2,
    'google': 2,
    'smart home': 2,
    'google nest': 2,
    'concrete mixers': 2,
}

hd_stores = [
    4409, # Sandy
    4408, # Centerville
    4406, # W Valley City
    4416, # Provo
    4461, # Saragota Springs
    4407, # Lindon
    4417, # American Fork
    4421, # E Sandy
    4409, # Sandy
    4410, # W Jordan
    # 4415, Park City
    4412, # St George
    4420, # W St George
    4418, # Cedar City
]

find_all_prices = input('Be slow?\n')
find_all_prices = True if find_all_prices.lower() == 'y' else False

driver = webdriver.Chrome()

logger = setup_logger()
clearance_logger = setup_clearance_logger()


for store in hd_stores:
    logger.info(f'-------------------------- Searching for store { store } --------------------------')

    while True:
        try:
            driver.get('https://www.homedepot.com/')
        except:
            logger.info(f'#{ store } - Failed to load homedepot.com')
            continue

        wait_random_time(2, 4)

        try:
            element = driver.find_element(By.XPATH, '//*[@id="header-content"]/div/div[2]/div/button/div[2]/p')
            element.click()
            
        except Exception as e:
            logger.info(f'#{ store } - Failed to find the element: {str(e)}')

            try:
                button = driver.find_element(By.XPATH, '//button[@data-testid="my-store-button"]')
                button.click()
            except:
                logger.info(f'#{ store } - Failed to click on My Store button')
                continue

        wait_random_time(2, 4)

        try:
            parent_element = driver.find_element(By.XPATH, '//div[@data-component="SearchInput"]')
            input_field = parent_element.find_element(By.XPATH, './/input[@type="text"]')
            input_field.send_keys(store)
            input_field.submit()
            logger.info(f'#{ store } - Found store using SearchInput search')
        except:
            try:
                input_field = driver.find_element(By.XPATH, '//input[@placeholder="ZIP Code, City, State, or Store #"]')
                input_field.send_keys(store)
                input_field.submit()
                logger.info(f'#{ store } - Found store using placeholder search')
                
            except:
                logger.info(f'#{ store } - Failed to search for store using placeholder search')
                continue


        wait_random_time(2, 4)

        try:
            shop_button = driver.find_element(By.XPATH, '//button[@data-testid="store-pod-localize__button"]')
            shop_button.click()
        except:
            logger.info(f'#{ store } - Failed to click on Shop button') 
            continue
        
        wait_random_time(2, 4)
        break
    
    store_sale_items = {}
    
    for search_term, pages in search_terms.items():
        logger.info(f'-------------------------- Searching for { search_term } --------------------------')

        while True:
            try:
                driver.get('https://www.homedepot.com/')
                wait_random_time(1, 2)

                search_bar = driver.find_element(By.XPATH, '//*[@id="typeahead-search-field-input"]')
                search_bar.clear()
                search_bar.send_keys(search_term)
                break
            except:
                continue

        wait_random_time(0, 1)
        search_bar.submit()
        wait_random_time(6, 9)

        logger.info(f'#{ store } - Found { search_term }')

        retry_counter = 0
        retry_limit = 15

        while True:
            if retry_counter > retry_limit:
                if retry_counter > retry_limit + 2:
                    break

                driver.refresh()
            
            wait_random_time(2, 3)
            try:
                driver.find_element(By.LINK_TEXT, 'In Stock at Store Today').click()
                logger.info(f'#{ store } - Clicked on In Stock at Store Today')
                break 
            except:
                logger.info(f'#{ store } - Failed to click on In Stock at Store Today')

                try: 
                    driver.find_element(By.PARTIAL_LINK_TEXT, 'Shop All').click()
                    logger.info(f'#{ store } - Clicked on Shop All')
                except:
                    logger.info(f'#{ store } - Failed to click on Shop All')
                    retry_counter += 1
                    continue
                
                continue
        
        if retry_counter > retry_limit:
            logger.info(f'#{ store } - Failed to find In Stock at Store Today or Shop All')
            continue
        
        wait_random_time(5, 10)

        number_of_pages = pages
        try:
            elements = driver.find_elements(By.CLASS_NAME, 'results-pagination__counts--number')
            number_of_pages = min(math.ceil(int(elements[1].text) / 24), 25)
            logger.info(f'#{ store } - Number of pages: { number_of_pages }')
        except:
            logger.info(f'#{ store } - Failed to find number of pages')
            pass

        for page_num in range(0, number_of_pages): 
            new_sale_items = {}
            should_break = False

            logger.info(f'#{ store } - Finding { search_term } on page { page_num + 1 } / { number_of_pages }')

            clearance_links = extract_clearance_links(driver.page_source)
            for link in clearance_links:
                new_sale_items['https://www.homedepot.com' + link] = ('N/A', 'N/A')

            # Comment this loop out to speed up the process
            if (find_all_prices):
                for _ in range(0, 7):
                    result = find_clearance()

                    if result:
                        new_sale_items.update(result)
                        should_break = True

                    result = find_clearance()
                    if result:
                        new_sale_items.update(result)
                    else:
                        should_break = False
                    
                    if should_break:
                        break
                    
                    driver.execute_script("window.scrollBy(0, 600);")  
            
            logger.info(f'#{ store } - Found { len(new_sale_items) } { search_term } on page { page_num + 1 }')

            for key, value in new_sale_items.items():
                # Remove duplicates
                if key in store_sale_items:
                    continue

                print(f'#{ store } - { value[0] } / { value[1] } - { key }')
                clearance_logger.info(f'#{ store } - { value[0] } / { value[1] } - { key }')
                log_clearance_item_to_excel(key, store, value)
            
            store_sale_items.update(new_sale_items)
            
            logger.info(f'#{ store } - Moving to next page')
            next_page()
            logger.info(f'#{ store } - Moved to next page')
            wait_random_time(2, 3)
