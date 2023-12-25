import os
import csv
import logging
import json
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException

# Additional Constants
CONFIG_FILE_PATH = 'config.json'

# Load Configurations from JSON file
try:
    with open(CONFIG_FILE_PATH, 'r') as config_file:
        CONFIG = json.load(config_file)
        output_directory = CONFIG.get('OUTPUT_DIRECTORY')
except FileNotFoundError:
    # Provide default values if the configuration file is missing
    CONFIG = {
        'CSV_FILE_NAME': 'urls.csv',
        'OUTPUT_IMAGES_FOLDER': 'output_images',
        'TIMEOUT': 10,
        'RETRIES': 2
    }
    output_directory = None

# Set up logging
logging.basicConfig(filename='script_log.txt', level=logging.INFO)
logger = logging.getLogger(__name__)

def wait_for_page_to_load(driver, timeout=CONFIG['TIMEOUT']):
    WebDriverWait(driver, timeout).until(
        lambda x: x.execute_script("return document.readyState === 'complete'")
    )

def wait_for_element(driver, by, value, timeout=CONFIG['TIMEOUT'], refresh_on_timeout=False):
    for _ in range(CONFIG['RETRIES']):
        try:
            return WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((by, value))
            )
        except TimeoutException:
            logger.warning("TimeoutException occurred. Retrying after refreshing the page...")
            if refresh_on_timeout:
                driver.refresh()
                time.sleep(2)
            else:
                logger.warning("Retrying without refreshing the page...")
            time.sleep(2)
        except Exception as e:
            logger.error(f"An error occurred while waiting for an element: {e}")
    raise TimeoutException("Exceeded maximum retries")

def search_images_and_extract_urls_bing(driver, image_folder, csv_writer):
    driver.get("https://www.bing.com/")
    wait_for_page_to_load(driver)
    time.sleep(2)

    try:
        search_by_image_button = wait_for_element(driver, By.XPATH, '//*[@id="sb_sbip"]', refresh_on_timeout=True)
        search_by_image_button.click()
    except TimeoutException:
        logger.error("Exceeded maximum retries for locating the 'Search by image' button. Aborting search.")
        return

    wait_for_element(driver, By.CSS_SELECTOR, "input[type='file']")

    image_paths = sorted([os.path.join(image_folder, file) for file in os.listdir(image_folder) if file.endswith(('.jpg', '.jpeg', '.png'))], key=lambda x: int(''.join(filter(str.isdigit, x))))

    for image_path in image_paths:
        try:
            file_input = wait_for_element(driver, By.CSS_SELECTOR, "input[type='file']")
            time.sleep(5)
            file_input.send_keys(image_path)
            time.sleep(3)

            element_to_right_click = wait_for_element(driver, By.XPATH, '//a[@class=\'richImgLnk\']', refresh_on_timeout=True)
            url = element_to_right_click.get_attribute("href")

            if url.startswith("https://www.bing.com/images"):
                logger.warning(f"Image link starts with 'https://www.bing.com/images'. Switching to Google...")

                # Switch to Google and perform the search again
                search_images_and_extract_urls_google(driver, image_folder, csv_writer, image_path)
            else:
                csv_writer.writerow([image_path, url])

        except TimeoutException:
            logger.error(f"Timeout for {image_path}: No element found to right-click. Writing 'no source available'.")

            # Additional logging to help diagnose the issue
            logger.debug(driver.page_source)

            # Switch to Google and perform the search again
            search_images_and_extract_urls_google(driver, image_folder, csv_writer, image_path)

        except Exception as e:
            logger.error(f"Error processing {image_path}: {e}")

        finally:
            driver.get("https://www.bing.com/")
            wait_for_page_to_load(driver)
            time.sleep(2)

            try:
                search_by_image_button = wait_for_element(driver, By.XPATH, '//*[@id="sb_sbip"]', refresh_on_timeout=True)
                search_by_image_button.click()
            except TimeoutException:
                logger.error("Exceeded maximum retries for locating the 'Search by image' button. Aborting search.")
                return

            wait_for_element(driver, By.CSS_SELECTOR, "input[type='file']")

def search_images_and_extract_urls_google(driver, image_folder, csv_writer, current_image_path):
    driver.get("https://www.google.com/imghp")
    wait_for_page_to_load(driver)
    time.sleep(2)

    try:
        search_by_image_button = wait_for_element(driver, By.XPATH, '//div[@jscontroller=\'lpsUAf\' and @jsname=\'R5mgy\' and @class=\'nDcEnd\' and @aria-label=\'Search by image\']', refresh_on_timeout=True)
        search_by_image_button.click()
    except TimeoutException:
        logger.error("Exceeded maximum retries for locating the 'Search by image' button. Aborting search.")
        return

    wait_for_element(driver, By.CSS_SELECTOR, "input[type='file']")

    try:
        file_input = wait_for_element(driver, By.CSS_SELECTOR, "input[type='file']")
        time.sleep(5)
        file_input.send_keys(current_image_path)
        time.sleep(3)

        element = driver.find_element(By.XPATH, "//div[@class='ICt2Q']")
        element.click()

        element_to_right_click = wait_for_element(driver, By.XPATH, '//li[@class=\'anSuc\']/a', refresh_on_timeout=True)
        url = element_to_right_click.get_attribute("href")

        csv_writer.writerow([current_image_path, url])

    except TimeoutException:
        logger.error(f"Timeout for {current_image_path}: No element found to right-click. Writing 'no source available'.")
        csv_writer.writerow([current_image_path, "no source available"])

    except Exception as e:
        logger.error(f"Error processing {current_image_path}: {e}")

    finally:
        driver.get("https://www.google.com/imghp")
        wait_for_page_to_load(driver)
        time.sleep(2)

        try:
            search_by_image_button = wait_for_element(driver, By.XPATH, '//div[@jscontroller=\'lpsUAf\' and @jsname=\'R5mgy\' and @class=\'nDcEnd\' and @aria-label=\'Search by image\']', refresh_on_timeout=True)
            search_by_image_button.click()
        except TimeoutException:
            logger.error("Exceeded maximum retries for locating the 'Search by image' button. Aborting search.")
            return

        wait_for_element(driver, By.CSS_SELECTOR, "input[type='file']")

def main():
    current_directory = os.path.dirname(os.path.realpath(__file__))
    dynamic_directory = CONFIG.get('OUTPUT_IMAGES_FOLDER', 'default_dynamic_directory')
    csv_file_path = os.path.join(current_directory, dynamic_directory, '..', CONFIG['CSV_FILE_NAME'])


    # Construct the path to ChromeDriver in the same directory as the script
    driver_path = os.path.join(current_directory, 'chromedriver')

    chrome_options = webdriver.ChromeOptions()

    try:
        with webdriver.Chrome(service=ChromeService(executable_path=driver_path), options=chrome_options) as driver:
            # Wait for the page to load before proceeding
            wait_for_page_to_load(driver)

            IMAGE_FOLDER = os.path.join(current_directory, dynamic_directory)

            if os.path.exists(IMAGE_FOLDER) and os.path.isdir(IMAGE_FOLDER):
                if os.path.exists(csv_file_path):
                    action = input(f"The CSV file '{CONFIG['CSV_FILE_NAME']}' already exists. Do you want to overwrite it? (yes/no): ").lower()
                    if action != 'yes':
                        logger.warning("User chose not to overwrite the existing CSV file. Aborting script.")
                        return

                with open(csv_file_path, 'w', newline='') as csv_file:
                    csv_writer = csv.writer(csv_file)
                    search_images_and_extract_urls_bing(driver, IMAGE_FOLDER, csv_writer)

                # All images processed
                logger.info("All images have been processed.")

            else:
                logger.error(f"The '{dynamic_directory}' directory is missing. Please run the script to create it.")
    except WebDriverException as wde:
        logger.error(f"WebDriverException occurred: {wde}")
    except Exception as e:
        logger.error(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    main()
