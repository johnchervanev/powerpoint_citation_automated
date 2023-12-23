import os
import time
import csv
import logging
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# Constants
CSV_FILE_NAME = 'urls.csv'
OUTPUT_IMAGES_FOLDER = 'output_images'
TIMEOUT = 3
RETRIES = 1

# Set up logging
logging.basicConfig(filename='script_log.txt', level=logging.ERROR)

def wait_for_element(driver, by, value, timeout=TIMEOUT, refresh_on_timeout=False):
    for _ in range(RETRIES):
        try:
            return WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((by, value))
            )
        except TimeoutException:
            logging.error("TimeoutException occurred. Retrying after refreshing the page...")
            if refresh_on_timeout:
                driver.refresh()
                time.sleep(2)
            else:
                logging.error("Retrying without refreshing the page...")
            time.sleep(2)
    raise TimeoutException("Exceeded maximum retries")

def search_images_and_extract_urls_bing(driver, image_folder, csv_writer):
    driver.get("https://www.bing.com/")
    time.sleep(2)

    try:
        search_by_image_button = wait_for_element(driver, By.XPATH, '//*[@id="sb_sbip"]', refresh_on_timeout=True)
        search_by_image_button.click()
    except TimeoutException:
        logging.error("Exceeded maximum retries for locating the 'Search by image' button. Aborting search.")
        return

    wait_for_element(driver, By.CSS_SELECTOR, "input[type='file']")

    image_paths = sorted([os.path.join(image_folder, file) for file in os.listdir(image_folder) if file.endswith(('.jpg', '.jpeg', '.png'))], key=lambda x: int(''.join(filter(str.isdigit, x))))

    for image_path in image_paths:
        file_input = wait_for_element(driver, By.CSS_SELECTOR, "input[type='file']")
        time.sleep(5)
        file_input.send_keys(image_path)
        time.sleep(3)

        try:
            element_to_right_click = wait_for_element(driver, By.XPATH, '//a[@class=\'richImgLnk\']', refresh_on_timeout=True)
            url = element_to_right_click.get_attribute("href")

            if url.startswith("https://www.bing.com/images"):
                logging.warning(f"Image link starts with 'https://www.bing.com/images'. Switching to Google...")

                # Switch to Google and perform the search again
                search_images_and_extract_urls_google(driver, image_folder, csv_writer, image_path)
            else:
                csv_writer.writerow([image_path, url])

        except TimeoutException:
            logging.error(f"Timeout for {image_path}: No element found to right-click. Writing 'no source available'.")
            
            # Switch to Google and perform the search again
            search_images_and_extract_urls_google(driver, image_folder, csv_writer, image_path)

        except Exception as e:
            logging.error(f"Error processing {image_path}: {e}")

        driver.get("https://www.bing.com/")
        time.sleep(2)

        try:
            search_by_image_button = wait_for_element(driver, By.XPATH, '//*[@id="sb_sbip"]', refresh_on_timeout=True)
            search_by_image_button.click()
        except TimeoutException:
            logging.error("Exceeded maximum retries for locating the 'Search by image' button. Aborting search.")
            return

        wait_for_element(driver, By.CSS_SELECTOR, "input[type='file']")

def search_images_and_extract_urls_google(driver, image_folder, csv_writer, current_image_path):
    driver.get("https://www.google.com/imghp")
    time.sleep(2)

    try:
        search_by_image_button = wait_for_element(driver, By.XPATH, '//div[@jscontroller=\'lpsUAf\' and @jsname=\'R5mgy\' and @class=\'nDcEnd\' and @aria-label=\'Search by image\']', refresh_on_timeout=True)
        search_by_image_button.click()
    except TimeoutException:
        logging.error("Exceeded maximum retries for locating the 'Search by image' button. Aborting search.")
        return

    wait_for_element(driver, By.CSS_SELECTOR, "input[type='file']")

    file_input = wait_for_element(driver, By.CSS_SELECTOR, "input[type='file']")
    time.sleep(5)
    file_input.send_keys(current_image_path)
    time.sleep(3)

    # Find the element using XPath
    element = driver.find_element(By.XPATH, "//div[@class='ICt2Q']")

    # Perform the click action
    element.click()

    try:
        # Find the element using the specified XPath
        element_to_right_click = wait_for_element(driver, By.XPATH, '//li[@class=\'anSuc\']/a', refresh_on_timeout=True)

        # Extract the 'href' attribute value from the found element
        url = element_to_right_click.get_attribute("href")

        # Write the current_image_path and url to the CSV file
        csv_writer.writerow([current_image_path, url])

    except TimeoutException:
        logging.error(f"Timeout for {current_image_path}: No element found to right-click. Writing 'no source available'.")
        csv_writer.writerow([current_image_path, "no source available"])

    except Exception as e:
        logging.error(f"Error processing {current_image_path}: {e}")

    driver.get("https://www.google.com/imghp")
    time.sleep(2)

    try:
        search_by_image_button = wait_for_element(driver, By.XPATH, '//div[@jscontroller=\'lpsUAf\' and @jsname=\'R5mgy\' and @class=\'nDcEnd\' and @aria-label=\'Search by image\']', refresh_on_timeout=True)
        search_by_image_button.click()
    except TimeoutException:
        logging.error("Exceeded maximum retries for locating the 'Search by image' button. Aborting search.")
        return

    wait_for_element(driver, By.CSS_SELECTOR, "input[type='file']")

def main():
    current_directory = os.path.dirname(os.path.realpath(__file__))
    csv_file_path = os.path.join(current_directory, CSV_FILE_NAME)

    # Construct the path to ChromeDriver in the same directory as the script
    driver_path = os.path.join(current_directory, 'chromedriver')

    chrome_options = webdriver.ChromeOptions()

    driver = webdriver.Chrome(service=ChromeService(executable_path=driver_path), options=chrome_options)

    try:
        IMAGE_FOLDER = os.path.join(current_directory, OUTPUT_IMAGES_FOLDER)

        if os.path.exists(IMAGE_FOLDER) and os.path.isdir(IMAGE_FOLDER):
            with open(csv_file_path, 'w', newline='') as csv_file:
                csv_writer = csv.writer(csv_file)
                search_images_and_extract_urls_bing(driver, IMAGE_FOLDER, csv_writer)
        else:
            logging.error(f"The '{OUTPUT_IMAGES_FOLDER}' directory is missing. Please run the script to create it.")
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
