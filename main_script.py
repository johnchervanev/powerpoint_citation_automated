# File: main_script.py
import os
from extract_images import extract_images_from_pptx, choose_file
from selenium_automation import main as search_images_and_extract_urls
from add_footer import add_footer_to_slides
from tkinter import Tk, filedialog

def run_scripts():
    # Script 1
    input_pptx = choose_file()
    script_directory = os.path.dirname(os.path.abspath(__file__))
    output_images_folder = os.path.join(script_directory, "output_images")
    extract_images_from_pptx(input_pptx, output_images_folder)

    # Script 2
    search_images_and_extract_urls()

    # Script 3
    csv_file = os.path.join(script_directory, "urls.csv")
    output_pptx_folder = os.path.join(script_directory, "output_pptx")

    # Create the output PowerPoint folder if it doesn't exist
    if not os.path.exists(output_pptx_folder):
        os.makedirs(output_pptx_folder)

    add_footer_to_slides(input_pptx, output_pptx_folder, csv_file)

if __name__ == "__main__":
    run_scripts()
