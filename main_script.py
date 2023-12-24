# File: main_script.py
import os
from extract_images import extract_images_from_pptx, choose_file
from selenium_automation import main as search_images_and_extract_urls
from add_footer import add_footer_to_slides
from tkinter import Tk, filedialog

def run_scripts():
    # Script 1
    try:
        input_pptx = choose_file()
        if not input_pptx:
            print("File selection canceled. Exiting script.")
            return

        script_directory = os.path.dirname(os.path.abspath(__file__))
        output_images_folder = os.path.join(script_directory, "output_images")
        
        # Use the name of the PowerPoint file as the template
        template = os.path.splitext(os.path.basename(input_pptx))[0]

        # Extract images from PowerPoint and mark slide numbers
        extract_images_from_pptx(input_pptx, output_images_folder, template)
    except Exception as e:
        print(f"Error in Script 1: {e}")

    # Script 2
    try:
        search_images_and_extract_urls()
    except Exception as e:
        print(f"Error in Script 2: {e}")

    # Script 3
    try:
        csv_file = os.path.join(script_directory, "urls.csv")
        output_pptx_folder = os.path.join(script_directory, "output_pptx")

        # Create the output PowerPoint folder if it doesn't exist
        if not os.path.exists(output_pptx_folder):
            os.makedirs(output_pptx_folder)

        add_footer_to_slides(input_pptx, output_pptx_folder, csv_file)
    except Exception as e:
        print(f"Error in Script 3: {e}")

if __name__ == "__main__":
    run_scripts()
