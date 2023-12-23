import os
import csv
import re
from pptx import Presentation
from pptx.util import Inches, Pt
from tkinter import Tk, filedialog

def create_textbox(slide, left, top, width, height):
    textbox = slide.shapes.add_textbox(left, top, width, height)
    return textbox.text_frame

def customize_paragraph(p):
    p.font.size = Pt(7)
    p.alignment = 2  # Right-aligned (adjust as needed)

def add_footer_to_slide(slide, urls, presentation):
    slide_width = presentation.slide_width
    slide_height = presentation.slide_height
    left = Inches(0.5)
    top = slide_height - Inches(0.5)
    width = slide_width - Inches(1.0)
    height = Inches(0.5)

    text_frame = create_textbox(slide, left, top, width, height)
    
    p = text_frame.add_paragraph()
    for url in urls:
        p.text += url + '\n'
    
    customize_paragraph(p)

def read_csv(csv_file):
    data = {}
    with open(csv_file, newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            if row and len(row) >= 2:
                image_path, url = row[0], row[1]

                match = re.search(r'slide_(\d+)_image_\d+', image_path)
                if match:
                    slide_number = int(match.group(1))
                    data.setdefault(slide_number, []).append(url)
    return data

def add_footer_to_slides(input_pptx, output_folder, csv_file):
    presentation = Presentation(input_pptx)
    data = read_csv(csv_file)

    for i, slide in enumerate(presentation.slides):
        if i + 1 in data:
            urls = data[i + 1]
            add_footer_to_slide(slide, urls, presentation)

    os.makedirs(output_folder, exist_ok=True)

    base_name = os.path.splitext(os.path.basename(input_pptx))[0]
    output_path = os.path.join(output_folder, f"{base_name}_modified.pptx")
    presentation.save(output_path)

if __name__ == "__main__":
    Tk().withdraw()
    input_pptx = filedialog.askopenfilename(title="Select PowerPoint file", filetypes=[("PowerPoint files", "*.pptx")])

    # Set the output folder to the main directory
    script_directory = os.path.dirname(os.path.abspath(__file__))
    output_folder = os.path.join(script_directory, "output_pptx")

    # Create the output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    csv_file = os.path.join(script_directory, "urls.csv")
    add_footer_to_slides(input_pptx, output_folder, csv_file)
