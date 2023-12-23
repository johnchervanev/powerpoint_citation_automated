import os
from pptx import Presentation
from PIL import Image, ImageTk, ImageFile
from io import BytesIO
from tkinter import filedialog, Tk

ImageFile.LOAD_TRUNCATED_IMAGES = True

def convert_to_rgb(image):
    if image.mode in ("RGBA", "P"):
        return image.convert("RGB")
    return image

def process_image(image, output_folder, slide_number, image_count, template):
    try:
        image_data = image.blob
        with BytesIO(image_data) as image_stream:
            img = Image.open(image_stream)
            img = convert_to_rgb(img)

            img_path = os.path.join(output_folder, f"{template}__fin_slide_{slide_number}_image_{image_count + 1}.jpg")
            img.save(img_path, "JPEG")

            return True
    except Exception as e:
        print(f"Error processing image: {e}")
        return False

def extract_and_process_images(presentation, output_folder, template):
    image_count = 0

    for slide_number, slide in enumerate(presentation.slides, start=1):
        for shape_number, shape in enumerate(slide.shapes, start=1):
            if hasattr(shape, "image"):
                if process_image(shape.image, output_folder, slide_number, image_count, template):
                    image_count += 1

    return image_count

def extract_images_from_pptx(input_pptx, output_folder, template):
    # Load the PowerPoint presentation
    presentation = Presentation(input_pptx)

    # Create the output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    # Extract and process images
    image_count = extract_and_process_images(presentation, output_folder, template)

    print(f"{image_count} images extracted from the PowerPoint presentation.")

def choose_file():
    Tk().withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(title="Select PowerPoint file", filetypes=[("PowerPoint files", "*.pptx")])
    return file_path

if __name__ == "__main__":
    # Choose the input PowerPoint file interactively
    input_pptx = choose_file()

    # Set the output folder to the directory of the script
    script_directory = os.path.dirname(os.path.abspath(__file__))
    output_folder = os.path.join(script_directory, "output_images")

    # Use the name of the PowerPoint file as the template
    template = os.path.splitext(os.path.basename(input_pptx))[0]

    # Extract images from PowerPoint and mark slide numbers
    extract_images_from_pptx(input_pptx, output_folder, template)
