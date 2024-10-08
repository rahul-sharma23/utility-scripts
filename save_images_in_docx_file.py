import os
import docx
from docx.shared import Inches
from PIL import Image
from docx import Document

def create_word_document(image_files, output_file):
    doc = Document()

    for image_file in image_files:
        # Open the image and rotate it by 90 degrees
        image_path = image_file
        image = Image.open(image_path)
        rotated_image = image.rotate(90, expand=True)  # Rotate the image by 90 degrees

        # Save the rotated image to a temporary file (to insert into the Word document)
        rotated_image_path = image_file
        rotated_image.save(rotated_image_path)

        # Get dimensions of the rotated image
        width, height = rotated_image.size

        # Calculate aspect ratio to scale the image properly to fit A4 page (8.27 x 11.69 inches)
        aspect_ratio = width / height
        page_width = 6.27  # A4 page width in inches
        page_height = 9.69  # A4 page height in inches

        # Scale image to fill the page while maintaining aspect ratio
        if aspect_ratio > (page_width / page_height):
            image_width = page_width
            image_height = page_width / aspect_ratio
        else:
            image_height = page_height
            image_width = page_height * aspect_ratio

        # Add the rotated image to the document
        doc.add_picture(rotated_image_path, width=Inches(image_width), height=Inches(image_height))


    doc.save(output_file)

if __name__ == "__main__":
    image_files = [f for f in os.listdir('.') if f.endswith('.jpg')]
    output_file = "images.docx"
    create_word_document(image_files, output_file)