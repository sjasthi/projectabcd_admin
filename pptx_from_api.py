import os
import requests
import json
from pptx import Presentation
from pptx.util import Inches, Pt

# Declare the API base URL as a global constant
API_BASE_URL = "https://abcd2.projectabcd.com/api/getinfo.php?id="

def fetch_data_from_api(id):
    url = API_BASE_URL + str(id)
    headers = {
        "Accept": "application/json",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36"
    }

    # Make API request and check if the response is successful (status code 200):
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        return response.json()['data']
    else:
        print(f"Failed to get data from the API for ID {id}. Status code:", response.status_code)
        return None

def create_slide(prs, data):
    slide_master = prs.slide_master
    slide_layout = slide_master.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)

    title_shape = slide.shapes.title
    title_shape.text = data['name']

    content_shape = slide.placeholders[1]
    content_text = f"Description: {data['description']}\n"
    content_text += f"Did You Know: {data.get('did_you_know', '')}\n"
    content_text += f"Category: {data.get('category', '')}\n"
    content_text += f"Type: {data.get('type', '')}\n"
    content_text += f"State Name: {data.get('state_name', '')}\n"
    content_text += f"Key Words: {data.get('key_words', '')}\n"
    content_text += f"Image URL: {data.get('image_url', '')}\n"
    content_text += f"Status: {data.get('status', '')}\n"
    content_text += f"Notes: {data.get('notes', '')}\n"
    content_text += f"Book: {data.get('book', '')}\n"
    content_shape.text = content_text
    
    # Add the image to the slide
    image_filename = data.get('image_url', '')
    if image_filename:
        image_path = os.path.join(os.path.dirname(__file__), image_filename)
        left = Inches(1)
        top = Inches(2)
        pic = slide.shapes.add_picture(image_path, left, top)

def main():
    prs = Presentation()
    prs.slide_width = Inches(8)
    prs.slide_height = Inches(11)

    # Read the ID numbers from 'slide_numbers.txt' file
    with open('slide_numbers.txt', 'r') as file:
        ids_str = file.read()

    # Convert the comma-separated numbers into a list of integers
    ids = [int(id.strip()) for id in ids_str.split(',')]

    for id in ids:
        data = fetch_data_from_api(id)
        if data:
            create_slide(prs, data)

    prs.save("api.pptx")
    print("PowerPoint file 'api.pptx' created successfully.")

if __name__ == "__main__":
    main()
