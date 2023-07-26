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

def create_slide(prs, data, preferences):
    # Get preferences
    display_option = preferences.get('DISPLAY OPTION', 3)
    text_font = preferences.get('TEXT_FONT', 'Arial')
    heading_font = preferences.get('HEADING_FONT', 'Arial')
    text_size = int(preferences.get('TEXT_SIZE', 18))
    heading_size = int(preferences.get('HEADING_SIZE', 24))

    slide_master = prs.slide_master
    slide_layout = slide_master.slide_layouts[display_option]
    slide = prs.slides.add_slide(slide_layout)

    # Check if the layout has a title placeholder
    has_title_placeholder = hasattr(slide, "placeholders") and len(slide.placeholders) > 0
    if has_title_placeholder:
        # Customize font size and font for title
        title_shape = slide.shapes.title
        title_shape.text = data['name']
        title_shape.text_frame.paragraphs[0].font.size = Pt(heading_size)
        title_shape.text_frame.paragraphs[0].font.name = heading_font
    else:
        # No title placeholder, set the title text directly on the slide
        slide.shapes.title.text = data['name']
        slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(heading_size)
        slide.shapes.title.text_frame.paragraphs[0].font.name = heading_font

    # Find the content placeholder shape on the slide
    content_shape = None
    for shape in slide.shapes:
        if shape.has_text_frame and shape != title_shape:
            content_shape = shape
            break

    if content_shape:
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

        # Apply font properties to all paragraphs in content text
        for paragraph in content_shape.text_frame.paragraphs:
            paragraph.font.size = Pt(text_size)
            paragraph.font.name = text_font

     # Add the image to the content placeholder on the slide
    image_filename = data.get('image_url', '')
    if image_filename:
        image_path = os.path.join(os.path.dirname(__file__), image_filename)

        # Set the desired width and height of the image (e.g., Inches(4) and Inches(3))
        image_width = Inches(5)
        image_height = Inches(8)

        left = Inches(1)
        top = Inches(2)
        pic = slide.shapes.add_picture(image_path, left, top, width=image_width, height=image_height)

def read_preferences_from_file(file_path):
    preferences = {}
    with open(file_path, 'r') as file:
        for line in file:
            key, value = line.strip().split('=')
            preferences[key.strip()] = value.strip()
    return preferences

# Function to handle numbers with dashes
def expand_ranges(ids_str):
    expanded_ids = []
    for part in ids_str.split(','):
        part = part.strip()
        if '-' in part:
            start, end = map(int, part.split('-'))
            expanded_ids.extend(range(start, end + 1))
        else:
            expanded_ids.append(int(part))
    return expanded_ids

def main():
    prs = Presentation()
    prs.slide_width = Inches(8)
    prs.slide_height = Inches(11)

    # Read preferences from 'preferences.txt' file
    preferences_file_path = 'preferences.txt'
    preferences = read_preferences_from_file(preferences_file_path)

    # Read the ID numbers from 'slide_numbers.txt' file
    with open('slide_numbers.txt', 'r') as file:
        ids_str = file.read()

    # Expand ranges and convert comma-separated numbers into a list of integers
    ids = expand_ranges(ids_str)

    for id in ids:
       data = fetch_data_from_api(id)
       if data:
            create_slide(prs, data, preferences)

    prs.save("api.pptx")
    print("PowerPoint file 'api.pptx' created successfully.")

if __name__ == "__main__":
    main()
