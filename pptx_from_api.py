import requests
import json
from pptx import Presentation
from pptx.util import Inches
from typing import List

# Global constant
API_BASE = 'https://abcd2.projectabcd.com/api/getinfo.php?id='

def generate_the_pptx(id_list: List[int], sort_order: str = 'ascending') -> None:
    presentation = Presentation()

    for id in id_list:
        api_url = API_BASE + str(id)
        #406 error- add header, did not solve 406
        headers = {
            'Accept': 'application/json',
        }
        response = requests.get(api_url, headers=headers)
         
        # Check if request was successful
        if response.status_code == 200:
            # Get JSON data from the response
            data = response.json()

            if 'data' in data:
                data = data['data']
                id = data.get('id', '')
                name = data.get('name', '')
                description = data.get('description', '')
                
                
                # Create a new slide
                slide_layout = presentation.slide_layouts[1]
                slide = presentation.slides.add_slide(slide_layout)

                # Add title
                title_placeholder = slide.shapes.title
                title_placeholder.text = f"ID: {id} - {name}"

                # Add content
                content_placeholder = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8), Inches(5))
                content_frame = content_placeholder.text_frame
                content_frame.text = description

                # Add sorting order
                if sort_order == 'ascending':
                    presentation.slides._sldIdLst = sorted(presentation.slides._sldIdLst, key=lambda x: x.idx)

            else:
                print(f"Data not found for ID {id}")
        else:
            print(f"Failed to retrieve data for ID {id}. Error: {response.status_code} - {response.text}")

    if len(presentation.slides) > 0:
        presentation.save('output.pptx')
    else:
        print("No valid data retrieved. Presentation not saved.")

# Prompt user to enter IDs
id_list_input = input("Enter the IDs: ")
id_list = [int(id.strip()) for id in id_list_input.split(",")]

# Call the function with user-specified IDs
generate_the_pptx(id_list)



'''
BUGS: 406 issue
API endpoint and url is correct 
Header format?  Does api endpoint need additional headers or authentication?
server issue? is it configured to handle the requested format
Mod_Security? -open source web application firewall that provides security for web applications-it intercepts
and inspects requests and responses-if it detects suspicious behavior, it can block the request
