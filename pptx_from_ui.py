import tkinter as tk
from tkinter import*
import os
import shutil
from os.path import basename
from collections.abc import Container
import pptx.util
from pptx import Presentation 
from pptx.util import Pt
from bs4 import BeautifulSoup
import requests
import pandas as pd
from sqlalchemy import create_engine
import mysql.connector



window = tk.Tk()

label_text = tk.StringVar()
label_text.set("Enter page numbers")
text_box =tk.Label(window, textvariable=label_text).grid(row=1,column=1)

nums = StringVar(None)
text_value = tk.Entry(window,textvariable=nums).grid(row=1,column=2)

label_text_1 = tk.StringVar()
label_text_1.set("preferences")
text_box_1 =tk.Label(window, textvariable=label_text_1).grid(row=4,column=1)

nums_1 = StringVar(value="preferences.txt")
text_value_1 = tk.Entry(window, textvariable=nums_1).grid(row=4,column=2)

label_text_2 = tk.StringVar()
label_text_2.set("Output file")
text_box_2 =tk.Label(window, textvariable=label_text_2).grid(row=5,column=1)

nums_2 = StringVar(value="abcd_book")
text_value_2 = Entry(window,textvariable=nums_2).grid(row=5,column=2)
btn_1 = tk.Button(window, text= ".pptx").grid(row=5,column=2, columnspan=3)

selected_option = tk.StringVar()
selected_option_1 = tk.StringVar()

btn_text = tk.StringVar()
btn_text.set("Layout")
btn_box =tk.Label(window, textvariable=btn_text).grid(row=2,column=1)

radio_button1 = tk.Radiobutton(window, text="Pic on left", variable=selected_option, value=3).grid(row=2,column=2)
radio_button1 = tk.Radiobutton(window, text="Pic on right", variable=selected_option, value=2).grid(row=2,column=3)
radio_button1 = tk.Radiobutton(window, text="Pic on top", variable=selected_option, value=1).grid(row=2,column=4)


btn_text_1 = tk.StringVar()
btn_text_1.set("Method")
btn_box_1 =tk.Label(window, textvariable=btn_text_1).grid(row=3,column=1)

# Create the second radio button
radio_button2 = tk.Radiobutton(window, text="Web", variable=selected_option_1 , value="Web" ).grid(row=3,column=2)
radio_button2 = tk.Radiobutton(window, text="Excel", variable=selected_option_1, value="Excel").grid(row=3,column=3)
radio_button2 = tk.Radiobutton(window, text="APi", variable=selected_option_1, value="Api").grid(row=3,column=4)
radio_button2 = tk.Radiobutton(window, text="db", variable=selected_option_1, value="Database").grid(row=3,column=5)

def generate_output():
   
    page_numbers = nums.get()
    preferences = nums_1.get()
    output_file = nums_2.get()
    layout_option = selected_option.get()
    method_option = selected_option_1.get()

    global methods,pages,layout,preference,output
    methods= method_option
    pages = page_numbers
    layout = layout_option
    preference = preferences
    output = output_file

    window.destroy()

    
btn = tk.Button(window, text="Generate", command=generate_output)
btn.grid(row=6, column=2, columnspan=2)


window.mainloop()



all_pages = []  
with open("slide_numbers.txt" , "w") as file:
   chosen_nums = pages.split(",")
   
   for word in chosen_nums:
       all_pages.append(word)
       file.write(word + "\n")
file.close

if(methods == "Web"):
    def buildPresentation():  
        with open(preference, "r") as f:
            slideOption = int(layout)
            textFont = f.readline().split("= ")
            textFont = textFont[1]
            titleFont = f.readline().split("= ")
            titleFont = titleFont[1]
            textSize = f.readline().split("= ")
            textSize = int(textSize[1])
            titleSize = f.readline().split("= ")
            titleSize = int(titleSize[1])
        prs = Presentation()
        presentationLength = len(all_pages)
        pictureSlide = 0
       
        if(slideOption == 2):
            pictureSlide = 1
        #web scrapes the URL to get all needed information from the page
        for i in range (0, presentationLength):
            URL = "https://projectabcd.com/display_the_dress.php?id=" + str(all_pages[i])
            page = requests.get(URL, headers={"User-Agent": "html"})
            soup = BeautifulSoup(page.content, "html.parser")
            logo = soup.find("img")
            printLogo = logo.attrs["src"]
            logoURL = "http://projectabcd.com/" + printLogo
            r = requests.get(logoURL, headers={"User-Agent": "html"}, stream=True)
            if r.status_code == 200:
                   with open(basename(printLogo), "wb") as f:
                        r.raw.decode_content = True
                        shutil.copyfileobj(r.raw, f)
            pageInfo = soup.find("div", class_="containerTitle")
            pageInfoImg = soup.find("div", class_="container")
            name = pageInfo.find("h2", class_="headTwo")
            printName = name.text
            image = pageInfoImg.find("div", class_="containerImage")
            img = image.find("image",class_= "image")
            printImage = img.get("src")
            pictureURL = "http://projectabcd.com/" + printImage
            r = requests.get(pictureURL, headers={"User-Agent": "html"}, stream=True)
            if r.status_code == 200:
                with open(basename(printImage), "wb") as f:
                    r.raw.decode_content = True
                    shutil.copyfileobj(r.raw, f)
            
            pagetext = pageInfoImg.find("div", class_="containerText")
            description = pagetext.find("p", class_="words")
            printDescription = description.text
            fact = description.find_next_sibling("p")
            printFact = fact.text
             

            #creates the slide presentation if slide option 1 is choosen
            if(slideOption == 1 ):
                #creates the slides and sets layout preferences
                slide_layout = prs.slide_layouts[6]
                slide = prs.slides.add_slide(slide_layout)
                prs.slide_width = pptx.util.Inches(8)
                prs.slide_height = pptx.util.Inches(11)
                #places the logo on the slide
                logoHolder = slide.shapes.add_picture(basename(printLogo), pptx.util.Inches(7), pptx.util.Inches(0),width=pptx.util.Inches(1), height=pptx.util.Inches(1))
                #places the title on the slide
                titleBox = slide.shapes.add_textbox(pptx.util.Inches(2.5), pptx.util.Inches(.5), width = pptx.util.Inches(3), height = pptx.util.Inches(1))
                titleBoxtf = titleBox.text_frame
                title = titleBoxtf.add_paragraph()
                title.text = printName
                title.font.name = titleFont
                title.font.size = Pt(titleSize)
                #places the picture on the slide
                pictureHolder = prs.slides[i].shapes
                pictureHolder.add_picture(basename(printImage), pptx.util.Inches(2.5),pptx.util.Inches(2),width = pptx.util.Inches(3), height = pptx.util.Inches(4))
                #creates a textbox for the description and fun fact
                contentBox = slide.shapes.add_textbox(pptx.util.Inches(1), pptx.util.Inches(6), width = pptx.util.Inches(6), height = pptx.util.Inches(5))
                contentBoxtf = contentBox.text_frame
                contentBoxtf.word_wrap = True
                descriptionTitle = contentBoxtf.add_paragraph()
                descriptionTitle.font.name = textFont
                descriptionTitle.font.bold = True
                descriptionTitle.font.size = Pt(textSize)
                descriptionTitle.text = "Description: "
                descriptionParagraph = contentBoxtf.add_paragraph()
                descriptionParagraph.font.name = textFont
                descriptionParagraph.font.size = Pt(textSize)
                descriptionParagraph.text = printDescription
                FunFactTitle = contentBoxtf.add_paragraph()
                FunFactTitle.font.name = textFont
                FunFactTitle.font.bold = True
                FunFactTitle.font.size = Pt(textSize)
                FunFactTitle.text = "\nFun Fact:"
                FunFactParagraph = contentBoxtf.add_paragraph()
                FunFactParagraph.font.name = textFont
                FunFactParagraph.font.size = Pt(textSize)
                FunFactParagraph.text = printFact
            #creates the slide presentation if slide option 2 is chosen
            elif(slideOption == 2):
                #creates the slide layout preferences
                slide_layout = prs.slide_layouts[6]
                prs.slide_width = pptx.util.Inches(8)
                prs.slide_height = pptx.util.Inches(11)
                #creates a title page
                if(i == 0):
                    slide = prs.slides.add_slide(slide_layout)
                    titleBox = slide.shapes.add_textbox(pptx.util.Inches(2.5), pptx.util.Inches(1.5),width=pptx.util.Inches(3), height=pptx.util.Inches(2))
                    titleBoxtf = titleBox.text_frame
                    title = titleBoxtf.add_paragraph()
                    title.text = "Project abcd abdul"
                    title.font.size = Pt(50)
                    title.font.name = titleFont
                slide = prs.slides.add_slide(slide_layout)
                #places the picture to cover the whole slide
                pictureHolder = prs.slides[i+1].shapes
                pictureHolder.add_picture(basename(printImage), pptx.util.Inches(4), pptx.util.Inches(2), width=pptx.util.Inches(4), height=pptx.util.Inches(6))
                                #creates next slide
                #places the logo on the slide
                logoHolder = slide.shapes.add_picture(basename(printLogo), pptx.util.Inches(7), pptx.util.Inches(0),width=pptx.util.Inches(1), height=pptx.util.Inches(1))
                #places title on the slide
                titleBox = slide.shapes.add_textbox(pptx.util.Inches(2), pptx.util.Inches(1.5), width=pptx.util.Inches(2),height=pptx.util.Inches(1))
                titleBoxtf = titleBox.text_frame
                title = titleBoxtf.add_paragraph()
                title.text = printName
                title.font.size = Pt(titleSize)
                title.font.name = titleFont
                #creates textbox for description and fun fact
                contentBox = slide.shapes.add_textbox(pptx.util.Inches(1), pptx.util.Inches(2), width=pptx.util.Inches(3),height=pptx.util.Inches(4))
                contentBoxtf = contentBox.text_frame
                contentBoxtf.word_wrap = True
                descriptionTitle = contentBoxtf.add_paragraph()
                descriptionTitle.font.name = textFont
                descriptionTitle.font.bold = True
                descriptionTitle.font.size = Pt(textSize)
                descriptionTitle.text = "Description: "
                descriptionParagraph = contentBoxtf.add_paragraph()
                descriptionParagraph.font.name = textFont
                descriptionParagraph.font.size = Pt(textSize)
                descriptionParagraph.text = printDescription
                FunFactTitle = contentBoxtf.add_paragraph()
                FunFactTitle.font.bold = True
                FunFactTitle.font.name = textFont
                FunFactTitle.font.size = Pt(textSize)
                FunFactTitle.text = "\nFun Fact:"
                FunFactParagraph = contentBoxtf.add_paragraph()
                FunFactParagraph.font.name = textFont
                FunFactParagraph.font.size = Pt(textSize)
                FunFactParagraph.text = printFact
                
            #creates the slide presentation if slide option 3 is chosen
            elif (slideOption == 3):
             #creates slide preferences
             slide_layout = prs.slide_layouts[6]
             prs.slide_width = pptx.util.Inches(8)
             prs.slide_height = pptx.util.Inches(11)
             slide2 = prs.slides.add_slide(slide_layout)
             #places picture to cover whole slide
             pictureHolder = prs.slides[pictureSlide].shapes
             pictureHolder.add_picture(basename(printImage), pptx.util.Inches(0), pptx.util.Inches(2),width=pptx.util.Inches(4), height=pptx.util.Inches(6))
             #creates next slide
             
             #place logo on the slide
             logoHolder = slide2.shapes.add_picture(basename(printLogo), pptx.util.Inches(7), pptx.util.Inches(0),width=pptx.util.Inches(1), height=pptx.util.Inches(1))
             #places the title
             titleBox = slide2.shapes.add_textbox(pptx.util.Inches(4), pptx.util.Inches(1.5),width=pptx.util.Inches(2), height=pptx.util.Inches(1))
             titleBoxtf = titleBox.text_frame
             title = titleBoxtf.add_paragraph()
             title.text = printName
             title.font.size = Pt(titleSize)
             title.font.name = titleFont
             #creates textbox for description and fun fact
             contentBox = slide2.shapes.add_textbox(pptx.util.Inches(4), pptx.util.Inches(2), width=pptx.util.Inches(3),height=pptx.util.Inches(4))
             contentBoxtf = contentBox.text_frame
             contentBoxtf.word_wrap = True
             descriptionTitle = contentBoxtf.add_paragraph()
             descriptionTitle.font.name = textFont
             descriptionTitle.font.bold = True
             descriptionTitle.font.size = Pt(textSize)
             descriptionTitle.text = "Description: "
             descriptionParagraph = contentBoxtf.add_paragraph()
             descriptionParagraph.font.name = textFont
             descriptionParagraph.font.size = Pt(textSize)
             descriptionParagraph.text = printDescription
             FunFactTitle = contentBoxtf.add_paragraph()
             FunFactTitle.font.bold = True
             FunFactTitle.font.name = textFont
             FunFactTitle.font.size = Pt(textSize)
             FunFactTitle.text = "\nFun Fact:"
             FunFactParagraph = contentBoxtf.add_paragraph()
             FunFactParagraph.font.name = textFont
             FunFactParagraph.font.size = Pt(textSize)
             FunFactParagraph.text = printFact
             pictureSlide = pictureSlide + 1
 
        test = output +".pptx"
        prs.save(test)
        return test
    test = buildPresentation()
    os.startfile(test)
    
    
elif(methods == "Excel"):
    df = pd.read_excel(r"C:\xampp\htdocs\projectabcd_admin\dresses.xlsx")

    
    def buildPresentation(df):
        
     print("Creating powerpoint slides.")
    
     with open(preference, "r") as f:
                slideOption = f.readline().split("= ")
                slideOption = int(slideOption[1])
                textFont = f.readline().split("= ")
                textFont = textFont[1]
                titleFont = f.readline().split("= ")
                titleFont = titleFont[1]
                textSize = f.readline().split("= ")
                textSize = int(textSize[1])
                titleSize = f.readline().split("= ")
                titleSize = int(titleSize[1])
                prs = Presentation()
                presentationLength = len(all_pages)
                
                
     for i in range(presentationLength): 
   
                slide = prs.slides.add_slide(prs.slide_layouts[6]) 
                slide2 = prs.slides.add_slide(prs.slide_layouts[6])
                
                prs.slide_width = pptx.util.Inches(8)
                prs.slide_height = pptx.util.Inches(11)        
                
                contentBox = slide2.shapes.add_textbox(pptx.util.Inches(1), pptx.util.Inches(2), width=pptx.util.Inches(6),height=pptx.util.Inches(7))
                titleBox = slide2.shapes.add_textbox(pptx.util.Inches(4), pptx.util.Inches(.5),width=pptx.util.Inches(4), height=pptx.util.Inches(1))
                
                titleBoxtf = titleBox.text_frame
                title = titleBoxtf.add_paragraph()
                title.font.name = titleFont
                title.font.size = Pt(titleSize)
                title.font.name = titleFont
                title.text = str(df['name'][i]) 
                
                contentBoxtf = contentBox.text_frame
                contentBoxtf.word_wrap = True
                descriptionTitle = contentBoxtf.add_paragraph()
                descriptionTitle.font.name = textFont
                descriptionTitle.font.bold = True
                descriptionTitle.font.size = Pt(textSize)
                descriptionTitle.text = "Description: "
                descriptionParagraph = contentBoxtf.add_paragraph()
                descriptionParagraph.font.name = textFont
                descriptionParagraph.font.size = Pt(textSize)
                descriptionParagraph.text = (df['description'][i])
                FunFactTitle = contentBoxtf.add_paragraph()
                FunFactTitle.font.bold = True
                FunFactTitle.font.name = textFont
                FunFactTitle.font.size = Pt(textSize)
                FunFactTitle.text = "\nFun Fact:"
                FunFactParagraph = contentBoxtf.add_paragraph()
                FunFactParagraph.font.name = textFont
                FunFactParagraph.font.size = Pt(textSize)
                FunFactParagraph.text = (df['did_you_know'][i])
                
                
                
                image_url = (df['image_url'][i]) 
                if image_url:
                    image_path = os.path.basename(image_url)
                    slide.shapes.add_picture(image_path, pptx.util.Inches(0), pptx.util.Inches(0),width=pptx.util.Inches(8), height=pptx.util.Inches(11))
                
     test = output +".pptx"
     prs.save(test)
     return test

    test = buildPresentation(df)
    os.startfile(test)

elif(methods == "Database"):
    conn = mysql.connector.connect(user='root', host='localhost', database='abcd_dress-500')
    cursor = conn.cursor()


    sql = "SELECT * FROM dresses"
    cursor.execute(sql)
    result = cursor.fetchall()

    def buildPresentation(data):
        print("Creating PowerPoint slides.")

        with open("preferences.txt", "r") as f:
            slideOption = f.readline().split("= ")
            slideOption = int(slideOption[1])
            textFont = f.readline().split("= ")
            textFont = textFont[1]
            titleFont = f.readline().split("= ")
            titleFont = titleFont[1]
            textSize = f.readline().split("= ")
            textSize = int(textSize[1])
            titleSize = f.readline().split("= ")
            titleSize = int(titleSize[1])

        prs = Presentation()
        presentationLength = len(all_pages)
    

        for i in range(0,presentationLength):
        
            slide = prs.slides.add_slide(prs.slide_layouts[6]) 
            slide2 = prs.slides.add_slide(prs.slide_layouts[6])
        
            prs.slide_width = pptx.util.Inches(8)
            prs.slide_height = pptx.util.Inches(11)        
        
            contentBox = slide2.shapes.add_textbox(pptx.util.Inches(1), pptx.util.Inches(2), width=pptx.util.Inches(6),height=pptx.util.Inches(7))
            titleBox = slide2.shapes.add_textbox(pptx.util.Inches(4), pptx.util.Inches(.5),width=pptx.util.Inches(4), height=pptx.util.Inches(1))
            
            titleBoxtf = titleBox.text_frame
            title = titleBoxtf.add_paragraph()
            title.font.name = titleFont
            title.font.size = Pt(titleSize)
            title.font.name = titleFont
            title.text = str(data[i][1]) 
        
            contentBoxtf = contentBox.text_frame
            contentBoxtf.word_wrap = True
            descriptionTitle = contentBoxtf.add_paragraph()
            descriptionTitle.font.name = textFont
            descriptionTitle.font.bold = True
            descriptionTitle.font.size = Pt(textSize)
            descriptionTitle.text = "Description: "
            descriptionParagraph = contentBoxtf.add_paragraph()
            descriptionParagraph.font.name = textFont
            descriptionParagraph.font.size = Pt(textSize)
            descriptionParagraph.text = str(data[i][2])
            FunFactTitle = contentBoxtf.add_paragraph()
            FunFactTitle.font.bold = True
            FunFactTitle.font.name = textFont
            FunFactTitle.font.size = Pt(textSize)
            FunFactTitle.text = "\nFun Fact:"
            FunFactParagraph = contentBoxtf.add_paragraph()
            FunFactParagraph.font.name = textFont
            FunFactParagraph.font.size = Pt(textSize)
            FunFactParagraph.text = str(data[i][3])
        
        
        
            image_url = data[ i][8]  
            if image_url:
                image_path = os.path.basename(image_url)
                slide.shapes.add_picture(image_path, pptx.util.Inches(0), pptx.util.Inches(0),width=pptx.util.Inches(8), height=pptx.util.Inches(11))
            
        test = output +".pptx"
        prs.save(test)
        return test
    test = buildPresentation(result)
    os.startfile(test)