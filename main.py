import os
import shutil
from os.path import basename

import pptx.util
from pptx import Presentation
from pptx.util import Pt
from bs4 import BeautifulSoup
import requests





def buildPresentation():
   start = int(input("Enter the slide number you'd like to start with: "))
   end = int(input("Enter the slide number you'd like to end with: "))
   prs = Presentation()
   slide_layout = prs.slide_layouts[6]
   presentationLength = end - start + 1
   for i in range (0, presentationLength):
        URL = "https://projectabcd.com/display_the_dress.php?id=" + str(start + i)
        page = requests.get(URL, headers={"User-Agent": "html"})
        soup = BeautifulSoup(page.content, "html.parser")
        pageInfo = soup.find("div", class_="container")
        name = pageInfo.find("h2", class_="head")
        printName = name.text
        image = pageInfo.find("image")
        printImage = image.attrs["src"]
        pictureURL = "http://projectabcd.com/" + printImage
        r = requests.get(pictureURL, headers={"User-Agent": "html"}, stream=True)
        if r.status_code == 200:
             with open(basename(printImage), "wb") as f:
                  r.raw.decode_content = True
                  shutil.copyfileobj(r.raw, f)
        element = pageInfo.p
        description = pageInfo.find("p", class_="words")
        printDescription = description.text
        fact = description.find_next_sibling("p")
        printFact = fact.text
        printWords = printDescription + printFact

        slide = prs.slides.add_slide(slide_layout)
        prs.slide_width = pptx.util.Inches(8)
        prs.slide_height = pptx.util.Inches(11)
        titleBox = slide.shapes.add_textbox(pptx.util.Inches(2.5), pptx.util.Inches(.5), width = pptx.util.Inches(3), height = pptx.util.Inches(1))
        titleBoxtf = titleBox.text_frame
        title = titleBoxtf.add_paragraph()
        title.text = printName
        title.font.size = Pt(48)
        pictureHolder = prs.slides[i].shapes
        pictureHolder.add_picture(basename(printImage), pptx.util.Inches(2.5),pptx.util.Inches(2),width = pptx.util.Inches(3), height = pptx.util.Inches(4))
        contentBox = slide.shapes.add_textbox(pptx.util.Inches(1), pptx.util.Inches(6), width = pptx.util.Inches(6), height = pptx.util.Inches(5))
        contentBoxtf = contentBox.text_frame
        contentBoxtf.word_wrap = True
        descriptionTitle = contentBoxtf.add_paragraph()
        descriptionTitle.font.bold = True
        descriptionTitle.font.size = Pt(16)
        descriptionTitle.text = "Description: "
        descriptionParagraph = contentBoxtf.add_paragraph()
        descriptionParagraph.font.size = Pt(16)
        descriptionParagraph.text = printDescription
        FunFactTitle = contentBoxtf.add_paragraph()
        FunFactTitle.font.bold = True
        FunFactTitle.font.size = Pt(16)
        FunFactTitle.text = "\nFun Fact:"
        FunFactParagraph = contentBoxtf.add_paragraph()
        FunFactParagraph.font.size = Pt(16)
        FunFactParagraph.text = printFact
   test = "test.pptx"
   prs.save(test)
   return test

test = buildPresentation()
os.startfile(test)

