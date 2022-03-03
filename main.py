import os

import pptx.util
from pptx import Presentation
from pptx.util import Pt
from bs4 import BeautifulSoup
import requests


def buildPresentation():
    URL = "https://projectabcd.com/display_the_dress.php?id=1"
    page = requests.get(URL, headers={"User-Agent": "html"})
    soup = BeautifulSoup(page.content, "html.parser")
    info = soup.find("div", class_="container")
    name = info.find("h2", class_="head")
    printName = name.text
    # image = info.find("img")
    # printImage = image.get("src")
    element = info.p
    description = info.find("p", class_="words")
    printDes = description.text
    fact = description.find_next_sibling("p")
    printFact = fact.text
    printWords = printDes + printFact

    prs = Presentation()
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    prs.slide_width = pptx.util.Inches(8)
    prs.slide_height = pptx.util.Inches(11)
    #title = slide.shapes.title.text = printName
    #placeholder = slide.placeholders[1]
    #picture = placeholder.insert_picture("Slide1.jpg")
    #words = slide.placeholders[2]
    #words.text = printWords
    txBox = slide.shapes.add_textbox(pptx.util.Inches(2.5), pptx.util.Inches(.5), width = pptx.util.Inches(3), height = pptx.util.Inches(1))
    tf = txBox.text_frame
    p = tf.add_paragraph()
    p.text = printName
    p.font.size = Pt(48)
    holder = prs.slides[0].shapes
    picture = holder.add_picture("Slide1.jpg", pptx.util.Inches(2.5),pptx.util.Inches(2),width = pptx.util.Inches(3), height = pptx.util.Inches(4))
    txBox2 = slide.shapes.add_textbox(pptx.util.Inches(1), pptx.util.Inches(6), width = pptx.util.Inches(6), height = pptx.util.Inches(5))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True
    p2 = tf2.add_paragraph()
    p2.font.bold = True
    p2.font.size = Pt(16)
    p2.text = "Description: "
    p3 = tf2.add_paragraph()
    p3.font.size = Pt(16)
    p3.text = printDes
    p4 = tf2.add_paragraph()
    p4.font.bold = True
    p4.font.size = Pt(16)
    p4.text = "\nFun Fact:"
    p5 = tf2.add_paragraph()
    p5.font.size = Pt(16)
    p5.text = printFact
    test = "test.pptx"
    prs.save(test)
    return test


test = buildPresentation()
os.startfile(test)

