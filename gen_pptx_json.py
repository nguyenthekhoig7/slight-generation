import json
import collections.abc
from pptx import Presentation
from pptx.util import Inches
import re
import os

DATA_FOLDER = r"data"
REPONSE_FILE = os.path.join(DATA_FOLDER, "output2.txt")
TEMPLATE_PPTX = os.path.join(DATA_FOLDER, "samples_template.pptx")
OUTPUT_PPTX = os.path.join(DATA_FOLDER, "output_slide_json2.pptx")

with open(REPONSE_FILE, "r") as f:
    response = f.read()


# get json data from text response
match = re.search(r"{(.*?)]\n}", response, re.DOTALL)
if match:
    jsontext = match.group(0)
else:
    print("Cannot find json in the response")


# create presentation (from template, or new blank one)
try:
    slides_json = json.loads(jsontext)
except:
    print("Cannot extract json from text, here is the original text: \n", jsontext)

try:
    prs = Presentation(TEMPLATE_PPTX)
except:
    prs = Presentation()

# delete all existed slides
for i in range(len(prs.slides) - 1, -1, -1):
    rId = prs.slides._sldIdLst[i].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[i]

# # Analyze the template's layouts
# print('About this template:')
# print("number of layouts: ", len(prs.slide_layouts))
# for i in range(len(prs.slide_layouts)):
#     layout = prs.slide_layouts[i]
#     print(f'layout #{i}')
#     for shape in layout.placeholders:
#         print(shape.name, '  ***  ', shape.placeholder_format.type)


# create slides
the_key = list(slides_json.keys())[0]
for slide in slides_json[the_key]:
    print(slide)
    header, content = (slide["header"], slide["content"])
    print(f"Header: {header},\ncontent: {content}")

    slide_layout = prs.slide_layouts[3]
    new_slide = prs.slides.add_slide(slide_layout)

    title = new_slide.shapes.title
    title.text = header

    left, top, width, height = (Inches(1.5), Inches(1.5), Inches(6), Inches(6))
    textbox = new_slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    text_frame.text = content
    text_frame.fit_text(max_size=22)
    text_frame.word_wrap = True


prs.save(OUTPUT_PPTX)
