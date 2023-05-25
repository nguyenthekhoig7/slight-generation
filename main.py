import os
import json
import random
from PIL import Image
import collections.abc
from pptx.util import Inches
from pptx import Presentation
from api_key import API_KEY

from src.utils import *
from src.image_download import Downloader
from src.text_gen import *
import re

DATA_FOLDER = r"data"
FONT_FOLDER = r"fonts"
TEMPLATE_PPTX = os.path.join(DATA_FOLDER, "template.pptx")
CHOSEN_FONT = os.path.join(FONT_FOLDER, "Calibri Regular.ttf")

downloader = Downloader()

print("############################# SLIGHT #############################")
input_mode = int(input("""Select how you want to create: 
    1. Upload 1 docx document.
    2. Enter a topic.
    Your answer [1/2]: """))
mode = {1: "document", 2: "topic"}

if mode[input_mode] == "document":
    docu_file = input("Enter the document(docx) path:\n >>> ")
    text_query = create_query_read_document(docu_file=docu_file)
    output_txt_path = os.path.join("data", re.sub(r'\.docx*$', '.txt', docu_file)) # .docx and .doc --> .txt
else:
    topic = input("What do you want to make a presentation about? \n >>> ")
    text_query = create_query(topic)
    output_txt_path = os.path.join("data", topic.replace(" ", "_") + ".txt")

success = query_from_API(query=text_query, token=API_KEY, output_path=output_txt_path)

if success:
    print(f"Successfully generate content.")
else:
    print("Cannot query from API. Please try again")

content_json = create_content_json(output_txt_path)
if content_json is None:
    print("Cannot extract json from text. Please try again !!!")

try:
    prs = Presentation(TEMPLATE_PPTX)
    print(f"Use template from {TEMPLATE_PPTX}")
except:
    prs = Presentation()
    print(f"Cannot use template from {TEMPLATE_PPTX}. Creating a blank file.")

for i in range(len(prs.slides) - 1, -1, -1):
    rId = prs.slides._sldIdLst[i].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[i]

key = list(content_json.keys())[0]
for item in content_json[key]:
    header, content = process_header(item["header"]), item["content"]
    try:
        topic
    except:
        topic = " "
    image_query = (header + topic).replace(" ", "_")
    image = None
    if image_query != "Introduction":
        try:
            downloader.download(image_query, limit=20, timer=50)
            image_names = os.listdir(os.path.join("simple_images", image_query))
            while True:
                path_to_image = os.path.join(
                    "simple_images",
                    image_query,
                    image_names[random.randint(0, len(image_names) - 1)],
                )
                image = Image.open(path_to_image)
                if image.size != (80, 36) and check_valid_image(path_to_image):
                    break
            downloader.flush_cache()
        except:
            print("Cannot download image")

    slide_layout = prs.slide_layouts[3]
    slide = prs.slides.add_slide(slide_layout)

    title = slide.shapes.title
    title.text = header

    for i, place_holder in enumerate(slide.placeholders):
        if i < 2:
            continue
        sp = place_holder.element
        sp.getparent().remove(sp)

    if image:
        picture = slide.shapes.add_picture(path_to_image, Inches(1), Inches(1))
        # TODO: change picture size

    left, top, width, height = (Inches(1.5), Inches(1.5), Inches(6), Inches(6))
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    text_frame.text = content
    text_frame.fit_text(font_file=CHOSEN_FONT, max_size=22)
    text_frame.word_wrap = True

output_pptx_path = output_txt_path.replace("txt", "pptx")
prs.save(output_pptx_path)
