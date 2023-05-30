import os
import json
import random
from PIL import Image
import collections.abc
from pptx.util import Inches
from pptx import Presentation
from api_key import POE_API_KEY

from src.utils import *
from src.image_download import Downloader
from src.text_gen import *
import re
import time

DATA_FOLDER = r"data"
FONT_FOLDER = r"fonts"
TEMPLATE_PPTX = os.path.join(DATA_FOLDER, "template.pptx")
CHOSEN_FONT = os.path.join(FONT_FOLDER, "Calibri Regular.ttf")

downloader = Downloader()

print("############################# SLIGHT #############################")
input_mode = int(
    input(
        """Select how you want to create: 
    1. Upload 1 docx document.
    2. Enter a topic.
    Your answer [1/2]: """
    )
)
mode = {1: "document", 2: "topic"}

if mode[input_mode] == "document":
    docu_file = input("Enter the document(docx/pdf) path:\n >>> ")
    while True:
        if ".doc" in docu_file or '.pdf' in docu_file:
            break
        print("Only accept .doc or .docx files. Please try again.")
        docu_file = input("Enter the document(docx/pdf) path:\n >>> ")
    text_query, output_txt_path = create_query_read_document(docu_file=docu_file)
    # output_txt_path = os.path.join("data", re.sub(r"\.[a-zA-Z0-9]+$", ".txt", docu_file))  # .docx and .doc --> .txt

    with open(output_txt_path, "r") as f:
        docu = f.read()
    get_title_query = (
        "Summarize the following document into a short title. Your response should contain only a short title of less than 10 words, nothing other than that, a less-than-10-word title. Here is the document: \n"
        + docu
    )
    topic = query_from_API(query=get_title_query, token=POE_API_KEY)

else:  # mode topic
    topic = input("What do you want to make a presentation about? \n >>> ")
    text_query = create_query(topic, n_slides=10, n_words_per_slide=70)
    output_txt_path = os.path.join("data", topic.replace(" ", "_") + ".txt")

st_time = time.time()
success = query_API__save_to_file(query=text_query, token=POE_API_KEY, output_path=output_txt_path)

if success:
    print(f"Successfully generate content about {topic}.")
else:
    print("Cannot query from API. Please try again")

response = read_response_file(output_txt_path)
content_json = create_content_json(response)

if content_json is None:
    print("Cannot extract json from text. Cannot create a presentation. Stopped.")
    exit()

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
    image_query = (header + topic).replace(" ", "_")
    image = None
    if "Introduction" not in image_query:
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
        if i < 1:
            continue
        sp = place_holder.element
        sp.getparent().remove(sp)

    if image:
        w, h = image.size
        try:
            if w > h:
                    picture = slide.shapes.add_picture(path_to_image, Inches(6), Inches(2.5), width=Inches(3.8))
            else:
                picture = slide.shapes.add_picture(path_to_image, Inches(6), Inches(2.5), height=Inches(5))
        except Exception as e:
            print('Cannot add picture. ', e)
    left, top, width, height = Inches(1), Inches(2.5), Inches(5), Inches(5)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    text_frame.text = content
    text_frame.fit_text(font_file=CHOSEN_FONT, max_size=22)
    text_frame.word_wrap = True

output_pptx_path = output_txt_path.replace("txt", "pptx")
prs.save(output_pptx_path)
print(f'Presentation saved to {output_pptx_path}')

end_time = time.time()

duration = end_time - st_time
print(f'time consumed: %.dm%ds' % (duration/60, duration%60))