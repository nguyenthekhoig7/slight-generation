import os
from PIL import Image
import collections.abc
from typing import Optional
from fastapi import FastAPI
from pptx.util import Inches
from pptx import Presentation
from api_key import POE_API_KEY
from bing_image_downloader import downloader

from src.utils import *
from src.text_gen import *


DATA_FOLDER = r"data"
FONT_FOLDER = r"fonts"
IMAGE_FOLDER = os.path.join(r"images")
TEMPLATE_PPTX = os.path.join(DATA_FOLDER, "template.pptx")
CHOSEN_FONT = os.path.join(FONT_FOLDER, "Calibri Regular.ttf")

app = FastAPI()


@app.post("/generate/")
async def generate(topic: str, n_slides: Optional[int] = 10, n_words_per_slide: Optional[int] = 70):
    text_query = create_query(topic, n_slides=n_slides, n_words_per_slide=n_words_per_slide)
    response = query_from_API(query=text_query, token=POE_API_KEY)
    content_json = create_content_json(response)

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

    if not os.path.isdir(IMAGE_FOLDER):
        os.makedirs(IMAGE_FOLDER)

    # for item in content_json[key]:
    #     header, content = process_header(item["header"]), item["content"]
    #     image_query = (header + topic).replace(" ", "_")
    #     image = None
    #     if "Introduction" not in image_query:
    #         try:
    #             downloader.download(
    #                 image_query,
    #                 limit=1,
    #                 output_dir=IMAGE_FOLDER,
    #                 force_replace=False,
    #                 timeout=10,
    #                 verbose=False,
    #             )

    #             image_folder_path = os.path.join(IMAGE_FOLDER, image_query)
    #             image_path = os.path.join(image_folder_path, os.listdir(image_folder_path)[0])
    #             image = Image.open(image_path)
    #         except Exception as e:
    #             print(e)

    #     slide_layout = prs.slide_layouts[3]
    #     slide = prs.slides.add_slide(slide_layout)

    #     title = slide.shapes.title
    #     title.text = header

    #     for i, place_holder in enumerate(slide.placeholders):
    #         if i < 1:
    #             continue
    #         sp = place_holder.element
    #         sp.getparent().remove(sp)

    #     if image:
    #         w, h = image.size
    #         try:
    #             if w > h:
    #                 picture = slide.shapes.add_picture(image_path, Inches(6), Inches(2.5), width=Inches(3.8))
    #             else:
    #                 picture = slide.shapes.add_picture(image_path, Inches(6), Inches(2.5), height=Inches(5))
    #         except Exception as e:
    #             print("Cannot add picture. ", e)

    #     left, top, width, height = Inches(1), Inches(2.5), Inches(5), Inches(5)
    #     textbox = slide.shapes.add_textbox(left, top, width, height)
    #     text_frame = textbox.text_frame
    #     text_frame.text = content
    #     text_frame.fit_text(font_file=CHOSEN_FONT, max_size=22)
    #     text_frame.word_wrap = True

    # output_pptx_path = output_txt_path.replace("txt", "pptx")
    # output_pptx_path = change_name_if_duplicated(output_pptx_path)
    # prs.save(output_pptx_path)
    # print(f"Presentation saved to {output_pptx_path}")

    return content_json


from pydantic import BaseModel


class Item(BaseModel):
    name: str
    description: str | None = None
    price: float
    tax: float | None = None


@app.post("/items/")
async def create_item(item: Item):
    return item
