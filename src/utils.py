import os
import re
# import aspose.slides as slides
# import aspose.pydrawing as drawing

regex = r"</g><text.*>.*Aspose.*</text>"


def process_header(title: str):
    index = title.find(":")
    return title[index + 1 :]


def check_valid_image(image_path: str):
    for extension in ["BMP", "GIF", "JPEG", "PNG", "TIFF", "WMF"]:
        if image_path.endswith(extension.lower()):
            return True
    return False


def change_name_if_duplicated(init_name):
    """
    Check name, add index to it if duplicated, to avoid overwriting an existed file.
    Only work with pptx file.
    """
    if ".pptx" not in init_name:
        print("Not a pptx file, cannot check name duplication.")
        return None
    if os.path.exists(init_name):
        i = 1
        while True:
            # print('name exists, making newname...')
            if i > 1:
                init_name = re.sub(r"\((\d+)\)", lambda match: "(" + str(int(match.group(1)) + 1) + ")", init_name)
            else:
                try:
                    init_name = init_name.replace(".pptx", f" ({i}).pptx")
                except:
                    print("Cannot change name")
                    return None

            if not os.path.exists(init_name):
                break
            i += 1
    new_name = init_name
    return new_name


def get_layout_id(presentation):
    layout_id = 1
    while True:
        try:
            slide_layout = presentation.slide_layouts[layout_id]
        except:
            print("Template does not have any title that is suitable.")
            return None
        slide = presentation.slides.add_slide(slide_layout)
        shape = slide.placeholders[0]
        if shape.top / presentation.slide_height < 0.2:
            # print('top/totalheight = ', shape.top/presentation.slide_height)
            break
        layout_id += 1
    return layout_id


# def convert_pptx_to_svg(pptx_file, output_folder):
#     with slides.Presentation(pptx_file) as presentation:
#         for slide in presentation.slides:
#             svg_file_name = os.path.join(output_folder, "slide_{0}.svg".format(str(slide.slide_number)))
#             with open(svg_file_name, "wb") as file:
#                 slide.write_as_svg(file)

#             with open(svg_file_name, encoding="utf-8") as f:
#                 data = f.readlines()

#             for i, d in enumerate(data):
#                 result = re.findall(regex, d)
#                 if len(re.findall(regex, d)) > 0:
#                     for r in result:
#                         data[i] = data[i].replace(r[4:], "")

#             with open(svg_file_name, "w", encoding="utf-8") as f:
#                 f.writelines(data)
