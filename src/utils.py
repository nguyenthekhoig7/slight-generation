def process_header(title: str):
    index = title.find(":")
    return title[index + 1 :]


def check_valid_image(image_path: str):
    for extension in ["BMP", "GIF", "JPEG", "PNG", "TIFF", "WMF"]:
        if image_path.endswith(extension.lower()):
            return True
    return False
