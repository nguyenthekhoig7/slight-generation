import poe
import logging
import re
import json
import pypandoc

def create_content_json(reponse_file: str):
    def _create_content_from_json(response):
        match = re.search(r"{(.*?)]\n}", response, re.DOTALL)
        if match: # response has json inside
            content_json = match.group(0)
            try:
                content_json = json.loads(content_json)
                return content_json
            except:
                return None
    def _create_content_from_python_code(response):
        match_pycode = re.search(r"```python(.*?)\)\n```", response, re.DOTALL)
        if not match_pycode:
            return None
        py_code = match_pycode.group(0)
        
        match_content = re.search(r"content =(.*?)]", py_code, re.DOTALL)
        if not match_content:
            return None
        content_str = match_content.group(0)
        contents = content_str.split(r"=")[1]
        contents_list = re.findall("\"(.*)\",", contents)

        match_header = re.search(r"header =(.*?)]", response, re.DOTALL)
        if not match_header:
            return None
        headers_str = match_header.group(0)
        headers_list = re.findall("\"Slide \d+: (.+)\"", headers_str)

        slides_json = {}
        slides_json['slides'] = []

        for header, content in zip(headers_list, contents_list):
            pair = {}
            pair['header'] = header
            pair['content'] = content
            slides_json['slides'].append(pair)

        return slides_json

    with open(reponse_file, "r") as f:
        response = f.read()
        
    content_json = _create_content_from_json(response)
    if content_json is None:
        content_json = _create_content_from_python_code(response)

    return content_json


def create_query(topic: str, n_slides: int = 10, n_words_per_slide: int = 55):
    query = """"{
    "input_text": "[[QUERY]]",
    "output_format": "json",
    "json_structure": {
        "slides":"{{presentation_slides}}"
       }
    }"""

    topic_query = (
        f"Generate a {n_slides} slide presentation for the topic. Produce {n_words_per_slide-5} to {n_words_per_slide+5} words per slide. "
        + topic
        + ". Each slide should have a  {{header}}, {{content}}. The final slide should be a list of discussion questions. Return as JSON, only JSON, not the code to generate JSON."
    )

    query = query.replace("[[QUERY]]", topic_query)
    return query


def create_query_read_document(docu_file: str, n_slides: int = 10, n_words_per_slide: int = 55):
    def _get_document(docu_file: str):
        assert 'doc' in docu_file
        docu_txt_file = docu_file.replace("docx", "txt")
        output = pypandoc.convert_file(docu_file, 'plain', outputfile=docu_txt_file)
        assert output == ""

        with open(docu_txt_file, 'r') as f:
            docu = f.read()
            if len(docu.split()) > 500:
                print("Warning: document input larger than 500 words, reduce it next time to have better performance.")
        return docu
    docu = _get_document(docu_file)

    query = """"{
    "input_text": "[[QUERY]]",
    "output_format": "json",
    "json_structure": {
        "slides":"{{presentation_slides}}"
       }
    }"""
    topic_query = (
        f"Generate a {n_slides} slide presentation from the document provided. Produce {n_words_per_slide-5} to {n_words_per_slide+5} words per slide. "
        + ". Each slide should have a  {{header}}, {{content}}. The first slide should only contain the short title. The final slide should be a list of discussion questions. Return as JSON, only JSON, not the code to generate JSON."
        + " Here is the document: \n"
        + docu
    )
    query = query.replace("[[QUERY]]", topic_query)
    return query


def query_from_API(query: str, token: str, output_path: str, bot_name="chinchilla"):
    try:
        poe.logger.setLevel(logging.INFO)
        client = poe.Client(token)

        with open(output_path, "w") as f:
            for chunk in client.send_message(bot_name, query, with_chat_break=True):
                print(chunk["text_new"], end="", flush=True)
                f.write(chunk["text_new"])

        # delete the 3 latest messages, including the chat break
        client.purge_conversation(bot_name, count=3)
    except:
        return False

    return True
