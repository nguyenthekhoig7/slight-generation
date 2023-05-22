import poe
import logging
import re


def create_content_json(reponse_file: str):
    with open(reponse_file, "r") as f:
        response = f.read()
    match = re.search(r"{(.*?)]\n}", response, re.DOTALL)
    if match:
        content_json = match.group(0)
    else:
        return None

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
        + ". Each slide should have a  {{header}}, {{content}}. The final slide should be a list of discussion questions. Return as JSON."
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
