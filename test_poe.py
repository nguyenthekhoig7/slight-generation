import poe
import logging
import sys
import json


OUTPUT_CHATBOT = r"./data/output2.txt"


query_json = """"{
    "input_text": "[[QUERY]]",
    "output_format": "json",
    "json_structure": {
        "slides":"{{presentation_slides}}"
       }
    }"""

presentation_title = input("What do you want to make a presentation about? \n >>> ")
question = (
    "Generate a 10 slide presentation for the topic. Produce 50 to 60 words per slide. "
    + presentation_title
    + ". Each slide should have a  {{header}}, {{content}}. The final slide should be a list of discussion questions. Return as JSON."
)

prompt = query_json.replace("[[QUERY]]", question)


# send a message and immediately delete it
token = sys.argv[1]

poe.logger.setLevel(logging.INFO)
client = poe.Client(token)

with open(OUTPUT_CHATBOT, "w") as f:
    for chunk in client.send_message("chinchilla", prompt, with_chat_break=True):
        print(chunk["text_new"], end="", flush=True)
        f.write(chunk["text_new"])


# delete the 3 latest messages, including the chat break
client.purge_conversation("chinchilla", count=3)
print(f"Results written to {OUTPUT_CHATBOT}")
