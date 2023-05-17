import poe
import logging
import sys


query_json = """"{
    "input_text": "[[QUERY]]",
    "output_format": "json",
    "json_structure": {
        "slides":"{{presentation_slides}}"
       }
    }"""

presentation_title = input("What do you want to make a presentation about? >")
question = "Generate a 10 slide presentation for the topic. Produce 50 to 60 words per slide. " + presentation_title + ".Each slide should have a  {{header}}, {{content}}. The final slide should be a list of discussion questions. Return as JSON."

prompt = query_json.replace("[[QUERY]]",question)



#send a message and immediately delete it
token = sys.argv[1]
poe.logger.setLevel(logging.INFO)
client = poe.Client(token)

message = prompt
print('fpt asking: ',message)
for chunk in client.send_message("chinchilla", message, with_chat_break=True):
  print(chunk["text_new"], end="", flush=True)

#delete the 3 latest messages, including the chat break
client.purge_conversation("chinchilla", count=3)