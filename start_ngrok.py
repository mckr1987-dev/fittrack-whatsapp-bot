import ngrok
import time
import os
from dotenv import load_dotenv

load_dotenv()

NGROK_TOKEN = os.environ["NGROK_TOKEN"]    # set this in your .env file

print("Starting ngrok tunnel...")
listener = ngrok.forward(5000, authtoken=NGROK_TOKEN)
print("")
print("Public URL: " + listener.url())
print("Webhook URL: " + listener.url() + "/webhook")
print("")
print("Keep this window open while testing!")

try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    print("Stopped.")
