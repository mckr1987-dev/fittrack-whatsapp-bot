import ngrok
import time

NGROK_TOKEN = "3FeuRXcZLXR9gkffkK5NzMa6Ktj_316aubTUiwND5aCUumnNv"    # paste your token here

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