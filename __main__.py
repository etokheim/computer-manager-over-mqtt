import paho.mqtt.client as mqtt
import ctypes
import time
import win32api
import json
from dotenv import load_dotenv
import os

load_dotenv()

DISPLAY_STATE = "ON"

COMPUTER_NAME = os.getenv("COMPUTERNAME")

HOMEASSISTANT_STATUS_TOPIC = "homeassistant/status" # TODO: Also subscribe to the status topic and republish discovery when getting events
CONFIGURATION_TOPIC = "homeassistant/switch/" + COMPUTER_NAME + "/display/config"
COMMAND_TOPIC = "homeassistant/switch/" + COMPUTER_NAME + "/display/set"
STATE_TOPIC = "homeassistant/switch/" + COMPUTER_NAME + "/display/state"
BROKER_ADDRESS = os.getenv("MQTT_BROKER_ADDRESS")
USERNAME = os.getenv("MQTT_USERNAME")
PASSWORD = os.getenv("MQTT_PASSWORD")

# Callback function to handle incoming MQTT messages
def on_message(client, userdata, message):
	payload = message.payload.decode("utf-8")
	
	print(f"New message. Client: {client}, userdata: {userdata}, payload: {payload}") 
	
	# Check the payload and perform corresponding actions
	if payload.lower() == "on":
		turn_display_on()
	elif payload.lower() == "off":
		turn_display_off()
	elif payload.lower() == "online":
		# This can be replaced by only sending one discovery message with the retain flag,
		# however, at first try it seemed unreliable.
		publish_discovery_payload()

# Function to turn on the display
def turn_display_on():
	print("Turn on!")

	# Simulate a key press to wake up the display (Control key)
	win32api.keybd_event(0x11, 0, 0, 0)  # Control key down
	win32api.keybd_event(0x11, 0, 0x0002, 0)  # Control key up
	
	global DISPLAY_STATE
	DISPLAY_STATE = "ON"
	publish_state()

# Function to turn off the display
def turn_display_off():
	print("Turn off!")
	
	ctypes.windll.user32.SendMessageTimeoutW(
		ctypes.windll.user32.GetForegroundWindow(),
		0x112,  # WM_SYSCOMMAND
		0xF170,  # SC_MONITORPOWER
		2,  # Turn monitor off
		0,
		5000,
		ctypes.pointer(ctypes.c_ulong())
	)
	
	global DISPLAY_STATE
	DISPLAY_STATE = "OFF"
	publish_state()

# Function to publish the discovery payload
def publish_discovery_payload():
	print("Publishing discovery payload")

	discovery_payload = {
		"name": "Display",
		"command_topic": "homeassistant/switch/" + COMPUTER_NAME + "/display/set",
		"state_topic": "homeassistant/switch/" + COMPUTER_NAME + "/display/state",
		"unique_id": COMPUTER_NAME + "_display",
		"payload_on": "ON",
		"payload_off": "OFF",
		"device": {
			"identifiers": [COMPUTER_NAME],
			"name": COMPUTER_NAME
		}
	}

	client.publish(CONFIGURATION_TOPIC, json.dumps(discovery_payload), retain=False)
	
	# From the docs:
	# After the configs have been published, the state topics will need an update, so they need to be republished.
	# https://www.home-assistant.io/integrations/mqtt/#use-the-birth-and-will-messages-to-trigger-discovery
	publish_state()

def publish_state():
	print("Publish state")

	global DISPLAY_STATE
	state_payload = DISPLAY_STATE

	# If these messages are published with a RETAIN flag, the MQTT switch will receive an instant state update after subscription, and will start with the correct state. Otherwise, the initial state of the switch will be unknown. A MQTT device can reset the current state to unknown using a None payload.
	client.publish(STATE_TOPIC, state_payload, retain=True)

# Set up the MQTT client with authentication
client = mqtt.Client()
client.username_pw_set(USERNAME, PASSWORD)
client.on_message = on_message
client.connect(BROKER_ADDRESS, 1883, 60)
client.subscribe([(COMMAND_TOPIC, 1), (HOMEASSISTANT_STATUS_TOPIC, 1)])

# Publish the discovery payload
publish_discovery_payload()

# Main loop to keep the program running and handle MQTT messages
try:
	client.loop_start()
	while True:
		time.sleep(1)  # Keep the program running
except KeyboardInterrupt:
	print("Exiting program.")
	client.disconnect()
	client.loop_stop()