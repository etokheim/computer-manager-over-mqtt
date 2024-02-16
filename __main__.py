import paho.mqtt.client as mqtt
import ctypes
import time
import win32api
import json
from dotenv import load_dotenv
import os
import types

load_dotenv()

DISPLAY_STATE = "ON"

COMPUTER_NAME = os.getenv("COMPUTERNAME")

HOMEASSISTANT_STATUS_TOPIC = "homeassistant/status"
STATE_TOPIC = "homeassistant/switch/" + COMPUTER_NAME + "/display/state"
BROKER_ADDRESS = os.getenv("MQTT_BROKER_ADDRESS")
USERNAME = os.getenv("MQTT_USERNAME")
PASSWORD = os.getenv("MQTT_PASSWORD")
MQTT_ENTITIES = {}

def distribute_message(client, userdata, message):
	"""Route messages to the correct entities"""
	# invoke on_message on all MqttEntities
	# First, add listener to home assistant status topic!
	payload = message.payload.decode("utf-8")
	topic = message.topic
	topic_stripped = topic.removesuffix("/set")

	print(f"Got message from topic: {topic_stripped}")
	
	MQTT_ENTITIES[topic_stripped].on_message(client, userdata, payload)

# Set up the MQTT client with authentication
client = mqtt.Client()
client.on_message = distribute_message
client.username_pw_set(USERNAME, PASSWORD)
client.connect(BROKER_ADDRESS, 1883, 60)
#client.subscribe(topic = HOMEASSISTANT_STATUS_TOPIC, qos = 1)
		#elif payload.lower() == "online":
		#	# This can be replaced by only sending one discovery message with the retain flag,
		#	# however, at first try it seemed unreliable.
		#	self.publish_discovery_payload()

# Publish the discovery payload
class MqttEntity():
	def __init__(self, topic, payload, state, client):
		self.topic = topic
		self.payload = payload
		self.state = state
		self.client = client
		self.payload_on = payload["payload_on"]
		self.payload_off = payload["payload_off"]

		self.subscribe_to_mqtt()
		self.publish_discovery_payload()

	def publish_discovery_payload(self):
		print(f"Publishing discovery payload for topic: {self.topic}")

		discovery_message = json.dumps(self.payload)
		client.publish(f"{self.topic}/config", discovery_message, qos=1, retain=False)
		
		# From the docs:
		# After the configs have been published, the state topics will need an update, so they need to be republished.
		# https://www.home-assistant.io/integrations/mqtt/#use-the-birth-and-will-messages-to-trigger-discovery
		self.publish_state()
		
	def publish_state(self):
		print(f"Publishing {self.topic} state")

		# If these messages are published with a RETAIN flag, the MQTT switch will receive an instant state update after subscription, and will start with the correct state. Otherwise, the initial state of the switch will be unknown. A MQTT device can reset the current state to unknown using a None payload.
		client.publish(self.payload["state_topic"], self.state, retain=True)

	def subscribe_to_mqtt(self):
		client.subscribe(topic = self.payload["command_topic"], qos = 1)
		
	def on_message(self, client, userdata, payload):
		print(f"on_message not implemented for topic: {self.topic}") 

topic_prefix = f"homeassistant/switch/{COMPUTER_NAME}"

MQTT_ENTITIES[f"{topic_prefix}/display"] = MqttEntity(
	topic = f"{topic_prefix}/display",

	payload = {
		"name": "Display",
		"command_topic": f"{topic_prefix}/display/set",
		"state_topic": f"{topic_prefix}/display/state",
		"unique_id": COMPUTER_NAME + "_display",
		"payload_on": "ON",
		"payload_off": "OFF",
		"device": {
			"identifiers": [COMPUTER_NAME],
			"name": COMPUTER_NAME
		}
	},
	
	state = "ON",

	client = client
)

def display_on_message(self, client, userdata, payload):
	# Callback function to handle incoming MQTT messages
	print(f"New message. Client: {client}, userdata: {userdata}, payload: {payload}") 
	return
	
	# Check the payload and perform corresponding actions
	if payload == self.payload_on:
		turn_display_on()
	elif payload == self.payload_off:
		turn_display_off()

entity_display = MQTT_ENTITIES[f"{topic_prefix}/display"]
entity_display.on_message = types.MethodType(display_on_message, entity_display) # types.MethodType to get access to self within the appended method.
print(f"MQTT_ENTITIES: {MQTT_ENTITIES}")


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
	
def enable_dark_mode():
	try:
		ctypes.windll.dwmapi.DwmSetPreferredColorization(2, 0, 0, 0)  # Enable dark mode
		print("Dark mode enabled.")
	except Exception as e:
		print(f"Error enabling dark mode: {e}")

def disable_dark_mode():
	try:
		ctypes.windll.dwmapi.DwmSetPreferredColorization(1, 0, 0, 0)  # Disable dark mode
		print("Dark mode disabled.")
	except Exception as e:
		print(f"Error disabling dark mode: {e}")


# Main loop to keep the program running and handle MQTT messages
try:
	client.loop_start()
	while True:
		time.sleep(1)  # Keep the program running
except KeyboardInterrupt:
	print("Exiting program.")
	client.disconnect()
	client.loop_stop()