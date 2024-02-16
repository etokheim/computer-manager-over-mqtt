import paho.mqtt.client as mqtt
import ctypes
import time
import win32api
import json
from dotenv import load_dotenv
import os
import subprocess

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
class EntityMqtt():
	def __init__(self, topic, payload, client, state = None):
		self.topic = topic
		self.payload = payload
		self.state = state
		self.client = client
		self.payload_on = payload["payload_on"]
		self.payload_off = payload["payload_off"]

	def activate(self):
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
		print(f"Publishing {self.topic} state ({self.state})")

		# If these messages are published with a RETAIN flag, the MQTT switch will receive an instant state update after subscription, and will start with the correct state. Otherwise, the initial state of the switch will be unknown. A MQTT device can reset the current state to unknown using a None payload.
		client.publish(self.payload["state_topic"], self.state, retain=True)

	def subscribe_to_mqtt(self):
		client.subscribe(topic = self.payload["command_topic"], qos = 1)
		
	def on_message(self, client, userdata, payload):
		print(f"on_message not implemented for topic: {self.topic}") 

class EntityDisplay(EntityMqtt):
	def __init__(self, topic, payload, client, state = None):
		super().__init__(topic, payload, client, state)
		self.activate()

	def on_message(self, client, userdata, payload):
		# Callback function to handle incoming MQTT messages
		print(f"New message. Client: {client}, userdata: {userdata}, payload: {payload}") 
		
		# Check the payload and perform corresponding actions
		if payload == self.payload_on:
			self.turn_display_on()
		elif payload == self.payload_off:
			self.turn_display_off()

	# Function to turn on the display
	def turn_display_on(self):
		print("Turn on!")

		# Simulate a key press to wake up the display (Control key)
		win32api.keybd_event(0x11, 0, 0, 0)  # Control key down
		win32api.keybd_event(0x11, 0, 0x0002, 0)  # Control key up
		
		self.state = self.payload_on
		self.publish_state()

	# Function to turn off the display
	def turn_display_off(self):
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

		self.state = self.payload_off
		self.publish_state()

class EntityDarkMode(EntityMqtt):
	def __init__(self, topic, payload, client, state = None):
		super().__init__(topic, payload, client, state)
		self.state = self.get_dark_mode_state()
		print(f"Is dark mode enabled?: {self.state}")

		self.activate()

	def on_message(self, client, userdata, payload):
		# Callback function to handle incoming MQTT messages
		print(f"New message. Client: {client}, userdata: {userdata}, payload: {payload}") 
		
		# Check the payload and perform corresponding actions
		if payload == self.payload_on:
			self.enable_dark_mode()
		elif payload == self.payload_off:
			self.disable_dark_mode()

	def enable_dark_mode(self):
		try:
			subprocess.run(["reg", "add", "HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "/v", "AppsUseLightTheme", "/t", "REG_DWORD", "/d", "0", "/f"], check=True)
			print("Dark mode enabled.")

			self.state = self.payload_on
			self.publish_state()
		except Exception as e:
			print(f"Error enabling dark mode: {e}")

	def disable_dark_mode(self):
		try:
			subprocess.run(["reg", "add", "HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "/v", "AppsUseLightTheme", "/t", "REG_DWORD", "/d", "1", "/f"], check=True)
			print("Dark mode disabled.")

			self.state = self.payload_off
			self.publish_state()
		except Exception as e:
			print(f"Error disabling dark mode: {e}")
			
	def get_dark_mode_state(self):
		try:
			result = subprocess.run(["reg", "query", "HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "/v", "AppsUseLightTheme"], capture_output=True, text=True, check=True)
			return self.payload_on if ("0x1" not in result.stdout) else self.payload_off
		except Exception as e:
			print(f"Error checking dark mode status: {e}")
			return None

topic_prefix = f"homeassistant/switch/{COMPUTER_NAME}"

MQTT_ENTITIES[f"{topic_prefix}/display"] = EntityDisplay(
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

MQTT_ENTITIES[f"{topic_prefix}/dark_mode"] = EntityDarkMode(
	topic = f"{topic_prefix}/dark_mode",

	payload = {
		"name": "Dark Mode",
		"command_topic": f"{topic_prefix}/dark_mode/set",
		"state_topic": f"{topic_prefix}/dark_mode/state",
		"unique_id": COMPUTER_NAME + "_dark_mode",
		"payload_on": "ON",
		"payload_off": "OFF",
		"device": {
			"identifiers": [COMPUTER_NAME],
			"name": COMPUTER_NAME
		}
	},
	
	client = client
)

print(f"MQTT_ENTITIES: {MQTT_ENTITIES}")
	
# Main loop to keep the program running and handle MQTT messages
try:
	client.loop_start()
	while True:
		time.sleep(1)  # Keep the program running
except KeyboardInterrupt:
	print("Exiting program.")
	client.disconnect()
	client.loop_stop()