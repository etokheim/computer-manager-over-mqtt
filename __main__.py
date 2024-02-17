import paho.mqtt.client as mqtt
import ctypes
import time
import win32api
import json
from dotenv import load_dotenv
import os
import subprocess
import logging
import sys
import colorlog

# Load environment variables from .env
load_dotenv()

# Set up logging
logger = logging.getLogger("logger")

stdout = colorlog.StreamHandler(stream=sys.stdout)

fmt = colorlog.ColoredFormatter(
	"%(white)s%(asctime)s%(reset)s | %(blue)s%(filename)s%(reset)s | %(log_color)s(%(levelname)s)%(reset)s %(log_color)s%(message)s%(reset)s"
)

stdout.setFormatter(fmt)
logger.addHandler(stdout)

LOG_LEVEL = os.getenv("CMOM_LOG_LEVEL")

if LOG_LEVEL:
	if LOG_LEVEL == "critical":
		logger.setLevel(logging.CRITICAL)
	elif LOG_LEVEL == "error":
		logger.setLevel(logging.ERROR)
	elif LOG_LEVEL == "warning":
		logger.setLevel(logging.WARNING)
	elif LOG_LEVEL == "info":
		logger.setLevel(logging.INFO)
	elif LOG_LEVEL == "debug":
		logger.setLevel(logging.DEBUG)
else:
	logger.setLevel(logging.WARNING)

logger.info("Setting log level to: %s", LOG_LEVEL)

# Setup
HOMEASSISTANT_STATUS_TOPIC = "homeassistant/status"
COMPUTER_NAME = os.getenv("COMPUTERNAME")
BROKER_ADDRESS = os.getenv("MQTT_BROKER_ADDRESS")
USERNAME = os.getenv("MQTT_USERNAME")
PASSWORD = os.getenv("MQTT_PASSWORD")
MQTT_ENTITIES = {}

def distribute_message(client, userdata, message):
	"""Route messages to the entity matching the message topic"""
	payload = message.payload.decode("utf-8")
	topic = message.topic
	topic_stripped = topic.removesuffix("/set")

	# logger.info("Got message from topic: %s. Forwarding.", topic_stripped)
	
	try:
		MQTT_ENTITIES[topic_stripped].on_message(client, userdata, payload)
	except KeyError as error:
		if topic == HOMEASSISTANT_STATUS_TOPIC:
			logger.info("Got status message from Home Assistant: %s", payload)
			
			if payload == "online":
				for topic, entity in MQTT_ENTITIES.items():
					entity.publish_discovery_payload()
		else:
			raise KeyError(error)

# Set up the MQTT client with authentication
client = mqtt.Client()
client.on_message = distribute_message
client.username_pw_set(USERNAME, PASSWORD)
client.connect(BROKER_ADDRESS, 1883, 60)

# Subscribe to homeassistant's status messages
# Needed to ie. trigger rediscovery after Home Assistant has been restarted
client.subscribe(topic = HOMEASSISTANT_STATUS_TOPIC, qos = 1)

class EntityMqtt():
	def __init__(self, topic, payload, client, state = None):
		self.topic = topic
		self.payload = payload
		self.state = state
		self.client = client
		self.payload_on = payload["payload_on"]
		self.payload_off = payload["payload_off"]

	def activate(self):
		"""Activates the entity - meaning subscribing to it's command topic and publishing discovery"""
		self.subscribe_to_mqtt()
		self.publish_discovery_payload()

	def publish_discovery_payload(self):
		"""Publish the entity's discovery payload, which registers the entity with the MQTT broker (and also enables auto discovery with Home Assistant)"""
		logger.info("Publishing discovery payload for topic: %s", self.topic)

		discovery_message = json.dumps(self.payload)
		client.publish(f"{self.topic}/config", discovery_message, qos=1, retain=False)
		
		# From the docs:
		# After the configs have been published, the state topics will need an update, so they need to be republished.
		# https://www.home-assistant.io/integrations/mqtt/#use-the-birth-and-will-messages-to-trigger-discovery
		self.publish_state()
		
	def publish_state(self):
		"""Publish the entity's state to the MQTT broker"""
		logger.info("Publishing state for %s: %s", self.topic, self.state)

		# If these messages are published with a RETAIN flag, the MQTT switch will receive an instant state update after subscription, and will start with the correct state. Otherwise, the initial state of the switch will be unknown. A MQTT device can reset the current state to unknown using a None payload.
		client.publish(self.payload["state_topic"], self.state, retain=True)

	def subscribe_to_mqtt(self):
		"""Subscribes to messages from its own command topic"""
		client.subscribe(topic = self.payload["command_topic"], qos = 1)
		
	def on_message(self, client, userdata, payload):
		"""Handle messages"""
		logger.error(f"on_message not implemented for topic: %s", self.topic) 

class EntityDisplay(EntityMqtt):
	def __init__(self, topic, payload, client, state = None):
		super().__init__(topic, payload, client, state)
		self.activate()

	def on_message(self, client, userdata, payload):
		"""EntityDisplay's own handling of incoming messages"""
		# Callback function to handle incoming MQTT messages
		logger.info("New message for topic %s with payload %s. Client: %s, userdata: %s", self.topic, payload, client, userdata)

		# Check the payload and perform corresponding actions
		if payload == self.payload_on:
			self.turn_display_on()
		elif payload == self.payload_off:
			self.turn_display_off()

	# Function to turn on the display
	def turn_display_on(self):
		"""Turn the device's display on"""
		logger.debug("Turning display on")

		# Simulate a key press to wake up the display (Control key)
		win32api.keybd_event(0x11, 0, 0, 0)  # Control key down
		win32api.keybd_event(0x11, 0, 0x0002, 0)  # Control key up
		
		self.state = self.payload_on
		self.publish_state()

	# Function to turn off the display
	def turn_display_off(self):
		"""Turn the device's display off"""
		logger.debug("Turning display off")
	
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
		logger.debug("Is dark mode enabled?: %s", self.state)

		self.activate()

	def on_message(self, client, userdata, payload):
		"""EntityDarkMode's own handing of messages"""
		logger.info("New message for topic %s with payload %s. Client: %s, userdata: %s", self.topic, payload, client, userdata)
		
		# Check the payload and perform corresponding actions
		if payload == self.payload_on:
			self.enable_dark_mode()
		elif payload == self.payload_off:
			self.disable_dark_mode()

	def enable_dark_mode(self):
		try:
			subprocess.run(["reg", "add", "HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "/v", "AppsUseLightTheme", "/t", "REG_DWORD", "/d", "0", "/f"], check=True)
			logger.debug("Dark mode enabled")

			self.state = self.payload_on
			self.publish_state()
		except Exception as e:
			logger.error("Error enabling dark mode: %s", e)

	def disable_dark_mode(self):
		try:
			subprocess.run(["reg", "add", "HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "/v", "AppsUseLightTheme", "/t", "REG_DWORD", "/d", "1", "/f"], check=True)
			logger.debug("Dark mode disabled")

			self.state = self.payload_off
			self.publish_state()
		except Exception as error:
			logger.error("Error disabling dark mode: %s", error)
			
	def get_dark_mode_state(self):
		try:
			result = subprocess.run(["reg", "query", "HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "/v", "AppsUseLightTheme"], capture_output=True, text=True, check=True)
			return self.payload_on if ("0x1" not in result.stdout) else self.payload_off
		except Exception as error:
			logger.error("Error checking dark mode status: %s", error)
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

logger.debug("MQTT_ENTITIES: %s", MQTT_ENTITIES)
	
# Main loop to keep the program running and handle MQTT messages
try:
	client.loop_start()
	while True:
		time.sleep(1)  # Keep the program running
except KeyboardInterrupt:
	logger.warning("Exiting program")
	client.disconnect()
	client.loop_stop()
	