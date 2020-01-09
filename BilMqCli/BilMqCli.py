import configparser
import json
from pyrabbit2.api import Client

config = configparser.ConfigParser()
print("Loading properties....")
config.read('BilMqCli.ini')
sections = config.sections()
for index in range(0, len(sections)):
    print(f"{index}. {sections[index]}")
env_choice = int(input("Select your environment to work with : "))
print(f"Connecting to Rabbit MQ - {sections[env_choice]}...")
selected_section = config[sections[env_choice]]
client = Client(selected_section['host'], selected_section['username'], selected_section['password'], 40, 'https')
print("Connected.... ")
json_queue_response = client.get_queues(vhost=selected_section['virtual_host'])
for json_str in json_queue_response:
    print(json_str.get('name'))
json_exchange_response = client.get_exchanges(vhost=selected_section['virtual_host'])
for json_exchange_str in json_exchange_response:
    print(json_exchange_str.get('name'))
