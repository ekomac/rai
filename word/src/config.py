import json
import os

os.chdir(os.path.dirname(os.path.realpath(__file__)))
path = os.getcwd()

CONFIG = None

with open('config.json', encoding='utf-8') as f:
    CONFIG = json.load(f)
