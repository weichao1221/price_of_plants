import json
import os

filename = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'static', 'res', '2024年3月雄安信息价字典.json')

with open(filename, 'r', encoding='utf-8') as f:
    data = json.load(f)

for k, v in data.items():
    print(k)