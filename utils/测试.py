import json
import os

filename = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'static', 'res', '2024年3月雄安信息价字典.json')

with open(filename, 'r', encoding='utf-8') as f:
    data = json.load(f)

result_key = []
result_value = []
for key, value in data.items():
    result_key.append(key)
    result_value.append(value)

print(result_value)
# for value in result_value:
#     for v in value:
#         print(v)