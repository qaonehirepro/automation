import json

a = {
    "Source": "n = input(\"Enter an integer\\n\")\r\nif (n%2 == 0):\r\n  print (\"Even\\n\")\r\nelse:\r\n  print (\"Odd\\n\")",
    "Lang": "python", "callback_url": "callback_url_string", "cid": "1787:128957:15237:1400389",
    "testcase_url": "testcase_url_string"}

# print(a['data'])
# new = json.loads(a['data'])
print(a.get('Source'))
