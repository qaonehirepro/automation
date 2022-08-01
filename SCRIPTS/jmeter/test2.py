from selenium import webdriver
import json
# from selenium.webdriver.chrome.options import Options
#
# chrome_options = Options()
# chrome_options.add_argument("--use-fake-ui-for-media-stream")
# driver = webdriver.Chrome(executable_path="F:\\qa_automation\\chromedriver.exe")
# driver.get("https://amsin.hirepro.in/crpo/#/login/AT")
# a = driver.get_network_conditions()
# print(a)

chrome_options = webdriver.ChromeOptions()
chrome_options.set_capability(
                        "goog:loggingPrefs", {"performance": "ALL", "browser": "ALL"}
                    )
driver = webdriver.Chrome(executable_path="F:\\qa_automation\\chromedriver.exe", options=chrome_options)
# driver = webdriver.Chrome()
driver.get('https://amsin.hirepro.in/crpo/#/login/AT')
##visit your website, login, etc. then:
log_entries = driver.get_log("performance")
print(log_entries)
for entry in log_entries:
    a = entry.get('message')
    b = json.loads(a)
    c = b.get('message')
    print(c)
    print("________________________________________________\n\n\n")

# for entry in log_entries:

    # try:
    #     obj_serialized: str = entry.get("message")
    #     obj = json.loads(obj_serialized)
    #     message = obj.get("message")
    #     method = message.get("method")
    #     if method in ['Network.requestWillBeSentExtraInfo' or 'Network.requestWillBeSent']:
    #         try:
    #             for c in message['params']['associatedCookies']:
    #                 if c['cookie']['name'] == 'authToken':
    #                     bearer_token = c['cookie']['value']
    #         except:
    #             pass
    #     print(type(message), method)
    #     print('--------------------------------------')
    # except Exception as e:
    #     raise e from None