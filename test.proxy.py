import requests

PROXY_LIST = [
    "http://51.254.69.243:3128",
    "http://81.171.24.199:3128",
    "http://176.31.200.104:3128",
    "http://163.172.182.164:3128",
]

def test_proxy(proxy):
    try:
        response = requests.get("https://www.google.com", proxies={"http": proxy, "https": proxy}, timeout=5)
        if response.status_code == 200:
            print(f"Proxy funzionante: {proxy}")
            return True
    except Exception as e:
        print(f"Errore con il proxy {proxy}: {e}")
    return False

working_proxies = [proxy for proxy in PROXY_LIST if test_proxy(proxy)]

print("Proxy validi:")
print(working_proxies)
