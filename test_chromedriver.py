from selenium import webdriver
from selenium.webdriver.chrome.service import Service

# Configura il percorso di Chromedriver
service = Service("C:\\WebDriver\\chromedriver.exe")
options = webdriver.ChromeOptions()
options.add_argument("--headless")  # Modalità senza finestra (opzionale)

# Avvia il browser
driver = webdriver.Chrome(service=service, options=options)

try:
    # Apri una pagina di test
    driver.get("https://www.google.com")
    print("Pagina caricata con successo:", driver.title)
except Exception as e:
    print("Errore:", e)
finally:
    driver.quit()
