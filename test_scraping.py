import requests
from lxml import html

def test_scraping(ticker):
    url = f"https://finance.yahoo.com/quote/{ticker}/key-statistics/"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        tree = html.fromstring(response.content)
        
        # Estrarre i dati usando gli XPath forniti
        pe_ratio = tree.xpath('//*[@id="nimbus-app"]/section/section/section/article/section[2]/div/table/tbody/tr[3]/td[2]/text()')
        pb_ratio = tree.xpath('//*[@id="nimbus-app"]/section/section/section/article/section[2]/div/table/tbody/tr[7]/td[2]/text()')
        peg_ratio = tree.xpath('//*[@id="nimbus-app"]/section/section/section/article/section[2]/div/table/tbody/tr[5]/td[2]/text()')
        revenue_growth = tree.xpath('//*[@id="nimbus-app"]/section/section/section/article/article/div/section[1]/div/section[4]/table/tbody/tr[3]/td[2]/text()')
        
        return {
            "P/E Ratio": pe_ratio[0] if pe_ratio else "N/A",
            "P/Book Ratio": pb_ratio[0] if pb_ratio else "N/A",
            "PEG Ratio (5y)": peg_ratio[0] if peg_ratio else "N/A",
            "Quarterly Revenue Growth (yoy)": revenue_growth[0] if revenue_growth else "N/A",
        }
    except Exception as e:
        print(f"Errore durante lo scraping per {ticker}: {e}")
        return None

# Test con un singolo ticker
ticker = "UCG.MI"
data = test_scraping(ticker)
print(f"Dati per {ticker}: {data}")
