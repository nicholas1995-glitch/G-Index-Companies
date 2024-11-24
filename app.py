from flask import Flask, render_template, jsonify, request
from flask_cors import CORS
import requests
from lxml import html
import random
import time
import os
from flask_socketio import SocketIO, emit


app = Flask(__name__)
CORS(app)  # Abilita CORS per tutte le route
socketio = SocketIO(app, cors_allowed_origins="*")


# Lista aziende con nome completo, ticker e ISIN
companies_info = {
   "DHL.DE": {"name": "Deutsche Post", "isin": "DE0005552004"},
    "CBK.DE": {"name": "Commerzbank", "isin": "DE000CBK1001"},
    "DB1.DE": {"name": "Deutsche Börse", "isin": "DE0005810055"},
    "BMW.DE": {"name": "BMW", "isin": "DE0005190003"},
    "DTE.DE": {"name": "Deutsche Telekom", "isin": "DE0005557508"},
    "MRK.DE": {"name": "Merck", "isin": "DE0006599905"},
    "IFX.DE": {"name": "Infineon", "isin": "DE0006231004"},
    "MUV2.DE": {"name": "Münchener Rück", "isin": "DE0008430026"},
    "SAP.DE": {"name": "SAP", "isin": "DE0007164600"},
    "RHM.DE": {"name": "Rheinmetall", "isin": "DE0007030009"},
    "BEI.DE": {"name": "Beiersdorf", "isin": "DE0005200000"},
    "FRE.DE": {"name": "Fresenius SE", "isin": "DE0005785604"},
    "HNR1.DE": {"name": "Hannover Rück", "isin": "DE0008402215"},
    "RWE.DE": {"name": "RWE", "isin": "DE0007037129"},
    "SY1.DE": {"name": "Symrise", "isin": "DE000SYM9999"},
    "HEI.DE": {"name": "Heidelberg Materials", "isin": "DE0006047004"},
    "QIA.DE": {"name": "Qiagen", "isin": "NL0012169213"},
    "EOAN.DE": {"name": "E.ON", "isin": "DE000ENAG999"},
    "SRT3.DE": {"name": "Sartorius", "isin": "DE0007165631"},
    "VOW3.DE": {"name": "Volkswagen", "isin": "DE0007664039"},
    "ZAL.DE": {"name": "Zalando", "isin": "DE000ZAL1111"},
    "SHL.DE": {"name": "Siemens Healthineers", "isin": "DE000SHL1006"},
    "P911.DE": {"name": "Dr. Ing. h.c. F. Porsche AG", "isin": "DE000PAG9113"},
    "AC.PA": {"name": "Accor", "isin": "FR0000120404"},
    "AI.PA": {"name": "Air Liquide", "isin": "FR0000120073"},
    "CAP.PA": {"name": "Capgemini", "isin": "FR0000125338"},
    "CS.PA": {"name": "Axa", "isin": "FR0000120628"},
    "BN.PA": {"name": "Danone", "isin": "FR0000120644"},
    "ENGI.PA": {"name": "Engie", "isin": "FR0010208488"},
    "BNP.PA": {"name": "BNP Paribas", "isin": "FR0000131104"},
    "EN.PA": {"name": "Bouygues", "isin": "FR0000120503"},
    "MC.PA": {"name": "Louis Vuitton (LVMH)", "isin": "FR0000121014"},
    "OR.PA": {"name": "L'Oréal", "isin": "FR0000120321"},
    "SU.PA": {"name": "Schneider Electric", "isin": "FR0000121972"},
    "SGO.PA": {"name": "Saint-Gobain", "isin": "FR0000125007"},
    "CA.PA": {"name": "Carrefour", "isin": "FR0000120172"},
    "EL.PA": {"name": "EssilorLuxottica", "isin": "FR0000121667"},
    "RI.PA": {"name": "Pernod Ricard", "isin": "FR0000120693"},
    "ORA.PA": {"name": "Orange", "isin": "FR0000133308"},
    "GLE.PA": {"name": "Société Générale", "isin": "FR0000130809"},
    "SAN.PA": {"name": "Sanofi", "isin": "FR0000120578"},
    "VIV.PA": {"name": "Vivendi", "isin": "FR0000127771"},
    "DG.PA": {"name": "Vinci", "isin": "FR0000125486"},
    "RNO.PA": {"name": "Renault", "isin": "FR0000131906"},
    "VIE.PA": {"name": "Veolia Environnement", "isin": "FR0000124141"},
    "KER.PA": {"name": "Kering", "isin": "FR0000121485"},
    "HO.PA": {"name": "Thales", "isin": "FR0000121329"},
    "DSY.PA": {"name": "Dassault Systèmes", "isin": "FR0000130650"},
    "ERF.PA": {"name": "Eurofins Scientific SE", "isin": "FR0000038259"},
    "PUB.PA": {"name": "Publicis Groupe", "isin": "FR0000130577"},
    "RMS.PA": {"name": "Hermès International", "isin": "FR0000052292"},
    "SAF.PA": {"name": "Safran", "isin": "FR0000073272"},
    "LR.PA": {"name": "Legrand", "isin": "FR0010307819"},
    "STLAM.MI": {"name": "Stellantis", "isin": "NL00150001Q9"},
    "G.MI": {"name": "Assicurazioni Generali", "isin": "IT0000062072"},
    "SRG.MI": {"name": "Snam Rete Gas", "isin": "IT0003153415"},
    "UCG.MI": {"name": "UniCredit", "isin": "IT0005239360"},
    "LDO.MI": {"name": "Leonardo", "isin": "IT0003856405"},
    "STMMI.MI": {"name": "STMicroelectronics", "isin": "NL0000226223"},
    "ENEL.MI": {"name": "Enel", "isin": "IT0003128367"},
    "ENI.MI": {"name": "Eni", "isin": "IT0003132476"},
    "TRN.MI": {"name": "Terna", "isin": "IT0003242622"},
    "SPM.MI": {"name": "Saipem", "isin": "IT0005252140"},
    "CPR.MI": {"name": "Campari", "isin": "NL0015435975"},
    "DIA.MI": {"name": "DiaSorin", "isin": "IT0003492391"},
    "REC.MI": {"name": "Recordati", "isin": "IT0003828271"},
    "MONC.MI": {"name": "Moncler", "isin": "IT0004965148"},
    "INW.MI": {"name": "Inwit", "isin": "IT0005090300"},
}

# Elenco statico di User-Agent
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.93 Safari/537.36",
    "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:88.0) Gecko/20100101 Firefox/88.0",
    "Mozilla/5.0 (iPhone; CPU iPhone OS 14_6 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0 Mobile/15E148 Safari/604.1",
]

session = requests.Session()
session.headers.update({"User-Agent": random.choice(USER_AGENTS)})

def scrape_stock_data(ticker):
    try:
        url_main = f"https://finance.yahoo.com/quote/{ticker}/"
        url_stats = f"https://finance.yahoo.com/quote/{ticker}/key-statistics/"
        
        response_main = session.get(url_main, timeout=20)
        response_main.raise_for_status()
        tree_main = html.fromstring(response_main.content)
        
        response_stats = session.get(url_stats, timeout=20)
        response_stats.raise_for_status()
        tree_stats = html.fromstring(response_stats.content)

        pe_ratio = tree_main.xpath('//*[@id="nimbus-app"]/section/section/section/article/div[2]/ul/li[11]/span[2]/fin-streamer/text()')
        pb_ratio = tree_stats.xpath('//*[@id="nimbus-app"]/section/section/section/article/section[2]/div/table/tbody/tr[7]/td[2]/text()')
        peg_ratio = tree_stats.xpath('//*[@id="nimbus-app"]/section/section/section/article/section[2]/div/table/tbody/tr[5]/td[2]/text()')

        return {
            "P/E Ratio": pe_ratio[0].strip() if pe_ratio else "--",
            "P/Book Ratio": pb_ratio[0].strip() if pb_ratio else "--",
            "PEG Ratio (5y)": peg_ratio[0].strip() if peg_ratio else "--",
            "Ticker": ticker,
            "Full Name": companies_info[ticker]["name"],
            "ISIN": companies_info[ticker]["isin"],
            "Calcolo": f"((P/E: {pe_ratio[0] if pe_ratio else '--'} / 15) * 0.4) + (P/B: {pb_ratio[0] if pb_ratio else '--'} * 0.4) + (PEG: {peg_ratio[0] if peg_ratio else '--'} * 0.2)",
        }
    except Exception as e:
        return {
            "P/E Ratio": "--",
            "P/Book Ratio": "--",
            "PEG Ratio (5y)": "--",
            "Ticker": ticker,
            "Full Name": companies_info[ticker]["name"],
            "ISIN": companies_info[ticker]["isin"],
            "Calcolo": "Errore nel calcolo",
        }

def calculate_g_index(company):
    try:
        pe = float(company["P/E Ratio"].replace(",", "")) if company["P/E Ratio"] != "--" else 0
        pb = float(company["P/Book Ratio"].replace(",", "")) if company["P/Book Ratio"] != "--" else 0
        peg = float(company["PEG Ratio (5y)"].replace(",", "")) if company["PEG Ratio (5y)"] != "--" else 0
        return round(((pe / 15) * 0.4) + (pb * 0.4) + (peg * 0.2), 2)
    except:
        return "--"

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/fetch_data", methods=["POST"])
def fetch_data():
    tickers = list(companies_info.keys())
    companies = []
    for i in range(0, len(tickers), 5):
        batch = tickers[i:i + 5]
        for ticker in batch:
            company_data = scrape_stock_data(ticker)
            company_data["Indice G"] = calculate_g_index(company_data)
            companies.append(company_data)
        if i + 5 < len(tickers):  # Attendi solo tra i batch
            print(f"Attesa di 1 secondi prima di elaborare il prossimo batch...")
            time.sleep(1)

    green = [c for c in companies if c["Indice G"] != "--" and c["Indice G"] < 1]
    orange = [c for c in companies if c["Indice G"] != "--" and 1 <= c["Indice G"] < 1.5]
    red = [c for c in companies if c["Indice G"] != "--" and c["Indice G"] >= 1.5]

    green.sort(key=lambda x: x["Indice G"])
    orange.sort(key=lambda x: x["Indice G"])
    red.sort(key=lambda x: x["Indice G"])

    return jsonify({"green": green, "orange": orange, "red": red})

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    socketio.run(app, host="0.0.0.0", port=port, debug=False)


@socketio.on("start_fetch")
def handle_start_fetch():
    tickers = list(companies_info.keys())
    for ticker in tickers:
        company_data = scrape_stock_data(ticker)
        socketio.emit("update_data", company_data)

