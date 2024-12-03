import requests
from lxml import html
import random
import os
import json
import logging
import time
from google.cloud import firestore
from google.oauth2 import service_account
from datetime import datetime
import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Configura il logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

# Percorso del Secret File montato da Render
FIRESTORE_KEY_PATH = '/etc/secrets/firestore-key.json'

# Inizializza Firestore con le credenziali
try:
    if not os.path.exists(FIRESTORE_KEY_PATH):
        logger.error(f"File delle credenziali di Firestore non trovato: {FIRESTORE_KEY_PATH}")
        exit(1)
    credentials = service_account.Credentials.from_service_account_file(FIRESTORE_KEY_PATH)
    db = firestore.Client(credentials=credentials)
    logger.info("Connessione a Firestore riuscita")
except Exception as e:
    logger.error(f"Errore nella connessione a Firestore: {e}")
    exit(1)

# Lista aziende con nome completo, ticker e ISIN
companies_info = {
    "DHL.DE": {"name": "Deutsche Post", "isin": "DE0005552004"},
    "CBK.DE": {"name": "Commerzbank", "isin": "DE000CBK1001"},
    "DB1.DE": {"name": "Deutsche Börse", "isin": "DE0005810055"},
    "BMW.DE": {"name": "BMW", "isin": "DE0005190003"},
    "DTE.DE": {"name": "Deutsche Telekom", "isin": "DE0005557508"},
    "MRK.DE": {"name": "Merck", "isin": "DE0006599905"},
    "IFX.DE": {"name": "Infineon", "isin": "DE0006231004"},
 
}

# Elenco User-Agent
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.93 Safari/537.36",
    "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:88.0) Gecko/20100101 Firefox/88.0",
    "Mozilla/5.0 (iPhone; CPU iPhone OS 14_6 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0 Mobile/15E148 Safari/604.1",
    # Gli altri User-Agent che avevi...
]

session = requests.Session()
session.headers.update({"User-Agent": random.choice(USER_AGENTS)})

def scrape_stock_data(ticker):
    try:
        logger.info(f"Inizio scraping per {ticker}")
        url_main = f"https://finance.yahoo.com/quote/{ticker}/"
        url_stats = f"https://finance.yahoo.com/quote/{ticker}/key-statistics/"
        url_history = f"https://finance.yahoo.com/quote/{ticker}/history/"

        session = requests.Session()
        session.headers.update({"User-Agent": random.choice(USER_AGENTS)})

        # Richiesta per la pagina principale
        logger.info(f"Richiesta GET a {url_main}")
        response_main = session.get(url_main, timeout=20)
        time.sleep(2)  # Attendi 2 secondi dopo la richiesta
        response_main.raise_for_status()
        tree_main = html.fromstring(response_main.content)
        logger.info(f"Risposta ricevuta da {url_main} (Status Code: {response_main.status_code})")

        # Richiesta per la pagina delle statistiche
        logger.info(f"Richiesta GET a {url_stats}")
        response_stats = session.get(url_stats, timeout=20)
        time.sleep(5)  # Attendi 5 secondi dopo la richiesta
        response_stats.raise_for_status()
        tree_stats = html.fromstring(response_stats.content)
        logger.info(f"Risposta ricevuta da {url_stats} (Status Code: {response_stats.status_code})")

        # Richiesta per la pagina storica per il prezzo
        logger.info(f"Richiesta GET a {url_history}")
        response_history = session.get(url_history, timeout=20)
        time.sleep(2)  # Attendi 2 secondi dopo la richiesta
        response_history.raise_for_status()
        tree_history = html.fromstring(response_history.content)
        logger.info(f"Risposta ricevuta da {url_history} (Status Code: {response_history.status_code})")

        # Usa i nuovi XPath
        pe_ratio = tree_main.xpath('//*[@id="nimbus-app"]/section/section/section/article/div[2]/ul/li[11]/span[2]/fin-streamer/text()')
        pb_ratio = tree_stats.xpath('//*[@id="nimbus-app"]/section/section/section/article/section[2]/div/table/tbody/tr[7]/td[2]/text()')
        peg_ratio = tree_stats.xpath('//*[@id="nimbus-app"]/section/section/section/article/section[2]/div/table/tbody/tr[5]/td[2]/text()')
        price = tree_history.xpath('//*[@id="nimbus-app"]/section/section/section/article/div[1]/div[3]/table/tbody/tr[1]/td[6]/text()')

        # Aggiungi log dei valori estratti
        logger.debug(f"P/E Ratio estratto: {pe_ratio}")
        logger.debug(f"P/Book Ratio estratto: {pb_ratio}")
        logger.debug(f"PEG Ratio estratto: {peg_ratio}")
        logger.debug(f"Price estratto: {price}")

        # Verifica se i valori sono stati estratti correttamente
        if not pe_ratio or not pb_ratio or not peg_ratio or not price:
            logger.warning(f"Dati incompleti per {ticker}: PE={pe_ratio}, PB={pb_ratio}, PEG={peg_ratio}, Price={price}")

        logger.info(f"Scraping completato per {ticker}")

        return {
            "P/E Ratio": pe_ratio[0].strip() if pe_ratio else "--",
            "P/Book Ratio": pb_ratio[0].strip() if pb_ratio else "--",
            "PEG Ratio (5y)": peg_ratio[0].strip() if peg_ratio else "--",
            "Price": price[0].strip() if price else "--",
            "Ticker": ticker,
            "Full Name": companies_info[ticker]["name"],
            "ISIN": companies_info[ticker]["isin"],
            "Calcolo": f"((P/E: {pe_ratio[0] if pe_ratio else '--'} / 15) * 0.4) + (P/B: {pb_ratio[0] if pb_ratio else '--'} * 0.4) + (PEG: {peg_ratio[0] if peg_ratio else '--'} * 0.2)",
        }
    except Exception as e:
        logger.error(f"Errore durante lo scraping di {ticker}: {e}")
        return {
            "P/E Ratio": "--",
            "P/Book Ratio": "--",
            "PEG Ratio (5y)": "--",
            "Price": "--",
            "Ticker": ticker,
            "Full Name": companies_info[ticker]["name"],
            "ISIN": companies_info[ticker]["isin"],
            "Calcolo": "Errore nel calcolo",
        }

def calculate_g_index(company):
    try:
        pe = float(company["P/E Ratio"].replace(",", ".")) if company["P/E Ratio"] not in ("--", "") else None
        pb = float(company["P/Book Ratio"].replace(",", ".")) if company["P/Book Ratio"] not in ("--", "") else None
        peg = float(company["PEG Ratio (5y)"].replace(",", ".")) if company["PEG Ratio (5y)"] not in ("--", "") else None

        # Verifica se tutti i valori sono disponibili
        if pe is None or pb is None or peg is None:
            logger.warning(f"Dati incompleti per il calcolo dell'Indice G di {company['Ticker']}")
            return "--"

        g_index = round(((pe / 15) * 0.4) + (pb * 0.4) + (peg * 0.2), 2)
        return g_index
    except Exception as e:
        logger.error(f"Errore nel calcolo dell'Indice G per {company['Ticker']}: {e}")
        return "--"

def generate_excel(data):
    """
    Genera un file Excel dove ogni foglio è nominato come il ticker e contiene i dati richiesti.
    """
    logger.info("Inizio generazione file Excel")
    file_name = "dati_aziende_ticker.xlsx"

    date = datetime.now().strftime("%Y-%m-%d")
    
    # Verifica se la lista 'data' è vuota
    if not data:
        logger.error("La lista 'companies' è vuota. Nessun dato da scrivere nel file Excel.")
        return file_name

    # Crea il file Excel se non esiste
    if not os.path.exists(file_name):
        wb = openpyxl.Workbook()
        # Non rimuovere il foglio attivo
        wb.save(file_name)

    # Carica il workbook esistente
    wb = openpyxl.load_workbook(file_name)

    for company in data:
        ticker = company["Ticker"]
        sheet_name = ticker

        if sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
        else:
            if 'Sheet' in wb.sheetnames and wb['Sheet'].max_row == 1 and wb['Sheet'].max_column == 1 and wb['Sheet'].cell(1,1).value is None:
                # Rinomina il foglio di default e usalo
                sheet = wb['Sheet']
                sheet.title = sheet_name
                # Scrivi l'intestazione
                sheet.append(["Data", "P/E Ratio", "P/Book Ratio", "PEG Ratio (5y)", "Price"])
            else:
                sheet = wb.create_sheet(sheet_name)
                # Scrivi l'intestazione
                sheet.append(["Data", "P/E Ratio", "P/Book Ratio", "PEG Ratio (5y)", "Price"])

        # Aggiungi i dati nella nuova riga
        sheet.append([
            date,
            company["P/E Ratio"],
            company["P/Book Ratio"],
            company["PEG Ratio (5y)"],
            company["Price"]
        ])

    # Rimuovi il foglio di default se è vuoto e ci sono altri fogli
    if 'Sheet' in wb.sheetnames and len(wb.sheetnames) > 1:
        sheet = wb['Sheet']
        if sheet.max_row == 1 and sheet.max_column == 1 and sheet.cell(1,1).value is None:
            wb.remove(sheet)

    # Assicurati che almeno un foglio sia visibile prima di salvare
    if not wb.sheetnames:
        # Creiamo un foglio temporaneo
        wb.create_sheet("Foglio_Temporaneo")
        logger.warning("Workbook vuoto. Aggiunto 'Foglio_Temporaneo' per evitare errori.")

    wb.save(file_name)
    logger.info("File Excel generato con successo")
    return file_name


def send_email(file_name):
    """
    Invia il file Excel via email.
    """
    logger.info("Inizio invio email")
    
    # Percorso del file segreto con la password
    EMAIL_PASSWORD_PATH = '/etc/secrets/EMAIL_PASSWORD'
    
    # Leggi la password dal file
    try:
        with open(EMAIL_PASSWORD_PATH, 'r') as f:
            password = f.read().strip()  # Usa .strip() per rimuovere eventuali spazi extra
    except Exception as e:
        logger.error(f"Errore nel caricamento della password email: {e}")
        exit(1)  # Termina l'esecuzione in caso di errore
    
    sender_email = "tua_email@gmail.com"
    receiver_email = "destinatario_email@gmail.com"
    
    subject = "Dati aggiornati aziende"
    body = "In allegato trovi i dati aggiornati delle aziende."

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Aggiungi l'allegato
    with open(file_name, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename={file_name}")
    msg.attach(part)

    # Invio email
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
            logger.info("Email inviata con successo")
    except Exception as e:
        logger.error(f"Errore nell'invio dell'email: {e}")

def get_previous_ticker_categories():
    """
    Recupera le liste precedenti dei ticker per tutte e tre le categorie da Firestore.
    """
    try:
        doc_ref = db.collection("ticker_categories")
        docs = doc_ref.stream()
        previous = {}
        for doc in docs:
            previous[doc.id] = doc.to_dict()
        if not previous:
            logger.info("Nessun dato precedente trovato per ticker_categories.")
        return previous
    except Exception as e:
        logger.error(f"Errore nel recuperare ticker_categories da Firestore: {e}")
        return {}

def update_ticker_categories(new_green, new_orange, new_red):
    """
    Aggiorna le liste dei ticker per tutte e tre le categorie in Firestore.
    """
    try:
        doc_ref_green = db.collection("ticker_categories").document("green")
        doc_ref_orange = db.collection("ticker_categories").document("orange")
        doc_ref_red = db.collection("ticker_categories").document("red")
        
        green_dict = {company["Ticker"]: company["Indice G"] for company in new_green}
        orange_dict = {company["Ticker"]: company["Indice G"] for company in new_orange}
        red_dict = {company["Ticker"]: company["Indice G"] for company in new_red}
        
        doc_ref_green.set(green_dict)
        doc_ref_orange.set(orange_dict)
        doc_ref_red.set(red_dict)
        
        logger.info("Liste ticker_categories aggiornate in Firestore.")
    except Exception as e:
        logger.error(f"Errore nell'aggiornare ticker_categories in Firestore: {e}")

def compare_ticker_categories(previous, current):
    """
    Confronta le liste precedente e corrente dei ticker per tutte e tre le categorie.
    Ritorna un dizionario con entrate, uscite e trasferimenti.
    """
    categories = ['green', 'orange', 'red']
    changes = {
        'entered': [],
        'exited': [],
        'transferred': []
    }
    
    # Creiamo un dizionario che mappa ticker a categorie precedenti
    ticker_to_previous_category = {}
    for category in categories:
        if category in previous:
            for ticker in previous[category]:
                ticker_to_previous_category[ticker] = category
    
    # Creiamo un dizionario che mappa ticker a categorie correnti
    ticker_to_current_category = {}
    for category in categories:
        if category in current:
            for ticker in current[category]:
                ticker_to_current_category[ticker] = category
    
    # Ticker attuali e precedenti
    current_tickers = set(ticker_to_current_category.keys())
    previous_tickers = set(ticker_to_previous_category.keys())
    
    # Identificare i ticker entrati (presenti ora ma non prima)
    entered_tickers = current_tickers - previous_tickers
    for ticker in entered_tickers:
        category = ticker_to_current_category[ticker]
        indice_g = current[category][ticker]
        changes['entered'].append({
            'Ticker': ticker,
            'Categoria': category,
            'Indice G': indice_g
        })
    
    # Identificare i ticker usciti (presenti prima ma non ora)
    exited_tickers = previous_tickers - current_tickers
    for ticker in exited_tickers:
        category = ticker_to_previous_category[ticker]
        indice_g = previous[category][ticker]
        changes['exited'].append({
            'Ticker': ticker,
            'Categoria': category,
            'Indice G': indice_g
        })
    
    # Identificare i ticker trasferiti (presente in entrambe ma categoria diversa)
    common_tickers = current_tickers & previous_tickers
    for ticker in common_tickers:
        previous_category = ticker_to_previous_category[ticker]
        current_category = ticker_to_current_category[ticker]
        if previous_category != current_category:
            changes['transferred'].append({
                'Ticker': ticker,
                'Da': previous_category,
                'A': current_category,
                'Indice G': current[current_category][ticker]
            })
    
    return changes

def send_change_notification(changes):
    """
    Invia un'email notificando le modifiche nelle liste delle categorie.
    """
    if not changes['entered'] and not changes['exited'] and not changes['transferred']:
        logger.info("Nessuna modifica nelle liste delle categorie.")
        return
    
    logger.info("Invio notifica via email per modifiche nelle liste delle categorie.")
    
    # Percorso del file segreto con la password
    EMAIL_PASSWORD_PATH = '/etc/secrets/EMAIL_PASSWORD'
    
    # Leggi la password dal file
    try:
        with open(EMAIL_PASSWORD_PATH, 'r') as f:
            password = f.read().strip()
    except Exception as e:
        logger.error(f"Errore nel caricamento della password email: {e}")
        return
    
    sender_email = "nicholas.gazzola@gmail.com"
    receiver_email = "nicholas.gazzola@gmail.com"
    
    subject = "Aggiornamenti nelle Liste delle Aziende"
    
    body = "Ci sono stati aggiornamenti nelle liste delle aziende:\n\n"
    
    if changes['entered']:
        body += "**Aziende Entrate nelle Liste:**\n"
        for company in changes['entered']:
            body += f"- {company['Ticker']} nella categoria **{company['Categoria'].capitalize()}** con Indice G: {company['Indice G']}\n"
        body += "\n"
    
    if changes['exited']:
        body += "**Aziende Uscite dalle Liste:**\n"
        for company in changes['exited']:
            body += f"- {company['Ticker']} dalla categoria **{company['Categoria'].capitalize()}** con Indice G: {company['Indice G']}\n"
        body += "\n"
    
    if changes['transferred']:
        body += "**Aziende Trasferite tra le Liste:**\n"
        for company in changes['transferred']:
            body += f"- {company['Ticker']} da **{company['Da'].capitalize()}** a **{company['A'].capitalize()}** con Indice G: {company['Indice G']}\n"
        body += "\n"
    
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    
    # Invio email
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
            logger.info("Email di notifica inviata con successo.")
    except Exception as e:
        logger.error(f"Errore nell'invio dell'email di notifica: {e}")

def update_data():
    """
    Esegue lo scraping dei dati, calcola l'indice G, salva i dati su Firestore,
    genera un Excel e invia il file via email.
    """
    logger.info("Inizio aggiornamento dei dati")
    tickers = list(companies_info.keys())
    companies = []

    for i in range(0, len(tickers), 5):  # Passo di 5
        batch = tickers[i:i + 5]
        logger.info(f"Elaborazione batch {i//5 + 1}: {batch}")
        for ticker in batch:
            logger.info(f"Elaborazione ticker: {ticker}")
            company_data = scrape_stock_data(ticker)
            company_data["Indice G"] = calculate_g_index(company_data)
            companies.append(company_data)
            logger.info(f"Ticker {ticker} elaborato con Indice G: {company_data['Indice G']}")
        if i + 5 < len(tickers):  # Attendi solo se ci sono altri batch
            logger.info(f"Attesa di 60 secondi prima di elaborare il prossimo batch...")
            time.sleep(60)

    green = [c for c in companies if c["Indice G"] != "--" and c["Indice G"] < 1]
    orange = [c for c in companies if c["Indice G"] != "--" and 1 <= c["Indice G"] < 1.5]
    red = [c for c in companies if c["Indice G"] != "--" and c["Indice G"] >= 1.5]

    green.sort(key=lambda x: x["Indice G"])
    orange.sort(key=lambda x: x["Indice G"])
    red.sort(key=lambda x: x["Indice G"])

    stored_data = {"green": green, "orange": orange, "red": red}

    # Nuove aggiunte: Gestione delle categorie e notifiche
    # Recupera le liste precedenti delle categorie
    previous_categories = get_previous_ticker_categories()

    # Costruisci il dizionario delle categorie correnti
    current_categories = {
        "green": {company["Ticker"]: company["Indice G"] for company in green},
        "orange": {company["Ticker"]: company["Indice G"] for company in orange},
        "red": {company["Ticker"]: company["Indice G"] for company in red}
    }

    # Confronta le categorie
    changes = compare_ticker_categories(previous_categories, current_categories)

    # Aggiorna le categorie in Firestore
    update_ticker_categories(green, orange, red)

    # Invia notifiche via email se ci sono cambiamenti
    send_change_notification(changes)

    # Salva i dati su Firestore
    try:
        db.collection("stored_data").document("data").set(stored_data)
        logger.info("Dati salvati su Firestore con successo")
    except Exception as e:
        logger.error(f"Errore durante il salvataggio dei dati su Firestore: {e}")

    # Genera l'Excel e invia via email
    file_name = generate_excel(companies)
    send_email(file_name)

    logger.info("Aggiornamento dei dati completato")

if __name__ == "__main__":
    update_data()
