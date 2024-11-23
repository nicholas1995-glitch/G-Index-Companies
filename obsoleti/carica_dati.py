import requests
from bs4 import BeautifulSoup
import openpyxl

# Funzione per recuperare i dati da Yahoo Finance
def recupera_dati(ticker):
    url = f"https://finance.yahoo.com/quote/{ticker}/key-statistics"
    response = requests.get(url)
    
    soup = BeautifulSoup(response.text, 'html.parser')
    
    # Trova il P/E ratio
    pe_ratio = soup.find('td', {'data-test': 'PE_RATIO-value'}).text if soup.find('td', {'data-test': 'PE_RATIO-value'}) else "N/A"
    
    # Trova il P/B ratio
    pb_ratio = soup.find('td', {'data-test': 'PRICE_TO_BOOK-value'}).text if soup.find('td', {'data-test': 'PRICE_TO_BOOK-value'}) else "N/A"
    
    # Trova il PEG ratio 5Y
    peg_ratio = soup.find('td', {'data-test': 'PEG_RATIO_5Y-value'}).text if soup.find('td', {'data-test': 'PEG_RATIO_5Y-value'}) else "N/A"
    
    return pe_ratio, pb_ratio, peg_ratio

# Funzione per caricare i dati nel file Excel
def carica_dati_in_excel():
    # Carica il workbook esistente
    wb = openpyxl.load_workbook(r'C:\Users\nicho\Desktop\Aziende_sottovalutate.xlsm')
    sheet = wb.active

    row = 2  # Iniziamo dalla seconda riga, considerando che la prima è l'intestazione
    
    while sheet.cell(row=row, 
