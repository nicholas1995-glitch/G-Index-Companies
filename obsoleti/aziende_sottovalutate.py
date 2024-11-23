import pandas as pd
import yfinance as yf

# Elenco dei ticker
tickers = [
    "UCG.MI", "ISP.MI", "STLAM.MI", "G.MI", "CNHI.MI", "ENEL.MI", "SRG.MI",
    "TRN.MI", "STM.MI", "INW.MI", "UNI.MI", "BPE.MI", "BAMI.MI", "IVG.MI",
    "BMPS.MI", "MB.MI", "ENI.MI", "PIRC.MI", "TEN.MI", "LDO.MI", "HER.MI",
    "PST.MI", "TIT.MI", "AZM.MI", "A2A.MI", "BMED.MI", "IG.MI", "MONC.MI",
    "PRY.MI", "ERG.MI", "IP.MI", "BGN.MI", "NRXI.MI", "FBK.MI", "DIA.MI",
    "CPR.MI", "REC.MI", "AMP.MI", "SPM.MI", "RACE.MI"
]

# Lista per memorizzare i dati
data = []

# Recupero dei dati per ogni ticker
for ticker in tickers:
    try:
        stock = yf.Ticker(ticker)
        info = stock.info
        
        # Recupero dati finanziari
        pe_ratio = info.get("trailingPE", None)
        pb_ratio = info.get("priceToBook", None)
        peg_ratio = info.get("pegRatio", None)
        
        # Calcolo dell'Indice G
        if pe_ratio is not None and pb_ratio is not None and peg_ratio is not None:
            indice_g = ((pe_ratio / 15) * 0.4) + (pb_ratio * 0.4) + (peg_ratio * 0.2)
        else:
            indice_g = None
        
        # Aggiunta dei dati alla lista
        data.append({
            "Ticker": ticker,
            "P/E Ratio": pe_ratio if pe_ratio is not None else "N/A",
            "P/B Ratio": pb_ratio if pb_ratio is not None else "N/A",
            "PEG Ratio 5Y": peg_ratio if peg_ratio is not None else "N/A",
            "Indice G": round(indice_g, 2) if indice_g is not None else "N/A"
        })
        print(f"Dati recuperati per {ticker}")
    except Exception as e:
        print(f"Errore per {ticker}: {e}")
        data.append({
            "Ticker": ticker,
            "P/E Ratio": "N/A",
            "P/B Ratio": "N/A",
            "PEG Ratio 5Y": "N/A",
            "Indice G": "N/A"
        })

# Creazione del DataFrame
df = pd.DataFrame(data)

# Ordinamento per Indice G
df = df[df["Indice G"] != "N/A"]  # Filtra le righe con Indice G valido
df.sort_values(by="Indice G", inplace=True)

# Evidenziamo le 4 aziende più sottovalutate
top_aziende = df.head(4)

# Salva i dati in Excel
df.to_excel("Aziende_sottovalutate.xlsx", index=False)
top_aziende.to_excel("Top_aziende_sottovalutate.xlsx", index=False)

# Stampa i risultati
print("Le 4 aziende più sottovalutate:")
print(top_aziende)
