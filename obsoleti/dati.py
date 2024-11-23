from flask import Flask, render_template
import requests

app = Flask(__name__)

# Chiave API e lista delle aziende
API_KEY = "y69VjjbkaNqnqUgKBcESwJi8oDGLgi6z"
AZIENDE = ["ENI.MI"]  # Aggiungi altre aziende

def calcola_punteggio(data):
    pe = data.get("priceEarningsRatio", None)
    pb = data.get("priceToBookRatio", None)
    eps_growth = 5  # Stima della crescita EPS

    if pe is None or pb is None:
        return None
    peg = pe / eps_growth
    return (pe / 15 * 0.4) + (pb * 0.4) + (peg * 0.2)

def ottieni_dati():
    risultati = []
    for azienda in AZIENDE:
        url = f"https://financialmodelingprep.com/api/v3/profile/{azienda}?apikey={API_KEY}"
        response = requests.get(url)
        if response.status_code == 200:
            data = response.json()[0]
            punteggio = calcola_punteggio(data)
            if punteggio is not None:
                risultati.append({
                    "nome": data["companyName"],
                    "ticker": azienda,
                    "punteggio": punteggio,
                    "pe": data.get("priceEarningsRatio"),
                    "pb": data.get("priceToBookRatio"),
                })
    # Ordina i risultati per punteggio
    risultati_ordinati = sorted(risultati, key=lambda x: x["punteggio"])
    return risultati_ordinati[:4]

@app.route("/")
def index():
    aziende = ottieni_dati()
    return render_template("index.html", aziende=aziende)

if __name__ == "__main__":
    app.run(debug=True)
