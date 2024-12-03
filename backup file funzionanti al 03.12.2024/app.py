from flask import Flask, render_template, jsonify
from flask_cors import CORS
from flask_socketio import SocketIO, emit
import os
import json
import logging
from google.cloud import firestore
from google.oauth2 import service_account
from flask import send_file 

app = Flask(__name__)
CORS(app)  # Abilita CORS per tutte le route
socketio = SocketIO(app, cors_allowed_origins="*", ping_timeout=120, ping_interval=25)

# Configura il logging
logging.basicConfig(
    level=logging.INFO,  # Puoi cambiare a DEBUG per più dettagli
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler()  # Stampa i log nel terminale
    ]
)
logger = logging.getLogger(__name__)

# Percorso del Secret File montato da Render
FIRESTORE_KEY_PATH = '/etc/secrets/firestore-key.json'

# Inizializza Firestore con le credenziali del file
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

def load_data():
    """
    Carica i dati da Firestore.
    Se il documento non esiste, inizializza con dati vuoti.
    """
    try:
        doc_ref = db.collection("stored_data").document("data")
        doc = doc_ref.get()
        if doc.exists:
            stored_data = doc.to_dict()
            logger.info("Dati caricati da Firestore")
            return stored_data
        else:
            logger.warning("Nessun documento trovato in Firestore. Inizializzazione con dati vuoti.")
            return {"green": [], "orange": [], "red": []}
    except Exception as e:
        logger.error(f"Errore nel caricamento dei dati da Firestore: {e}")
        return {"green": [], "orange": [], "red": []}

def load_last_update_time():
    """
    Recupera l'ora dell'ultimo aggiornamento.
    """
    try:
        doc_ref = db.collection("stored_data").document("metadata")
        doc = doc_ref.get()
        if doc.exists:
            metadata = doc.to_dict()
            return metadata.get("last_update", "N/A")
        else:
            return "Nessun aggiornamento trovato"
    except Exception as e:
        logger.error(f"Errore nel recupero dell'ultimo aggiornamento: {e}")
        return "Errore"


@app.route("/")
def home():
    """
    Route principale che serve la pagina index.html.
    """
    return render_template("index.html")

@app.route("/fetch_data", methods=["POST"])
def fetch_data():
    """
    Endpoint per recuperare i dati memorizzati.
    """
    logger.info("Richiesta di fetch_data ricevuta")
    stored_data = load_data()
    logger.info("Dati inviati al client con successo")
    return jsonify(stored_data)

@app.route('/download_excel/<filename>', methods=['GET'])
def download_excel(filename):
    """
    Consente di scaricare i file Excel generati.
    """
    file_path = os.path.join('generated_excel', filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return jsonify({"error": "File non trovato"}), 404


@app.route("/last_update", methods=["GET"])
def last_update():
    """
    Endpoint per verificare l'ultimo aggiornamento.
    """
    last_update_time = load_last_update_time()
    return jsonify({"last_update": last_update_time})



@socketio.on("start_fetch")
def handle_start_fetch():
    """
    Evento SocketIO per inviare i dati al client.
    """
    logger.info("Inizio handle_start_fetch")
    stored_data = load_data()
    socketio.emit("update_data", stored_data)
    logger.info("Handle_start_fetch completato")

if __name__ == "__main__":
    # Carica i dati all'avvio (opzionale, può essere gestito dinamicamente)
    # load_data()

    port = int(os.environ.get("PORT", 5000))
    logger.info(f"Avvio del server Flask su porta {port}")
    try:
        socketio.run(app, host="0.0.0.0", port=port, debug=False)
    except (KeyboardInterrupt, SystemExit):
        logger.info("Server Flask interrotto")