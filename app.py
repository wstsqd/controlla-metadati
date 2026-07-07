# -*- coding: utf-8 -*-
import os
import re
import uuid
import urllib3
import requests
import openpyxl
import openpyxl.styles.stylesheet
from openpyxl.styles import Font
import io
from flask import Flask, request, jsonify, render_template, send_from_directory, send_file

# Salva il metodo originale ed applica la patch per evitare l'IndexError di openpyxl
original_expand = openpyxl.styles.stylesheet.Stylesheet._expand_named_style
def patched_expand(self, style):
    try:
        original_expand(self, style)
    except IndexError:
        pass
openpyxl.styles.stylesheet.Stylesheet._expand_named_style = patched_expand


# Disabilita gli avvisi per i certificati SSL non verificati
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.path.join(os.getcwd(), 'uploads')
app.config['PROCESSED_FOLDER'] = os.path.join(os.getcwd(), 'processed')

# Assicurati che le cartelle esistano
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['PROCESSED_FOLDER'], exist_ok=True)

# Endpoint del catalogo GeoNetwork
URL_CATALOGO = "https://rsdi.regione.basilicata.it/geonetwork/srv/api/search/records/_search"

# Parole comuni GIS da ignorare o penalizzare nella ricerca specifica
PAROLE_GENERICHE_GIS = {
    'area', 'aree', 'zona', 'zone', 'mappa', 'mappe', 'carta', 'carte', 'punti', 'punto', 'linea', 'linee', 
    'poligono', 'poligoni', 'servizio', 'servizi', 'dati', 'dato', 'db', 'dbgt', 'basilicata', 'regione', 
    'provincia', 'provincie', 'comune', 'comuni', 'limiti', 'limite', 'confini', 'confine', 'wms', 'wfs', 
    'raster', 'vettoriale', 'shapefile', 'vista', 'volumi', 'nan', 'vigente', 'vigente1', 'vigente2', 'vigente3'
}

# Parole vuote in italiano da escludere (comprensive di articoli e preposizioni semplici e articolate)
PAROLE_VUOTE = {
    'il', 'la', 'lo', 'i', 'gli', 'le', 'un', 'una', 'uno', 'di', 'a', 'da', 'in', 'con', 'su', 'per', 'tra', 'fra',
    'del', 'dello', 'della', 'dei', 'degli', 'delle', 'al', 'allo', 'alla', 'ai', 'agli', 'alle', 
    'dal', 'dallo', 'dalla', 'dai', 'dagli', 'dalle', 'nel', 'nello', 'nella', 'nei', 'negli', 'nelle', 
    'col', 'coi', 'sul', 'sullo', 'sulla', 'sui', 'sugli', 'sulle', 'e', 'o', 'ed', 'ad', 'od', 'si', 'no', 'non'
}

# Sinonimi per mappare le discrepanze tra Excel e Catalogo
MAPPA_SINONIMI = {
    'incendi': ['fuoco', 'incendio'],
    'incendio': ['fuoco', 'incendi'],
    'fuoco': ['incendi', 'incendio'],
    'cave': ['coltivazione', 'mineraria'],
    'idro': ['idraulica', 'idraulico', 'idrografico', 'idrogeologico', 'idrografia', 'acque', 'idrico'],
    'idrografia': ['idrografico', 'idro', 'acque'],
    'idraulica': ['idro', 'idraulico', 'acque'],
    'idraulico': ['idro', 'idraulica', 'acque'],
    'nocciolo': ['corilicola', 'noccioli'],
    'corilicola': ['nocciolo', 'noccioli'],
    'coste': ['costa', 'coste', 'costiero', 'litorale', 'spiaggia', 'portuale'],
    'costa': ['coste', 'costiero', 'litorale', 'spiaggia', 'portuale'],
    'apistica': ['apistico', 'api'],
    'apistico': ['apistica', 'api'],
    'pedologica': ['pedologia', 'suolo'],
    'pedologia': ['pedologica', 'suolo'],
    'granulometrica': ['granulometria', 'suolo'],
    'tessitura': ['tessiture', 'suolo'],
    'carbonati': ['carbonato', 'suolo'],
    'reazione': ['ph', 'suolo'],
    'alluvioni': ['alluvione', 'idraulica', 'idraulico'],
    'alluvione': ['alluvioni', 'idraulica', 'idraulico'],
    'frane': ['frana', 'iffi', 'geologico', 'geologica', 'movimenti franosi'],
    'frana': ['frane', 'iffi', 'geologico', 'geologica', 'movimento franoso']
}

def pulisci_testo(testo):
    """Pulisce il testo sostituendo le parentesi con spazi ed eliminando caratteri non alfanumerici."""
    if not testo:
        return ""
    testo = str(testo).lower().strip()
    testo = testo.replace('(', ' ').replace(')', ' ')
    testo = re.sub(r'[^a-z0-9]', ' ', testo)
    testo = re.sub(r'\s+', ' ', testo).strip()
    return testo

def ottieni_token(testo):
    """Estrae i token significativi (escludendo le parole vuote)."""
    testo_pulito = pulisci_testo(testo)
    parole = testo_pulito.split()
    return [p for p in parole if p not in PAROLE_VUOTE and len(p) >= 2]

def ottieni_radice(parola):
    """Ritorna la radice (stem) di 5 caratteri per gestire le flessioni in italiano."""
    return parola[:5] if len(parola) >= 5 else parola

def verifica_valore_numerico_o_corto(parola):
    """Riconosce se una parola rappresenta un valore numerico, anno o scala (es: 2007, 10k)."""
    if parola.isdigit():
        return True
    if re.match(r'^\d+(k|m|cm|tr|l\d+|k\d+|g\d+)?$', parola):
        return True
    return False

def contiene_parola_intera(parola, testo):
    """Verifica se 'parola' appare come parola intera nel 'testo' (non come sottostringa)."""
    if not parola or not testo:
        return False
    return bool(re.search(r'\b' + re.escape(parola) + r'\b', testo))

def sono_incompatibili(layer_tokens, db_tokens, title_tokens):
    """Verifica se i token del layer/db e il titolo appartengono a categorie escludentesi a vicenda."""
    query_set = set(layer_tokens + db_tokens)
    titolo_set = set(title_tokens)
    
    # Gruppi mutuamente esclusivi
    incendi = {'incendio', 'incendi', 'fuoco', 'crdi'}
    frane = {'frane', 'frana', 'iffi', 'geologico', 'geologica', 'liguridi', 'suolo', 'suoli'}
    idraulica_acque = {'idraulica', 'idraulico', 'idro', 'idrografico', 'idrografia', 'acque', 'invasi', 'laghi', 'alluvione', 'alluvioni', 'idrico', 'fiume', 'sinni'}
    ortofoto = {'ortofoto', 'ortofotocarta'}
    apistica = {'apistica', 'apistico', 'api'}
    pedologia = {'pedologia', 'pedologica', 'granulometrica', 'tessitura', 'carbonati', 'reazione', 'suolo', 'suoli'}
    archeologia = {'archeologiche', 'archeologia'}
    
    categorie = [incendi, frane, idraulica_acque, ortofoto, apistica, pedologia, archeologia]
    
    cat_query = set()
    for i, cat in enumerate(categorie):
        if query_set & cat:
            cat_query.add(i)
            
    cat_titolo = set()
    for i, cat in enumerate(categorie):
        if titolo_set & cat:
            cat_titolo.add(i)
            
    if cat_query and cat_titolo and cat_query != cat_titolo:
        # Escludi valori numerici (anni, scale) dall'intersezione: "2007" in comune
        # non deve salvare un match tra categorie incompatibili
        intersezione_specifica = set()
        for t in (query_set & titolo_set):
            if t in PAROLE_GENERICHE_GIS or t in {'rischio', 'servizio'}:
                continue
            if verifica_valore_numerico_o_corto(t):
                continue
            intersezione_specifica.add(t)
        if not intersezione_specifica:
            return True
            
    return False

def valuta_corrispondenza(nome_layer, nome_db, titolo_record, uuid_record, identificatore_record, testo_completo_record):
    """Assegna un punteggio di corrispondenza tra una query dell'Excel e una scheda metadati."""
    if not titolo_record:
        return 0
        
    punteggio = 0
    
    layer_pulito = pulisci_testo(nome_layer)
    db_pulito = pulisci_testo(nome_db)
    titolo_pulito = pulisci_testo(titolo_record)
    
    token_layer = ottieni_token(nome_layer)
    token_db = ottieni_token(nome_db)
    token_titolo = ottieni_token(titolo_record)
    
    # Incompatibility check
    if sono_incompatibili(token_layer, token_db, token_titolo):
        return 0
        
    id_minuscolo = str(identificatore_record).lower()
    uuid_minuscolo = str(uuid_record).lower()
    testo_completo_minuscolo = str(testo_completo_record).lower()
    
    # Tokenizza anche il testo completo per matching a parola intera
    token_testo_completo = set(pulisci_testo(testo_completo_record).split())
    token_id = set(pulisci_testo(identificatore_record).split())
    token_uuid = set(pulisci_testo(uuid_record).split())
    
    # 1. Corrispondenza frasale esatta
    if layer_pulito and titolo_pulito == layer_pulito:
        punteggio += 300
    elif layer_pulito and contiene_parola_intera(layer_pulito, titolo_pulito):
        non_generici_in_layer = [t for t in token_layer if t not in PAROLE_GENERICHE_GIS and not verifica_valore_numerico_o_corto(t)]
        if non_generici_in_layer:
            punteggio += 150
        else:
            punteggio += 40
            
    if db_pulito and db_pulito not in PAROLE_GENERICHE_GIS and not verifica_valore_numerico_o_corto(db_pulito) and titolo_pulito == db_pulito:
        punteggio += 250
    elif db_pulito and db_pulito not in PAROLE_GENERICHE_GIS and not verifica_valore_numerico_o_corto(db_pulito) and contiene_parola_intera(db_pulito, titolo_pulito):
        punteggio += 120

    if db_pulito and db_pulito not in PAROLE_GENERICHE_GIS and not verifica_valore_numerico_o_corto(db_pulito):
        if db_pulito == id_minuscolo or contiene_parola_intera(db_pulito, id_minuscolo):
            punteggio += 200
        if contiene_parola_intera(db_pulito, uuid_minuscolo):
            punteggio += 180

    # 2. Corrispondenza basata su Token (solo parole esatte e sinonimi)
    corrispondenze_specifiche_non_numeriche = 0
    
    def controlla_corrispondenza_token(t):
        """Cerca un token nei token del titolo (match esatto di parola, non sottostringa)."""
        if t in token_titolo:
            return True, True
        if t in MAPPA_SINONIMI:
            for sinonimo in MAPPA_SINONIMI[t]:
                if sinonimo in token_titolo:
                    return True, False
        return False, False

    # Token del Layer
    specifiche_layer_trovate = 0
    generiche_layer_trovate = 0
    for t in token_layer:
        generico = t in PAROLE_GENERICHE_GIS
        numerico = verifica_valore_numerico_o_corto(t)
        
        trovato, esatto = controlla_corrispondenza_token(t)
        
        if trovato:
            valore = 40 if esatto else 25
            if generico:
                punteggio += 5
                generiche_layer_trovate += 1
            else:
                punteggio += valore
                specifiche_layer_trovate += 1
                if not numerico:
                    corrispondenze_specifiche_non_numeriche += 1
        elif t in token_id or t in token_uuid:
            # Match esatto a parola intera nell'identificatore o UUID
            if generico:
                punteggio += 2
            else:
                punteggio += 20
                specifiche_layer_trovate += 1
                if not numerico:
                    corrispondenze_specifiche_non_numeriche += 1
        elif t in token_testo_completo:
            # Match esatto a parola intera nel testo completo (punteggio ridotto)
            if not generico:
                punteggio += 3
                # NON contare come corrispondenza specifica: il testo completo è troppo ampio
                
    # Token del Database
    specifiche_db_trovate = 0
    for t in token_db:
        if t in PAROLE_GENERICHE_GIS:
            continue
        numerico = verifica_valore_numerico_o_corto(t)
        
        trovato, esatto = controlla_corrispondenza_token(t)
        if trovato:
            valore = 30 if esatto else 20
            punteggio += valore
            specifiche_db_trovate += 1
            if not numerico:
                corrispondenze_specifiche_non_numeriche += 1
        elif t in token_id or t in token_uuid:
            punteggio += 50
            specifiche_db_trovate += 1
            if not numerico:
                corrispondenze_specifiche_non_numeriche += 1
        elif t in token_testo_completo:
            # Match nel testo completo: punteggio minimo e NON conta come match specifico
            punteggio += 3

    # 3. Bonus per il completamento dei token specifici richiesti
    specifiche_layer_richieste = [t for t in token_layer if t not in PAROLE_GENERICHE_GIS]
    if specifiche_layer_richieste and specifiche_layer_trovate == len(specifiche_layer_richieste):
        punteggio += 50
        
    specifiche_db_richieste = [t for t in token_db if t not in PAROLE_GENERICHE_GIS]
    if specifiche_db_richieste and specifiche_db_trovate == len(specifiche_db_richieste):
        punteggio += 50

    # 4. Tie-breaker bonus: se il titolo inizia esattamente con il nome cercato
    if layer_pulito and titolo_pulito.startswith(layer_pulito):
        punteggio += 10
    elif db_pulito and titolo_pulito.startswith(db_pulito):
        punteggio += 10

    # REGOLA CRITICA 1: Se la richiesta contiene parole specifiche non numeriche,
    # almeno una deve coincidere nel TITOLO (non nel testo completo)
    specifiche_richieste_totali = [t for t in (token_layer + token_db) if t not in PAROLE_GENERICHE_GIS and not verifica_valore_numerico_o_corto(t)]
    if specifiche_richieste_totali and corrispondenze_specifiche_non_numeriche == 0:
        return 0
    
    # REGOLA CRITICA 2: Richiedi una percentuale minima di corrispondenze specifiche
    # Per evitare match con 1 sola parola su molte
    totale_specifiche = len(specifiche_richieste_totali)
    if totale_specifiche >= 2 and corrispondenze_specifiche_non_numeriche < max(1, totale_specifiche // 2):
        # Se meno della metà delle parole specifiche corrisponde, penalizza pesantemente
        punteggio = punteggio // 3
    
    # REGOLA CRITICA 3: Se TUTTI i token specifici sono solo numerici (es. anni: 2007, 2008),
    # la query non ha contesto semantico sufficiente. Richiedi che almeno i token
    # generici/di contesto (es. "ortofoto") corrispondano nel titolo.
    token_tutti = [t for t in (token_layer + token_db) if t not in PAROLE_GENERICHE_GIS]
    token_non_numerici = [t for t in token_tutti if not verifica_valore_numerico_o_corto(t)]
    if token_tutti and not token_non_numerici:
        # Tutti i token specifici sono numerici - verifica il contesto generico
        token_contesto = [t for t in (token_layer + token_db) if t in PAROLE_GENERICHE_GIS]
        if token_contesto:
            contesto_trovato = any(t in token_titolo for t in token_contesto)
            if not contesto_trovato:
                # Il contesto generico (es. "ortofoto") non matcha il titolo → rifiuta
                return 0
        # Se non c'è nemmeno contesto generico (solo un anno), ridurre pesantemente
        # per evitare che "2007" da solo matchi qualsiasi record del 2007
        if not token_contesto:
            punteggio = punteggio // 4
    
    # REGOLA CRITICA 4: Se il nome_db contiene suffissi numerici (es. "apistica_1", "apistica_4"),
    # quei numeri devono essere presenti anche nel titolo del record.
    # Altrimenti "apistica_1" e "apistica_4" matcherebbero entrambi "Carta Apistica".
    if db_pulito:
        numeri_nel_db = [t for t in db_pulito.split() if t.isdigit()]
        if numeri_nel_db:
            # Controlla se almeno uno dei numeri del db è nel titolo
            parole_titolo = set(titolo_pulito.split())
            numeri_trovati = sum(1 for n in numeri_nel_db if n in parole_titolo)
            if numeri_trovati == 0:
                # Il suffisso numerico specifico non è nel titolo → match molto debole
                punteggio = punteggio // 4
        
    return punteggio

def scarica_catalogo_rsdi():
    """Scarica l'elenco completo dei metadati dal catalogo RSDI Basilicata."""
    payload = {
        "size": 600,
        "query": {
            "match_all": {}
        }
    }
    intestazioni = {
        "Accept": "application/json",
        "Content-Type": "application/json"
    }
    try:
        risposta = requests.post(URL_CATALOGO, json=payload, headers=intestazioni, verify=False, timeout=20)
        if risposta.status_code == 200:
            dati = risposta.json()
            hits = dati.get('hits', {}).get('hits', [])
            # Esclude le schede template
            risultati = [h for h in hits if h.get('_source', {}).get('isTemplate') != 'y']
            return risultati
        else:
            print(f"Errore caricamento catalogo (Status {risposta.status_code})")
            return []
    except Exception as e:
        print(f"Eccezione durante il download del catalogo: {e}")
        return []

def formatta_data(valore_data):
    """Formatta la data in formato italiano DD-MM-YYYY."""
    if not valore_data:
        return ""
    val_str = str(valore_data).strip()
    match = re.match(r'^(\d{4})-(\d{2})-(\d{2})', val_str)
    if match:
        anno, mese, giorno = match.groups()
        return f"{giorno}-{mese}-{anno}"
    return val_str

def estrai_date_metadato(source_record):
    """Estrae le date di pubblicazione e revisione dal record GeoNetwork."""
    data_pub = ""
    data_rev = ""
    
    lista_pub = source_record.get('publicationDateForResource', [])
    lista_rev = source_record.get('revisionDateForResource', [])
    
    if lista_pub and isinstance(lista_pub, list):
        data_pub = lista_pub[0]
    if lista_rev and isinstance(lista_rev, list):
        data_rev = lista_rev[0]
        
    if not data_pub or not data_rev:
        date_risorsa = source_record.get('resourceDate', [])
        if date_risorsa and isinstance(date_risorsa, list):
            for rd in date_risorsa:
                if isinstance(rd, dict):
                    tipo = rd.get('type')
                    valore = rd.get('date')
                    if tipo == 'publication' and not data_pub:
                        data_pub = valore
                    elif tipo == 'revision' and not data_rev:
                        data_rev = valore
                        
    return formatta_data(data_pub), formatta_data(data_rev)

def determina_visualizzatore(uid, cache=None):
    """Determina se un progetto è NUOVO o STANDARD in base al suo UID.
    
    Controlla l'endpoint crea-progetti/initFile per verificare se il progetto
    usa il nuovo visualizzatore. Restituisce ('NUOVO', link) o ('STANDARD', link).
    """
    if cache is not None and uid in cache:
        return cache[uid]
    
    url_nuovo_json = f"https://rsdi.regione.basilicata.it/crea-progetti/initFile/{uid}.json"
    url_nuovo_viewer = f"https://rsdi-view.regione.basilicata.it/#https://rsdi.regione.basilicata.it/crea-progetti/initFile/{uid}.json"
    url_standard = f"https://rsdi.regione.basilicata.it/viewGis/?project={uid}"
    
    try:
        risposta = requests.get(url_nuovo_json, verify=False, timeout=10)
        if risposta.status_code == 200:
            contenuto = risposta.text.strip()
            if contenuto and contenuto != '{}':
                risultato = ('NUOVO', url_nuovo_viewer)
                if cache is not None:
                    cache[uid] = risultato
                return risultato
    except Exception as e:
        print(f"Errore nel controllo visualizzatore per UID {uid}: {e}")
    
    risultato = ('STANDARD', url_standard)
    if cache is not None:
        cache[uid] = risultato
    return risultato


@app.route('/')
def home():
    """Carica la pagina principale."""
    return render_template('index.html')


@app.route('/carica', methods=['POST'])
def carica_e_elabora():
    """Riceve il file Excel, lo elabora e restituisce i dettagli in formato JSON."""
    if 'file' not in request.files:
        return jsonify({"errore": "Nessun file inviato"}), 400
        
    file_excel = request.files['file']
    if file_excel.filename == '':
        return jsonify({"errore": "Nome del file vuoto"}), 400
        
    if not file_excel.filename.endswith('.xlsx'):
        return jsonify({"errore": "Il file deve essere in formato Excel (.xlsx)"}), 400

    # Genera identificativi univoci
    id_sessione = str(uuid.uuid4())
    nome_salvato = f"{id_sessione}_{file_excel.filename}"
    percorso_upload = os.path.join(app.config['UPLOAD_FOLDER'], nome_salvato)
    percorso_elaborato = os.path.join(app.config['PROCESSED_FOLDER'], nome_salvato)
    
    file_excel.save(percorso_upload)
    
    # Scarica il catalogo metadati
    catalogo = scarica_catalogo_rsdi()
    if not catalogo:
        return jsonify({"errore": "Impossibile connettersi al catalogo RSDI Basilicata. Riprova più tardi."}), 503
        
    try:
        # Carica il workbook mantenendo lo stile e la formattazione originale
        wb = openpyxl.load_workbook(percorso_upload)
        if not wb.sheetnames:
            return jsonify({"errore": "File Excel vuoto o non valido"}), 400
            
        sheet = wb.active
        
        # Identifica gli indici delle colonne basandosi sull'intestazione (riga 1)
        riga_intestazione = [str(cella.value).strip().lower() if cella.value is not None else "" for cella in sheet[1]]
        
        # Mappa dei nomi delle colonne alle posizioni (1-indexed per openpyxl)
        # Gestione speciale: 'nome' appare due volte (col 2 = nome progetto, col 7 = nome layer)
        mappa_colonne = {}
        colonne_richieste = ['nome_completo', 'catalogo', 'data_pubblicazione', 'data_revisione', 'visualizzatore', 'uid', 'ufficio']
        for col_name in colonne_richieste:
            if col_name in riga_intestazione:
                mappa_colonne[col_name] = riga_intestazione.index(col_name) + 1
        
        # Per la colonna 'nome' (layer), prendi l'ULTIMA occorrenza (col 7, non col 2)
        indici_nome = [i + 1 for i, val in enumerate(riga_intestazione) if val == 'nome']
        if len(indici_nome) >= 2:
            mappa_colonne['nome_layer'] = indici_nome[-1]  # Ultima occorrenza = nome del layer
        elif len(indici_nome) == 1:
            mappa_colonne['nome_layer'] = indici_nome[0]
                
        # Verifica se mancano le colonne critiche di input
        if 'nome_completo' not in mappa_colonne and 'nome_layer' not in mappa_colonne:
            return jsonify({"errore": "Il file Excel deve contenere almeno le colonne 'nome_completo' o 'nome'."}), 400
            
        # Se mancano le colonne di output, le creiamo in fondo alla tabella
        nuovo_indice_colonna = len(riga_intestazione) + 1
        for col_name in ['catalogo', 'data_pubblicazione', 'data_revisione', 'visualizzatore', 'ufficio']:
            if col_name not in mappa_colonne:
                sheet.cell(row=1, column=nuovo_indice_colonna, value=col_name)
                mappa_colonne[col_name] = nuovo_indice_colonna
                nuovo_indice_colonna += 1
                
        # Elaborazione delle righe (dalla riga 2 in poi)
        cronologia_elaborazione = []
        
        idx_nome_completo = mappa_colonne.get('nome_completo')
        idx_nome_layer = mappa_colonne.get('nome_layer')
        idx_catalogo = mappa_colonne.get('catalogo')
        idx_pub = mappa_colonne.get('data_pubblicazione')
        idx_rev = mappa_colonne.get('data_revisione')
        idx_visualizzatore = mappa_colonne.get('visualizzatore')
        idx_uid = mappa_colonne.get('uid')
        idx_ufficio = mappa_colonne.get('ufficio')
        
        soglia_matching = 80
        cache_visualizzatore = {}  # Cache per evitare richieste duplicate per lo stesso UID
        
        for riga in range(2, sheet.max_row + 1):
            valore_nome_completo = sheet.cell(row=riga, column=idx_nome_completo).value if idx_nome_completo else ""
            valore_nome_layer = sheet.cell(row=riga, column=idx_nome_layer).value if idx_nome_layer else ""
            valore_uid = sheet.cell(row=riga, column=idx_uid).value if idx_uid else ""
            
            # Se la riga è completamente vuota, la saltiamo
            if not valore_nome_completo and not valore_nome_layer:
                continue
                
            valore_nome_completo_str = str(valore_nome_completo).strip() if valore_nome_completo is not None else ""
            valore_nome_layer_str = str(valore_nome_layer).strip() if valore_nome_layer is not None else ""
            valore_uid_str = str(valore_uid).strip() if valore_uid is not None else ""
            
            # Gestione namespace: se il nome contiene ':', prepara anche la versione senza namespace
            valore_nome_layer_senza_ns = ""
            if ':' in valore_nome_layer_str:
                valore_nome_layer_senza_ns = valore_nome_layer_str.split(':', 1)[1]
            
            miglior_record = None
            punteggio_massimo = 0
            
            for scheda in catalogo:
                source = scheda.get('_source', {})
                rto = source.get('resourceTitleObject')
                if isinstance(rto, dict):
                    titolo = rto.get('langita', rto.get('default', ''))
                else:
                    titolo = ""
                uuid_sch = source.get('uuid', '')
                identificatore = source.get('resourceIdentifier', '')
                
                if isinstance(identificatore, list):
                    parti_id = []
                    for item in identificatore:
                        if isinstance(item, dict):
                            parti_id.extend([str(v) for v in item.values() if v])
                        else:
                            parti_id.append(str(item))
                    identificatore = " ".join(parti_id)
                else:
                    identificatore = str(identificatore)
                    
                testo_completo = " ".join(source.get('anyText', [])) if isinstance(source.get('anyText'), list) else str(source.get('anyText', ''))
                
                # Ricerca 1: nome_completo come 'layer', nome_layer come 'nome_db'
                score = valuta_corrispondenza(valore_nome_completo_str, valore_nome_layer_str, titolo, uuid_sch, identificatore, testo_completo)
                
                # Ricerca 2: se il nome_layer contiene ':', prova con la parte dopo i due punti
                if valore_nome_layer_senza_ns:
                    score_senza_ns = valuta_corrispondenza(valore_nome_completo_str, valore_nome_layer_senza_ns, titolo, uuid_sch, identificatore, testo_completo)
                    if score_senza_ns > score:
                        score = score_senza_ns
                            
                if score > punteggio_massimo:
                    punteggio_massimo = score
                    miglior_record = scheda
            
            # --- Determinazione del Visualizzatore ---
            tipo_visualizzatore = ""
            link_visualizzatore = ""
            if valore_uid_str:
                tipo_visualizzatore, link_visualizzatore = determina_visualizzatore(valore_uid_str, cache_visualizzatore)
                    
            esito = "NO"
            titolo_trovato = ""
            link_trovato = "no"
            pub_data = ""
            rev_data = ""
            valore_ufficio = ""
            
            if punteggio_massimo >= soglia_matching:
                esito = "YES"
                source = miglior_record.get('_source', {})
                rto = source.get('resourceTitleObject')
                titolo_trovato = rto.get('langita', rto.get('default', '')) if isinstance(rto, dict) else ""
                uuid_trovato = source.get('uuid', '')
                link_trovato = f"https://rsdi.regione.basilicata.it/geonetwork/srv/ita/catalog.search#/metadata/{uuid_trovato}"
                pub_data, rev_data = estrai_date_metadato(source)
                
                # Estrai l'ufficio dal record (orgForResource o contact)
                org_list = source.get('orgForResource', source.get('OrgForResource', []))
                if org_list and isinstance(org_list, list):
                    # Prendi la prima organizzazione come nome ufficio
                    primo_org = org_list[0]
                    if isinstance(primo_org, dict):
                        valore_ufficio = primo_org.get('default', primo_org.get('langita', ''))
                    else:
                        valore_ufficio = str(primo_org)
                elif isinstance(org_list, str):
                    valore_ufficio = org_list
                
                # Scrittura nel foglio Excel
                sheet.cell(row=riga, column=idx_catalogo, value=link_trovato)
                sheet.cell(row=riga, column=idx_pub, value=pub_data if pub_data else None)
                sheet.cell(row=riga, column=idx_rev, value=rev_data if rev_data else None)
                if idx_ufficio:
                    sheet.cell(row=riga, column=idx_ufficio, value=valore_ufficio if valore_ufficio else None)
            else:
                # Nessuna corrispondenza trovata
                sheet.cell(row=riga, column=idx_catalogo, value="no")
                sheet.cell(row=riga, column=idx_pub, value=None)
                sheet.cell(row=riga, column=idx_rev, value=None)
                if idx_ufficio:
                    sheet.cell(row=riga, column=idx_ufficio, value=None)
            
            # Scrivi il visualizzatore e sostituisci l'UID con il link
            if idx_visualizzatore:
                sheet.cell(row=riga, column=idx_visualizzatore, value=tipo_visualizzatore if tipo_visualizzatore else None)
            if idx_uid and link_visualizzatore:
                sheet.cell(row=riga, column=idx_uid, value=link_visualizzatore)
                
            cronologia_elaborazione.append({
                "riga": riga,
                "nome_completo": valore_nome_completo_str,
                "nome": valore_nome_layer_str,
                "esito": esito,
                "titolo_metadato": titolo_trovato,
                "link": link_trovato,
                "pubblicazione": pub_data,
                "revisione": rev_data,
                "visualizzatore": tipo_visualizzatore,
                "ufficio": valore_ufficio,
                "punteggio": punteggio_massimo
            })
            
        # Salva il file modificato preservando la formattazione originale
        wb.save(percorso_elaborato)
        wb.close()
        
        # Elimina il file caricato originariamente per pulizia
        if os.path.exists(percorso_upload):
            os.remove(percorso_upload)
            
        return jsonify({
            "status": "successo",
            "fileId": nome_salvato,
            "risultati": cronologia_elaborazione
        })
        
    except Exception as e:
        # Pulisce i file in caso di errore
        if os.path.exists(percorso_upload):
            os.remove(percorso_upload)
        print(f"Errore durante l'elaborazione del file Excel: {e}")
        return jsonify({"errore": f"Errore durante l'elaborazione del file Excel: {str(e)}"}), 500


@app.route('/scarica/<file_id>', methods=['GET'])
def scarica_file(file_id):
    """Permette al client di scaricare il file Excel elaborato."""
    # Sicurezza: convalida il nome del file
    if not re.match(r'^[a-f0-9\-]{36}_.+\.xlsx$', file_id):
        return jsonify({"errore": "ID file non valido"}), 400
        
    percorso_file = os.path.join(app.config['PROCESSED_FOLDER'], file_id)
    if not os.path.exists(percorso_file):
        return jsonify({"errore": "File non trovato o scaduto"}), 404
        
    # Rimuovi il prefisso UUID dal nome del file per il download dell'utente
    nome_originale = file_id[37:]
    
    ancorato = request.args.get('ancorato', 'false').lower() == 'true'
    
    if ancorato:
        try:
            wb = openpyxl.load_workbook(percorso_file)
            sheet = wb.active
            
            # Identifica gli indici delle colonne basandosi sull'intestazione (riga 1)
            riga_intestazione = [str(cella.value).strip().lower() if cella.value is not None else "" for cella in sheet[1]]
            idx_catalogo = None
            if 'catalogo' in riga_intestazione:
                idx_catalogo = riga_intestazione.index('catalogo') + 1
                
            if idx_catalogo:
                font_link = Font(color="0563C1", underline="single")
                for riga in range(2, sheet.max_row + 1):
                    cell = sheet.cell(row=riga, column=idx_catalogo)
                    val = str(cell.value or "").strip()
                    if val.startswith("http://") or val.startswith("https://"):
                        # Assegna l'hyperlink come stringa (preserva l'URL completo con #)
                        cell.hyperlink = val
                        cell.value = "sì"
                        cell.font = font_link
            
            # Salva in un buffer in memoria
            buffer = io.BytesIO()
            wb.save(buffer)
            buffer.seek(0)
            wb.close()
            
            return send_file(
                buffer,
                as_attachment=True,
                download_name=nome_originale,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            print(f"Errore durante l'ancoraggio dei link nel download: {e}")
            # In caso di errore, fall back al file originale sul server
            pass
            
    return send_file(
        percorso_file,
        as_attachment=True,
        download_name=nome_originale,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('FLASK_ENV') != 'production'
    app.run(host='0.0.0.0', port=port, debug=debug)
