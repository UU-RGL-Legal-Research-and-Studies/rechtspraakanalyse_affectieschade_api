from flask import Flask, render_template, request, send_file, session, redirect, url_for
import requests
import xml.etree.ElementTree as ET
import pandas as pd
from io import BytesIO
import tempfile
from ECLI_affectieschade1 import unique_list
from ECLI_uitkoopprocedure import unique_list1
import pickle
import shutil
import os
import time

app = Flask(__name__)
app.secret_key = 'hello_world'

# # Sample ECLIs
ECLIs = unique_list #list from other file where ECLIs have been subtracted (using Webparsing)
ECLIs.sort(reverse=True)
ECLI_texts = {} # ECLI_texts[ecli] = {'texts': [], 'current_index': 0}
ECLI_cache = {}  # Cache for XML roots

def highlight_term(text, term):
    #return text.replace(term, f'<span class="highlight">{term}</span>')
    return text

def get_case_number(ecli):
    root, zaaknummer = api_request(ecli)
    return zaaknummer

def api_request(ecli):
    if ecli in ECLI_cache:
        # Laad de XML-root van het tijdelijke bestand
        with open(ECLI_cache[ecli], 'rb') as temp_file:
            root = ET.parse(temp_file).getroot()
        return root, None

    # Voer de API-aanvraag uit
    url = f"https://data.rechtspraak.nl/uitspraken/content?id={ecli}"
    response = requests.get(url, stream=True)

    # Maak een tijdelijk bestand en schrijf de respons erin
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xml') as temp_file:
        shutil.copyfileobj(response.raw, temp_file)

    # Sla het pad naar het tijdelijke bestand op in de cache
    ECLI_cache[ecli] = temp_file.name

    # Laad de XML-root van het tijdelijke bestand
    with open(temp_file.name, 'rb') as temp_file:
        root = ET.parse(temp_file).getroot()

    # Extraheer het zaaknummer
    zaaknummer_elem = root.find(".//psi:zaaknummer", namespaces={"psi": "http://psi.rechtspraak.nl/"})
    zaaknummer = zaaknummer_elem.text if zaaknummer_elem is not None else "N/A"

    return root, zaaknummer

@app.route('/start_request', methods=['POST'])
def start_request():
    # Simulate API request
    time.sleep(5)  # Simulate a long API request
    return '', 200  # Return a success response

@app.route('/', methods=['GET', 'POST'])
def index():
    global ECLI_texts  # Declare ECLI_texts as global

    # Bepaal de geselecteerde database op basis van de sessie-variabele
    selected_db = session.get('selected_db', 'db1')
    if selected_db == 'db1':
        active_ECLIs = unique_list
    else:
        active_ECLIs = unique_list1

    ECLI_data = {}  # Store both texts and case numbers
    total_search_results = 0  # Initialize search results count

    if request.method == 'POST':
        # Verkrijg ruwe zoektermen van de gebruiker
        raw_search_terms = request.form.get('search_terms').split(',')
        # Verwerk de ruwe zoektermen
        search_terms = [term.split('|') for term in raw_search_terms]

        for ecli in active_ECLIs:  # Gebruik de actieve ECLIs op basis van de geselecteerde database
            root, zaaknummer = api_request(ecli)  # Get the XML root and the case number
            # Verzamel alle teksten die voldoen aan de zoekcriteria
            texts = [elem.text for elem in root.iter() if elem.text and all(
                any(synonym.lower() in elem.text.lower() for synonym in term)
                for term in search_terms
            )]
            # Voeg de gemarkeerde tekst toe
            highlighted_texts = []
            for text in texts:
                for term_group in search_terms:
                    for synonym in term_group:
                        text = highlight_term(text, synonym)
                highlighted_texts.append(text)
            ECLI_data[ecli] = {
                'texts': highlighted_texts if highlighted_texts else ['No result'],
                'current_index': 0,
                'zaaknummer': zaaknummer  # Store the case number
            }
            total_search_results += len(highlighted_texts)  # Update search results count
    return render_template('index.html', ECLI_texts=ECLI_data, search_results_count=total_search_results, selected_db=selected_db)

@app.route('/switch_db/<string:db_name>')
def switch_db(db_name):
    if db_name in ['db1', 'db2']:
        session['selected_db'] = db_name
    return redirect(url_for('index'))

def update_excel_file():
    data = []
    for ecli, texts in ECLI_texts.items():
        row = (ecli, texts['zaaknummer'], texts['texts'][texts['current_index']] if texts['texts'] else 'none')
        data.append(row)
    df = pd.DataFrame(data, columns=['ECLI', 'Case Number', 'Result'])

    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
        df.to_excel(temp_file, index=False)

    session['temp_excel_file'] = temp_file.name
    session.modified = True


# Voeg ergens in je code een functie toe om het 'vorige' en 'volgende' element te krijgen
def get_sibling_elements(ecli, index):
    root = api_request(ecli)
    current_elem = root.find(f".//*[@index='{index}']")
    prev_elem = current_elem.getprevious() if current_elem is not None else None
    next_elem = current_elem.getnext() if current_elem is not None else None
    return prev_elem, next_elem

@app.route('/previous/<ecli>', methods=['GET'])
def previous(ecli):
    if ecli in ECLI_texts:
        ECLI_texts[ecli]['current_index'] = max(0, ECLI_texts[ecli]['current_index'] - 1)
    update_excel_file()
    return redirect(url_for('index'))

@app.route('/next/<ecli>', methods=['GET'])
def next(ecli):
    if ecli in ECLI_texts:
        ECLI_texts[ecli]['current_index'] = min(len(ECLI_texts[ecli]['texts']) - 1, ECLI_texts[ecli]['current_index'] + 1)
    update_excel_file()
    return redirect(url_for('index'))

@app.route('/delete/<ecli>', methods=['GET'])
def delete(ecli):
    if ecli in ECLI_texts:
        del ECLI_texts[ecli]['texts'][ECLI_texts[ecli]['current_index']]
        ECLI_texts[ecli]['texts'] = [text for text in ECLI_texts[ecli]['texts'] if text]  # remove any empty strings which may occur due to deletion
    update_excel_file()
    return redirect(url_for('index'))

@app.route('/download/excel', methods=['GET'])
def download_excel():
    temp_excel_file = session.get('temp_excel_file', None)
    if temp_excel_file:
        return send_file(
            temp_excel_file,
            as_attachment=True,
            download_name='ECLI_results.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        return "No Excel file generated", 404

# Vergeet niet om de tijdelijke bestanden op te ruimen wanneer u klaar bent!
for temp_file_path in ECLI_cache.values():
    os.remove(temp_file_path)

if __name__ == '__main__':
    app.run(debug=True)
