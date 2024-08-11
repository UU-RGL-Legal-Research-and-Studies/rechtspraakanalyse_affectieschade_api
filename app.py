from flask import Flask, render_template, request, send_file, session, redirect, url_for
import requests
import xml.etree.ElementTree as ET
import pandas as pd
from io import BytesIO
import tempfile
from ECLI_affectieschade1 import unique_list  # Import a 'unique_list' function from another file
import re
from datetime import datetime
from flask_sqlalchemy import SQLAlchemy

# Create a Flask application and set a secret key for sessions
app = Flask(__name__)
app.secret_key = 'hello_world'

# Configure SQLAlchemy with a SQLite database
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///ecli_cache.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

# Define a database model for storing ECLI data
class ECLIEntry(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    ecli = db.Column(db.String, unique=True, nullable=False)
    xml_content = db.Column(db.Text, nullable=False)
    identifier_link = db.Column(db.String)
    date_link = db.Column(db.Date)

# Define and initialize some variables and data structures
ECLIs = unique_list  # A list of ECLIs retrieved from the other file
ECLIs.sort(reverse=True)  # Sort the ECLIs in reverse order
ECLI_texts = {}  # An empty dictionary to store text data for each ECLI

# Define a function to highlight search terms in text
def highlight_term(text, term):
    return text.replace(term, f'<span class="highlight">{term}</span>')

# Define a function to make an API request for a specific ECLI
def api_request(ecli):
    # Define XML namespaces
    namespaces = {
        'rdf': "http://www.w3.org/1999/02/22-rdf-syntax-ns#",
        'dcterms': "http://purl.org/dc/terms/",
    }

    # Check if the ECLI is already in the database
    entry = ECLIEntry.query.filter_by(ecli=ecli).first()
    if entry:
        # If it is, retrieve the XML data from the database
        root = ET.fromstring(entry.xml_content)
        identifier_link = entry.identifier_link
        date_link = entry.date_link
    else:
        # If the ECLI is not in the database, fetch it via the API and save it in the database
        url = f"https://data.rechtspraak.nl/uitspraken/content?id={ecli}"
        response = requests.get(url, stream=True)
        root = ET.fromstring(response.content)

        # Look for the dcterms:identifier tags and retrieve the identifier link
        identifier_link = None
        for identifier_tag in root.findall('.//rdf:Description/dcterms:identifier', namespaces):
            if identifier_tag.text and identifier_tag.text.startswith('http'):
                identifier_link = identifier_tag.text
                break

        date_link = None
        for date_tag in root.findall('.//rdf:Description/dcterms:issued', namespaces):
            if date_tag.text:
                try:
                    date_link = datetime.strptime(date_tag.text, '%Y-%m-%d').date()
                except ValueError:
                    date_link = None

        # Store the XML content and other data in the database
        new_entry = ECLIEntry(ecli=ecli, xml_content=ET.tostring(root, encoding='unicode'), identifier_link=identifier_link, date_link=date_link)
        db.session.add(new_entry)
        db.session.commit()

    return root, identifier_link, date_link  # Return the XML root, identifier link, and date link

# Define a route for the main page ("/") of the web application
@app.route('/', methods=['GET', 'POST'])
def index():
    global ECLI_texts  # Indicate that ECLI_texts should be treated as a global variable
    search_results_count = 0  # Initialize the number of search results

    if request.method == 'POST':
        # Get raw search terms from the user and process them
        raw_search_terms = request.form.get('search_terms').split(',')
        search_terms = [term.split('|') for term in raw_search_terms]
        ECLI_texts = {}  # Clear the ECLI_texts dictionary for new search results

        # Iterate over all ECLIs and perform search queries
        for ecli in ECLIs:
            root, identifier_link, date_link = api_request(ecli)  # Fetch the XML data for the current ECLI
            texts = [elem.text for elem in root.iter() if elem.text and all(
                any(synonym.lower() in elem.text.lower() for synonym in term)
                for term in search_terms
            )]  # Search for text that meets the search criteria
            highlighted_texts = []  # List to store highlighted text
            for text in texts:
                for term_group in search_terms:
                    for synonym in term_group:
                        text = highlight_term(text, synonym)  # Highlight the search terms in the text
                highlighted_texts.append(text)
            ECLI_texts[ecli] = {'texts': highlighted_texts, 'identifier_link': identifier_link, 'current_index': 0, 'date_link': date_link}
            search_results_count += len(highlighted_texts)  # Update the number of search results

        # Update the Excel file
        update_excel_file()

        # Sla de scrollpositie op in de sessie
        session['scrollPosition'] = request.form.get('scrollPosition', 0)

        # Redirect naar dezelfde pagina met een GET-verzoek
        return redirect(url_for('index'))

    # Voor een GET-verzoek
    scroll_position = session.get('scrollPosition', 0)

    # Sort ECLI_text op datum, waarbij None-waarden worden behandeld
    sorted_ECLI_texts = dict(sorted(ECLI_texts.items(), key=lambda item: (item[1]['date_link'] is None, item[1]['date_link']), reverse=True))

    return render_template('index.html', ECLI_texts=sorted_ECLI_texts, search_results_count=search_results_count, scroll_position=scroll_position)

# Define a function to remove HTML tags from a string
def remove_html_tags(text):
    clean = re.compile('<.*?>')
    return re.sub(clean, '', text)

# Define a function to update the Excel file with search results
def update_excel_file():
    data = []  # A list of tuples for each row in the DataFrame
    for ecli, texts in ECLI_texts.items():
        date = texts.get('date_link', 'No date available')
        result_text = remove_html_tags(texts['texts'][ECLI_texts[ecli]['current_index']]) if texts['texts'] else 'none'
        link = texts.get('identifier_link', 'No link available')
        data.append((date, ecli, link, result_text))  # Add data to the list

    df = pd.DataFrame(data, columns=['Date', 'ECLI', 'Link', 'Result'])  # Create a DataFrame from the data

    # Save the DataFrame as an Excel file
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
        df.to_excel(temp_file, index=False)

    session['temp_excel_file'] = temp_file.name  # Save the path to the Excel file in the session
    session.modified = True

# Define a route to download the Excel file ("/download/excel")
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

# Ensure the database is created before the first request
@app.before_first_request
def initialize_database():
    db.create_all()

# Start the Flask application if the script is run directly
if __name__ == '__main__':
    app.run(debug=True)
