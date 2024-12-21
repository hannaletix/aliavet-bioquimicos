from flask import Flask, jsonify, request, render_template, send_file
from io import BytesIO
from docx import Document
import os
import random
from waitress import serve
import shutil
from helpers import extract_information_from_table_as_array, find_collection_date, data_processing, create_word_document, format_filename, get_number_laudo, changeStyle

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')  # Exibe o formulário de upload

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'fileUpload' not in request.files:
        return 'Nenhum arquivo parte do pedido', 400

    file = request.files['fileUpload']
    laudo_name = file.filename

    if laudo_name == '':
        return 'Nenhum arquivo selecionado', 400
    
    bio_file_name = format_filename(laudo_name)

    if file:
        file_stream = BytesIO(file.read())
        document = Document(file_stream)

        try:
            tables_infos = extract_information_from_table_as_array(document)
            date_collection = find_collection_date(tables_infos)
            data_final = data_processing(tables_infos, date_collection)
            template = Document("template.docx")
            create_word_document(data_final, date_collection, template, bio_file_name)

            # number_laudo = get_number_laudo(laudo_name)
  
            downloads_folder = r'C:\Users\Hanna\OneDrive\Documentos\Aliavet\Bioquimico\101.061-22\\'
            shutil.copy(bio_file_name, downloads_folder)

            changeStyle(downloads_folder, bio_file_name)

        except Exception as e:
            return str(e), 500

        return 'Arquivo recebido com sucesso', 200


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    serve(app, host="0.0.0.0", port=port)