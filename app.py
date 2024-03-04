from flask import Flask, request, render_template, send_file
from io import BytesIO
from docx import Document
import os
from helpers import extract_information_from_table_as_array, find_collection_date, organize_data, create_word_document_with_date, create_zip_with_word_documents

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('upload.html')  # Exibe o formulário de upload

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'fileUpload' not in request.files:
        return 'Nenhum arquivo parte do pedido', 400

    file = request.files['fileUpload']

    if file.filename == '':
        return 'Nenhum arquivo selecionado', 400

    if file:
        # Converte o arquivo carregado para um objeto BytesIO
        file_stream = BytesIO(file.read())
        
        # Usa o objeto BytesIO para criar um documento com python-docx
        document = Document(file_stream)

        # Extrair as informações das tabelas diretamente do documento
        try:
            tables_info = extract_information_from_table_as_array(document)
            date_collection = find_collection_date(tables_info)
            organized_data = organize_data(tables_info)

            docs_names = []
            for data in organized_data:
                doc_name = create_word_document_with_date(date_collection, data)
                docs_names.append(doc_name)

            create_zip_with_word_documents(docs_names)

        except Exception as e:
            return str(e), 500

        return 'Arquivo recebido com sucesso', 200

@app.route('/download_zip')
def download_zip():
    zip_filepath = 'animalsData.zip'

    if os.path.exists(zip_filepath):
        return send_file(zip_filepath, as_attachment=True)
    else:
        return 'Arquivo zip não encontrado', 404

if __name__ == '__main__':
    app.run(debug=True)
