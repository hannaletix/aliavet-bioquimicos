from flask import Flask, request, render_template, send_file
from io import BytesIO
from docx import Document
import os
import random
from waitress import serve
from helpers import extract_information_from_table_as_array, find_collection_date, extract_relevant_data, create_zip_with_word_documents, create_word_document_fibs, create_word_document_tp, create_word_document_ttpa

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')  # Exibe o formulário de upload

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'fileUpload' not in request.files:
        return 'Nenhum arquivo parte do pedido', 400

    file = request.files['fileUpload']

    if file.filename == '':
        return 'Nenhum arquivo selecionado', 400

    if file:
        file_stream = BytesIO(file.read())
        
        # Usa o objeto BytesIO para criar um documento com python-docx
        document = Document(file_stream)

        try:
            tables_info = extract_information_from_table_as_array(document)
            date_collection = find_collection_date(tables_info)
            relevant_data = extract_relevant_data(tables_info)   # Extrai apenas os dados úteis
            relevant_data.pop(0)

            first_exam_hour = random.randint(14, 17)
            first_exam_minutes = random.randint(0, 59)
            exam_time = str(first_exam_hour) + ':' + str(first_exam_minutes)

            docs_names = ["fib_data.docx", "tp_data.docx", "ttpa_data.docx"]
            data_fib = []
            data_tp = []
            data_ttpa = []

            for data in relevant_data:
                data_fib.append({"id": data[0], "fib": data[1]})
                data_tp.append({"id": data[0], "tp": data[2]})
                data_ttpa.append({"id": data[0], "ttpa": data[3]})

            last_time_fib = create_word_document_fibs(data_fib, date_collection, exam_time)
            last_time_tp = create_word_document_tp(data_tp, date_collection, last_time_fib)
            create_word_document_ttpa(data_ttpa, date_collection, last_time_tp)

            
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

if __name__ == "__main__":
    # Obtém a porta da variável de ambiente do Heroku, ou define como 5000 se não estiver disponível
    port = int(os.environ.get("PORT", 5000))
    serve(app, host="0.0.0.0", port=port)