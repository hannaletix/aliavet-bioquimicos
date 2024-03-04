from docx import Document
import zipfile
import re
import os

def extract_information_from_table_as_array(docx_file):
    tables_info = []

    for table in docx_file.tables:
        # Primeiro, adiciona os cabeçalhos como a primeira 'linha' da tabela
        headers = [cell.text.strip() for cell in table.rows[0].cells]
        tables_info.append(headers)

        # Em seguida, para cada linha subsequente, adiciona os valores das células
        for row in table.rows[1:]:
            row_values = [cell.text.strip() for cell in row.cells]
            tables_info.append(row_values)

    return tables_info

def find_collection_date(array_data):
    for item in array_data:
        # Cada 'item' é uma lista. Verifica se "Data de coleta:" está no primeiro elemento da lista.
        if item[0] == "Data de coleta:":
            return item[1]  # Retorna o segundo elemento da lista, que é o valor da data de coleta
    return "Data de coleta não encontrada"

def get_all_headers(array_data):
    headers = []
    for item in array_data:
        if item[0].startswith("Identificação"):
            headers.append(item)
    return headers if headers else "Identificação não encontrada"

def separate_by_identifier(data):
    separated_data = []
    identifiers_seen = set()
    current_data = []

    for item in data:
        identifier = item[0]
        if identifier in identifiers_seen:
            # Se o identificador atual já foi visto, começa um novo array
            separated_data.append(current_data)
            current_data = [item]  # Começa um novo array com o item atual
            identifiers_seen = {identifier}  # Reinicia o conjunto de identificadores vistos
        else:
            # Se o identificador atual não foi visto, apenas adiciona ao array atual
            current_data.append(item)
            identifiers_seen.add(identifier)

    # Adiciona o último array ao resultado final
    separated_data.append(current_data)

    return separated_data

def filter_and_organize_numeric_lines(array_data):
    numeric_lines = []

    for line in array_data:
        # Verifica se o primeiro elemento é numérico (contém apenas dígitos)
        if line[0].replace('.', '', 1).isdigit() or line[0].isdigit():
            numeric_lines.append(line)

    separeted_numeric_lines = separate_by_identifier(numeric_lines)

    return separeted_numeric_lines

def merge_data_with_same_identifier(data):
    merged_data = {}
    for item in data:
        identifier = item['Identificação']
        if identifier not in merged_data:
            merged_data[identifier] = item
        else:
            merged_data[identifier].update(item)
    return list(merged_data.values())

def organize_data(array_data): 
    headers = get_all_headers(array_data)
    filtered_lines = filter_and_organize_numeric_lines(array_data)

    formatted_data = []
    for i, header_row in enumerate(headers):
        for data_row in filtered_lines[i]:
            formatted_data.append({header_row[j]: data_row[j] for j in range(len(header_row))})

    formatted_data_by_id = merge_data_with_same_identifier(formatted_data)

    return formatted_data_by_id 

def format_key_value_pair(key, value):
    key = key.replace('\n', '')

    match = re.search(r'\((.*?)\)', key)
    if match:
        unit = match.group(1)
        key = re.sub(r'\(.*?\)', '', key)
        return f"{key}  {value} {unit}"
    else:
        return f"{key}  {value}"

def create_word_document_with_date(date, animalData):
    doc = Document()
    id = animalData['Identificação']
    doc.add_paragraph(f"ID: {id}")
    doc.add_paragraph(f"Time: {date}")

    for key, value in animalData.items():
        if key == 'Identificação':
            continue

        formatted_value = format_key_value_pair(key, value)
        doc.add_paragraph(f"{formatted_value}")

    doc_name = f"{id}_data.docx"
    doc.save(doc_name)
    return doc_name

def create_zip_with_word_documents(word_filenames):
    with zipfile.ZipFile('animalsData.zip', 'w') as zip_file:
        for word_filename in word_filenames:
            zip_file.write(word_filename)
            os.remove(word_filename)
