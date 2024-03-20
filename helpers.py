from docx import Document
import zipfile
import re
import os
import random
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def extract_information_from_table_as_array(docx_file):
    tables_info = []

    for table in docx_file.tables:
        headers = [cell.text.strip() for cell in table.rows[0].cells]
        tables_info.append(headers)

        for row in table.rows[1:]:
            row_values = [cell.text.strip() for cell in row.cells]
            tables_info.append(row_values)

    return tables_info

def find_collection_date(array_data):
    for item in array_data:
        if item[0] == "Data de coleta:":
            return item[1]
    return "Data de coleta não encontrada"

def is_id(data):
    for char in data:
        if char.isdigit():
            return True
    return False

def organize_data(data):
    header = data[0]
    defaultValues = [r"Identificação", r"Fibrinogênio", r"TP", r"TTPA"]
    indices = [next((i for i, cabecalho in enumerate(header) if re.search(padrao, cabecalho)), None) for padrao in defaultValues]
    header = [header[i] for i in indices]

    for linha in data:
        linha[:] = [linha[i] for i in indices]

    return data

def extract_relevant_data(array_data): 
    relevant_data = []
    correct_data = False

    for row in array_data:
        if len(row) > 2:
            if 'Fibrinogênio' in row[1] or 'TP' in row[1] or 'TTPA' in row[1]:
                correct_data = True
            
            if correct_data:
                relevant_data.append([row[0], row[1], row[2], row[3]])

    organizedData = organize_data(relevant_data)

    return organizedData

def format_key_value_pair(key, value):
    key = key.replace('\n', '')

    match = re.search(r'\((.*?)\)', key)
    if match:
        unit = match.group(1)
        key = re.sub(r'\(.*?\)', '', key)
        return f"{key}  {value} {unit}"
    else:
        return f"{key}  {value}"

def format_date(date_str):
    parts = date_str.split('/')
    formatted_date = '-'.join([parts[2], parts[1], parts[0]])
    
    return formatted_date

def add_minutes(time):
    time_separate = time.split(":")
    hour = int(time_separate[0]) 
    minute = int(time_separate[1])  

    minute += 1

    if minute >= 60:
        hour += 1
        minute -= 60

        if hour >= 24:
            hour = 0

    return f"{hour:02d}:{int(round(minute)):02d}"

def insert_title_document(doc):
    p = doc.add_paragraph()
    run = p.add_run('MAX COAG 1')
    run.font.name = 'Bodoni MT'

def create_word_document_fibs(dataFib, date, time):
    doc = Document()
    timeFib = time

    for fib in dataFib:
        id = fib['id'][-4:]  # Seleciona apenas as 4 últimas strings
        fibValue = fib['fib'].replace(",", ".")

        insert_title_document(doc)
        doc.add_paragraph(f"ID: {id}")
        doc.add_paragraph(f"{format_date(date)}    {timeFib}")
        doc.add_paragraph(f"FIB-C   {fibValue}    mg/dL")
        doc.add_paragraph(f"    (0.0         0.0)   ")

        timeFib = add_minutes(timeFib)
    
    doc.save("fib_data.docx")

    return timeFib

def create_word_document_tp(dataTp, date, time):
    doc = Document()
    timeTp = time

    for tp in dataTp:
        id = tp['id'][-3:]  
        tpValue = tp['tp'].replace(",", ".")

        insert_title_document(doc)
        doc.add_paragraph(f"ID: {id}")
        doc.add_paragraph(f"{format_date(date)}    {timeTp}")
        doc.add_paragraph(f"PT   {tpValue}    S")
        doc.add_paragraph(f"    (0.0         0.0)   ")

        timeTp = add_minutes(timeTp)
    
    doc.save("tp_data.docx")

    return timeTp

def create_word_document_ttpa(dataTtpa, date, time):
    doc = Document()
    timeTtpa = time

    for ttpa in dataTtpa:
        id = ttpa['id'][-3:]
        ttpaValue = ttpa['ttpa'].replace(",", ".")

        insert_title_document(doc)
        doc.add_paragraph(f"ID: {id}")
        doc.add_paragraph(f"{format_date(date)}    {timeTtpa}")
        doc.add_paragraph(f"APTT   {ttpaValue}    S")
        doc.add_paragraph(f"    (0.0         0.0)   ")

        timeTtpa = add_minutes(timeTtpa)
    
    doc.save("ttpa_data.docx")

def create_zip_with_word_documents(word_filenames):
    with zipfile.ZipFile('animalsData.zip', 'w') as zip_file:
        for word_filename in word_filenames:
            zip_file.write(word_filename)
            os.remove(word_filename)
