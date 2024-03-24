from docx import Document
import zipfile
import re
import os
from docx.shared import Cm, Pt, Length
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
from datetime import datetime
from docx.enum.text import WD_LINE_SPACING
from docx.shared import Mm
import win32com.client as win32
import pythoncom

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
    date_object = datetime.strptime(date_str, "%d/%m/%Y")
    formatted_date = date_object.strftime("%Y –%m –%d")

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

def keep_last_digits(input_string):
    numbers = ''.join(filter(str.isdigit, input_string))
    last_digits = numbers[-4:] 
    return last_digits

def configFirstLineTable(table):
    table.cell(0, 0).merge(table.cell(0, 1))
    table.cell(0, 0).text = 'MAX COAG 1'
    table.rows[0].height = Cm(0.85)

    first_cell = table.rows[0].cells[0]
    first_cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    first_row = first_cell.paragraphs[0]
    run = first_row.runs[0] if first_row.runs else first_row.add_run()
    run.font.size = Pt(13)
    first_row.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

def configSecondLineTable(table, type):
    id = keep_last_digits(type['id'])

    table.cell(1, 0).merge(table.cell(1, 1))
    table.cell(1, 0).text = f"ID:   {id}"
    table.rows[1].height = Cm(0.4)

    second_row = table.cell(1, 0).paragraphs[0]
    second_row.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    run = second_row.runs[0] if second_row.runs else second_row.add_run()
    run.font.size = Pt(11)   

def configThirdLineTable(table, date, time):
    table.rows[2].height = Cm(0.6)
    dateText = format_date(date)

    # Configurações de largura para cada coluna
    widths = [Cm(2.95), Cm(1.45)]

    for idx, (text, width) in enumerate(zip([dateText, time], widths)):
        cell = table.cell(2, idx)
        cell.width = width
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

        while len(cell.paragraphs) > 1:
            p = cell.paragraphs[1]._element
            p.getparent().remove(p)

        paragraph = cell.paragraphs[0]
        paragraph.clear()
        run = paragraph.add_run(text)
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        run.font.size = Pt(11)

def configFourthLineTable(table, texts):
    table.rows[3].height = Cm(0.6)

    for idx, text in enumerate(texts):
        cell = table.cell(3, idx)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM
 
        for paragraph in cell.paragraphs[1:]:
            p = paragraph._element
            p.getparent().remove(p)
        
        paragraph = cell.paragraphs[0]
        paragraph.clear()  
        run = paragraph.add_run(text)
        run.font.size = Pt(11)

def configFifthLineTable(table, text):
    table.rows[4].height = Cm(0.6)
    cell = table.cell(4, 0).merge(table.cell(4, 1))
    cell.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM

    for paragraph in cell.paragraphs[1:]:
        p = paragraph._element
        p.getparent().remove(p)

    paragraph = cell.paragraphs[0]
    paragraph.clear() 
    run = paragraph.add_run(text)
    run.font.size = Pt(11)

def configDocument(doc):
    style = doc.styles['Normal']
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    font = style.font
    font.name = 'Tinos'
    section = doc.sections[0]
    section.top_margin = Cm(1)
    section.bottom_margin = Cm(1)
    section.left_margin = Cm(0.9)
    section.right_margin = Cm(0.9)
    section.page_width = Mm(58)
    section.page_height = Mm(297)

def configTable(table):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    columns = table.columns
    num_columns = len(columns)
    per_cell_width = Cm(4.4 / num_columns)

    total_preferred_width = Cm(4.4)
    table.cell(2, 0).width = Cm(2.95)  
    table.cell(2, 1).width = total_preferred_width - table.cell(2, 0).width 
    table.cell(3, 0).width = Cm(1.5) 
    table.cell(3, 1).width = total_preferred_width - table.cell(3, 0).width 

    for row in table.rows:
        for cell in row.cells:
            cell.width = per_cell_width

def create_word_document_fibs(dataFib, date, time):
    doc = Document()
    configDocument(doc)
    timeFib = time

    for fib in dataFib:
        fibValue = fib['fib'].replace(",", ".")

        table = doc.add_table(rows=5, cols=2)
        configTable(table)

        configFirstLineTable(table)
        configSecondLineTable(table, fib)
        configThirdLineTable(table, date, timeFib)
        configFourthLineTable(table, ["FIB-C", f"{fibValue}      mg/dL"])
        configFifthLineTable(table, "    (   0. 0    – ‒ – ‒  0. 0    )")

        timeFib = add_minutes(timeFib)
        doc.add_paragraph() 
    
    doc.save("fib_data.docx")

    return timeFib

def add_top_border_to_fourth_row(doc_path):
    pythoncom.CoInitialize()  # Inicializa o COM no thread atual
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(doc_path)
    word.Visible = False

    try:
        for table in doc.Tables:
            if len(table.Rows) >= 4:
                for cell in table.Rows(4).Cells:
                    cell.Borders(win32.constants.wdBorderTop).LineStyle = win32.constants.wdLineStyleDashLargeGap
                    cell.Borders(win32.constants.wdBorderTop).LineWidth = win32.constants.wdLineWidth050pt 
                    cell.Borders(win32.constants.wdBorderTop).Color = win32.constants.wdColorAutomatic  
            
    except Exception as e:
        print(f"Erro ao adicionar borda à quarta linha: {e}")
    finally:
        doc.Save()
        doc.Close()
        word.Quit()

    pythoncom.CoUninitialize() 
    

def create_word_document_tp(dataTp, date, time):
    doc = Document()
    configDocument(doc)
    timeTp = time

    for tp in dataTp: 
        tpValue = tp['tp'].replace(",", ".")

        table = doc.add_table(rows=5, cols=2)
        configTable(table)

        configFirstLineTable(table)
        configSecondLineTable(table, tp)
        configThirdLineTable(table, date, timeTp)
        configFourthLineTable(table, ["PT", f"{tpValue}      S"])
        configFifthLineTable(table, "    (   0. 0    – ‒ – ‒  0. 0    )")

        timeTp = add_minutes(timeTp)
        doc.add_paragraph() 

    docx_path = "tp_data.docx"
    doc.save(docx_path)

    current_dir = os.path.dirname(os.path.abspath(__file__))
    docx_path_finally = os.path.join(current_dir, docx_path)
    add_top_border_to_fourth_row(docx_path_finally)

    return timeTp

def create_word_document_ttpa(dataTtpa, date, time):
    doc = Document()
    configDocument(doc)
    timeTtpa = time

    for ttpa in dataTtpa:
        ttpaValue = ttpa['ttpa'].replace(",", ".")

        table = doc.add_table(rows=5, cols=2)
        configTable(table)

        configFirstLineTable(table)
        configSecondLineTable(table, ttpa)
        configThirdLineTable(table, date, timeTtpa)
        configFourthLineTable(table, ["APTT", f"{ttpaValue}      S"])
        configFifthLineTable(table, "    (   0. 0    – ‒ – ‒  0. 0    )")

        timeTtpa = add_minutes(timeTtpa)
        doc.add_paragraph() 
    
    doc.save("ttpa_data.docx")

def create_zip_with_word_documents(word_filenames):
    with zipfile.ZipFile('animalsData.zip', 'w') as zip_file:
        for word_filename in word_filenames:
            zip_file.write(word_filename)
            os.remove(word_filename)
