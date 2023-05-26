from docx import Document
import pandas as pd
from docx2python import docx2python
###
#Open the document in Word, open the Visual Basic editor (F11), open the immediate window (ctrl-G), type the following macro and press enter:
#ActiveDocument.Range.ListFormat.ConvertNumbersToText
###
import docx
import pandas as pd
file_path = "C:/Users/NadavHa/Downloads/doc10.docx"
def containsNumber(value):
    return any([char.isdigit() for char in value])
def extract_terms_from_docx(file_path):
    global line_list
    line_list=[]

    doc = docx.Document(file_path)
    data = []
    term=''
    term_number =''
    term_subject=''
    term_letter=''
    no_letters=''
    for paragraph in doc.paragraphs:
        line=paragraph.text.split('\t')
        line_list.append(line)
        no_letters = sum(c.isdigit() for c in line[0])
        if len(line)==2 and no_letters<2:
            term_subject_1=line[1]
        if len(line)==2 and no_letters==2:
            term_number = line[0]
            term = line[1]
            term_subject_2=line[1]
            data.append((term_number, term, term_subject_1))
        if len(line)==2 and no_letters==3:
            data = (list(filter(lambda x: x[0] != line[0][:-2], data)))
            term_number = line[0]
            term = line[1]
            term_subject_3=line[1]
            data.append((term_number, term, term_subject_2))

        if len(line)==2 and no_letters==4:
            data = (list(filter(lambda x: x[0] != line[0][:-2], data)))
            term_number = line[0]
            term = line[1]
            term_subject_4=line[1]
            data.append((term_number, term, term_subject_3))
        if len(line)==2 and no_letters==5:
            data = (list(filter(lambda x: x[0] != line[0][:-2], data)))

       #     data=data[:-1]
            term_number = line[0]
            term = line[1]
            data.append((term_number, term, term_subject_4))
        else:
            term=term+' '+line[0]
    return data

data = extract_terms_from_docx(file_path)
df = pd.DataFrame(data, columns=["Term Number", "Term","subject"]).drop_duplicates()
df.to_excel("amnira.xlsx")
