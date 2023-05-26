from docx import Document
import pandas as pd
from docx2python import docx2python

import docx
import pandas as pd
###
#Open the document in Word, open the Visual Basic editor (F11), open the immediate window (ctrl-G), type the following macro and press enter:
#ActiveDocument.Range.ListFormat.ConvertNumbersToText

###
file_path = "C:/Users/NadavHa/Downloads/amiash.docx"
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
    for paragraph in doc.paragraphs:
        line=paragraph.text.strip().split('\t')
        line_list.append(line)
        no_letters = sum(c.isdigit() for c in line[0])
        print(line)
        if len(line)==2 and no_letters==1 and '(' not in line[0]:
            term_subject_1=line[1]
        if len(line)==2 and no_letters>=2:
            term_number = line[0]
            term = line[1]
            term_subject_2=line[1]
            term_subject_2=line[1]

            data.append((term_number, term, term_subject_1))
        elif len(line)==2 and no_letters==1 and '(' in line[0]:
            term_number = term_number_2+line[0]
            term=line[1]
            data.append((term_number, term, term_subject_2))
        else:
            term=term+' '+line[0]
    return data

data = extract_terms_from_docx(file_path)
df = pd.DataFrame(data, columns=["Term Number", "Term","subject"]).drop_duplicates()
df.to_excel("tyrex3c.xlsx")
