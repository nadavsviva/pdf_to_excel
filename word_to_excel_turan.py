from docx import Document
import pandas as pd
from docx2python import docx2python

import docx
import pandas as pd
import re
###
#Open the document in Word, open the Visual Basic editor (F11), open the immediate window (ctrl-G), type the following macro and press enter:
#ActiveDocument.Range.ListFormat.ConvertNumbersToText

###
file_path = "C:/Users/NadavHa/Downloads/doc6.docx"
def containsNumber(value):
    return any([char.isdigit() for char in value])
def extract_terms_from_docx(file_path):
    doc = docx.Document(file_path)
    pattern = r'^([\d\.]+|[א-ת]\.)'
    data = []
    term=''
    term_number =''
    term_subject=''
    term_letter=''
    for paragraph in doc.paragraphs:
        line=paragraph.text.strip().split('\t')
        match = re.match(pattern, line[0])
        no_letters = sum(c.isdigit() for c in line[0])
        print(line)
        if len(line)==2 and containsNumber(line[0])==True:
            data.append((term_number, term, term_subject,term_letter))
            term_number = line[0]
            term = line[1]
        elif len(line)==2 and match:
            term_subject=line[1]
            term_letter=line[0]
        else:
            term=term+' '+line[0]
    return data

data = extract_terms_from_docx(file_path)
df = pd.DataFrame(data, columns=["Term Number", "Term","subject","letter"]).drop_duplicates()
df["term_chapter"]=df["letter"]+df["Term Number"]
df.to_excel("turan.xlsx")
