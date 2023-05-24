import PyPDF2
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import re
import os


def extract_segments(content):
    global segments
    global lines
    global site_no
    segments = {}
    lines = content.split('\n')
    pattern = r'^\d+(\.\d+)*\.?$'
    for line in lines:
        line = line.strip()
        first_string=line.split(' ')[0]
        match = re.match(pattern, first_string)
        print(match)
        if match:
            segment_no=first_string
            rest_line=line.partition(' ')[2]
            segments.update({segment_no:rest_line})
        elif "מספר הגנת הסביבה" in line:
            site_no = [int(i) for i in line.split() if i.isdigit()][0]
        else:     
            try:
                segments[segment_no]=segments[segment_no]+' '+line
            except UnboundLocalError:
                pass
    print(segments)
    return segments

def convert_pdf_to_table(file_path):
    pdf_file = open(file_path, 'rb')
    pdf_reader = PyPDF2.PdfReader(pdf_file)

    content = ""
    for page in pdf_reader.pages:
        content += page.extract_text()

    segments = extract_segments(content)
    return segments

def save_table_to_excel(segments):
    df=pd.DataFrame(segments.items(), columns=['מספר סעיף בתנאים/בחוק', 'נוסח הסעיף בחוק/בהיתר'])
    df['מספר אתר סביבתי']=site_no
    df=df[['נוסח הסעיף בחוק/בהיתר','מספר אתר סביבתי','מספר סעיף בתנאים/בחוק']]
    save_path = filedialog.asksaveasfilename(defaultextension='.xlsx', initialfile=pdf_file_name, filetypes=[('Excel Files', '*.xlsx'), ('All Files', '*.*')])
    df.to_excel(save_path, index=False)
    if save_path:
        # Get the file extension from the save_path
        file_extension = os.path.splitext(save_path)[1]

        # Check if the file extension is not ".xlsx"
        if file_extension.lower() != ".xlsx":
            # Append ".xlsx" to the save_path
            save_path += ".xlsx"

        df.to_excel(save_path, index=False)




def browse_pdf_file():
    global pdf_file_name
    file_path = filedialog.askopenfilename(filetypes=[('PDF Files', '*.pdf')])
    pdf_file_name = os.path.basename(file_path).split('.')[0]+'.xlsx'
    if file_path:
        segments = convert_pdf_to_table(file_path)
        save_table_to_excel(segments)
        print("Table generated and saved to Excel successfully.")

# Create the Tkinter GUI window
window = tk.Tk()
window.title("PDF License Converter")

# Create the Browse button
browse_button = tk.Button(window, text="Browse PDF File", command=browse_pdf_file)
browse_button.pack(pady=20)

window.mainloop()
