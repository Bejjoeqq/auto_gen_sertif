from docx import Document
from docx2pdf import convert
import pandas,os
def nameser(nama):
    names = nama
    if len(nama)>20:
        names = ""
        lens = 0
        phrase = nama.split()
        for x in phrase:
            lens += len(x)
            if lens>20:
                names += x[:1] + ". "
            else:
                names += x + " "
    return names

def replace_string(filename,nama):
    doc = Document(filename)
    for p in doc.paragraphs:
        if '<name>' in p.text:
            inline = p.runs
            for i in range(len(inline)):
                if '<name>' in inline[i].text:
                    text = inline[i].text.replace('<name>', nama)
                    inline[i].text = text
            print(p.text)

    try:
        doc.save(f'hasil/{nama}.docx')
    except Exception as e:
        print(e)

if __name__ == '__main__':
    # df1=pandas.read_excel(os.path.join("daftar.xlsx"),engine='openpyxl')
    # for x in df1["Nama"]:
    #     if type(x)==str:
    #         replace_string("template.docx",nameser(x.upper()))
    # print("Done")
    convert("hasil/")