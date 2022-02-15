import os
import sys
from argparse import ArgumentParser
from docx.shared import Pt,Mm
import docx
from progress.bar import IncrementalBar


def list_files(path,ignore,hr,file_name):
    create_doc(hr,file_name)
    filelist = []
    for root, dirs, files in os.walk(path):
        for file in files:
            filelist.append(os.path.join(root, file))
    check_ignore(path,filelist,ignore,hr,file_name)



def create_doc(hr,file_name):
    doc = docx.Document()
    par = doc.add_paragraph()
    par.add_run(f'Приложение {hr}').bold = True
    par.alignment = 1
    doc.save(f'{file_name}.docx')


def check_ignore(path,filelist,ignore,hr,file_name):
    newlist = []
    if ignore==None:
        entry(path, filelist, hr, file_name)
    else:
        ignore = ignore.split()
        for i in filelist:
            found = False
            for j in ignore:
                if f'\\{j}\\' in i:
                    found = j
                    break
            if found:
                continue
            else:
                newlist.append(i)
        filelist=newlist
        print(filelist,len(filelist))
        entry(path, filelist, hr, file_name)



def entry (path,filelist,hr,file_name):
    bar = IncrementalBar('Progress', max=len(filelist))
    doc = docx.Document(f'{file_name}.docx')
    for name in filelist:
        bar.next()
        name1 = name.replace(f'{path}\\', "")
        name=name.replace("\\","\\\\")
        dock_formation(doc,name1, name,hr,file_name)
    bar.finish()
    doc.save(f'{file_name}.docx')

i=1
def dock_formation(doc,name1,name,hr,file_name):
    global i
    # doc = docx.Document(f'{file_name}.docx')
    p =doc.add_paragraph(f'Листинг {hr}.{i} - {name1}')
    i+=1
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'
    try:
        # table = doc.add_table(rows=1, cols=1)
        # table.style = 'Table Grid'
        info = open(name,encoding="utf8",errors='ignore').read().strip()
        table.cell(0, 0).text = info
    except:
        # table = doc.add_table(rows=1, cols=1)
        # table.style = 'Table Grid'
        table.cell(0, 0).text = "Измените кодировку"
    fmt = p.paragraph_format
    fmt.space_before = Mm(3)
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(14)
    # doc.add_page_break()
    # doc.save(f'{file_name}.docx')


if __name__ == '__main__':
    parser = ArgumentParser(description="Formation of the program listing")
    parser.add_argument("-td", "--tirgert_dir", dest="path", required=True,
                        help="Directory path")
    parser.add_argument("-ig", "--ignore_dir", dest="ignore", default=".git",
                        help="Ignore directory")
    parser.add_argument("-hr", "--header", dest="hr", default="А",
                        help="Application number")
    parser.add_argument("-o", "--output", dest="file_name", required=True,
                        help="File name")
    args = parser.parse_args()

    list_files(args.path, args.ignore, args.hr, args.file_name)


# C:\Users\Admin\PycharmProjects\pythonProject
# C:\Users\Admin\AppData\Local\Programs\Python\Python310\python.exe