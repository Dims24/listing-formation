import os
import sys
from argparse import ArgumentParser
from docx.shared import Pt
import docx


def list_files(path,ignore,hr,file_name):
    create_doc(hr,file_name)
    filelist = []
    for root, dirs, files in os.walk(path):
        for file in files:
            filelist.append(os.path.join(root, file))
    check_ignore(filelist,ignore)
    entry(path,filelist,hr,file_name)

def create_doc(hr,file_name):
    doc = docx.Document()
    par = doc.add_paragraph()
    par.add_run(f'Приложение {hr}.').bold = True
    par.alignment = 1
    doc.save(f'{file_name}.docx')


def check_ignore(filelist,ignore):
    print(filelist[0])
    if ignore==None:
        return filelist
    else:
        return filelist
        # ignore=ignore.split()
        # for check in ignore:
        #     for mark in filelist:
        #         if check in mark:
        #             print(1)

def entry (path,filelist,hr,file_name):

    for name in filelist:
        name1 = name.replace(f'{path}\\', "")
        name=name.replace("\\","\\\\")

        # print(open('C:\\Users\\Admin\\PycharmProjects\\pythonProject\\выаыва\\counter.py').read())
        dock_formation(name1, name,hr,file_name)
i=1
def dock_formation(name1,name,hr,file_name):
    global i
    doc = docx.Document(f'{file_name}.docx')



    doc.add_paragraph(f'Листинг {hr}.{i} - {name1}')
    i+=1
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'
    info = open(name,encoding="utf8").read().strip()
    table.cell(0, 0).text = info
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    doc.add_page_break()
    doc.save(f'{file_name}.docx')




parser = ArgumentParser(description="Formation of the program listing")
parser.add_argument("-td", "--tirgert_dir", dest="path",required=True,
                        help="Directory path")
parser.add_argument("-ig", "--ignore_dir", dest="ignore",default=None,
                        help="Ignore directory")
parser.add_argument("-hr", "--header", dest="hr", default="А",
                        help="Application number")
parser.add_argument("-n", "--name", dest="file_name",required=True,
                        help="File name")

args = parser.parse_args()
list_files(args.path,args.ignore,args.hr,args.file_name)



if __name__ == '__main__':
    pass
    # path = input("Введите путь:")
    # listing=input("папки")
    # ignore=listing.split()
    # print(listing)
    # list_files(path,listing)

# C:\Users\Admin\PycharmProjects\pythonProject
# C:\Users\Admin\AppData\Local\Programs\Python\Python310\python.exe