import os
import sys
from argparse import ArgumentParser
from docx.shared import Pt,Mm
import docx
from progress.bar import IncrementalBar
import math


def list_files(path,ignore,hr,file_name,intcount):
    create_doc(hr,file_name)
    filelist = []
    for root, dirs, files in os.walk(path):
        for file in files:
            filelist.append(os.path.join(root, file))
    check_ignore(path,filelist,ignore,hr,file_name,intcount)

o=1

def create_doc(hr,file_name):
    global o
    doc = docx.Document()
    par = doc.add_paragraph()
    par.add_run(f'Приложение {hr}').bold = True
    par.alignment = 1
    doc.save(f'{file_name} - {o}.docx')


def check_ignore(path,filelist,ignore,hr,file_name,intcount):
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
        count=len(filelist)
        # print(filelist,len(filelist))
        intcount=breaking(count, intcount)
        crutch(path, filelist, hr, file_name,count,intcount)

it=0
def crutch(path, filelist, hr, file_name,count,intcount):
    it=count
    bar = IncrementalBar('Loading...', max=it, suffix=f' %(index).d/%(max).d - %(percent).1f%% - %(elapsed).ds')
    while it>0:
        it=it-1
        entry(path, filelist, hr, file_name, count,bar,intcount)
    bar.finish()
    return True

def breaking(count,intcount):
    intcount=math.ceil((count/int(intcount)))
    return intcount

j=0
def entry (path,filelist,hr,file_name,count,bar,intcount):
    global j
    doc = docx.Document(f'{file_name} - {o}.docx')
    if len(filelist) == 0:
        return True
    for name in filelist:
        if j >= int(intcount):
            create_doc1(hr, file_name, path, filelist,count,bar,intcount)
        else:
            filelist.remove(name)
            j += 1
            name1 = name.replace(f'{path}\\', "")
            name = name.replace("\\", "\\\\")
            dock_formation(doc, name1, name, hr,bar)
            doc.save(f'{file_name} - {o}.docx')
            bar.next()



def create_doc1(hr,file_name,path, filelist,count,bar,intcount):
    global o
    global j
    o += 1
    j = 0
    doc = docx.Document()
    par = doc.add_paragraph()
    par.add_run(f'Приложение {hr}').bold = True
    par.alignment = 1
    doc.save(f'{file_name} - {o}.docx')
    entry(path, filelist, hr, file_name,count,bar,intcount)




i=1
def dock_formation(doc,name1,name,hr,bar):
    global i
    p =doc.add_paragraph(f'Листинг {hr}.{i} - {name1}')
    i+=1
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'
    try:
        info = open(name,encoding="utf8",errors='ignore').read().strip()
        table.cell(0, 0).text = info
    except:
        table.cell(0, 0).text = "Измените кодировку"
    fmt = p.paragraph_format
    fmt.space_before = Mm(3)
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(14)

    # doc.add_page_break()


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
    parser.add_argument("-n", "--num", dest="intcount", required=True,
                        help="Number of files to split into")
    args = parser.parse_args()

    list_files(args.path, args.ignore, args.hr, args.file_name,args.intcount)



# C:\Users\Admin\PycharmProjects\pythonProject
# C:\Users\Admin\AppData\Local\Programs\Python\Python310\python.exe