import math
import os
from argparse import ArgumentParser

import docx
from docx.shared import Pt, Mm
from progress.bar import IncrementalBar


def list_files(path, ignore, hr, file_name, number):
    if os.path.exists(path):
        create_doc(hr, file_name)
        filelist = []
        for root, dirs, files in os.walk(path):
            for file in files:
                filelist.append(os.path.join(root, file))
        check_ignore(path, filelist, ignore, hr, file_name, number)
    else:
        print(f'Каталога \'{path}\' не существует')


o = 1


def create_doc(hr, file_name):
    global o
    doc = docx.Document()
    par = doc.add_paragraph()
    par.add_run(f'Приложение {hr}').bold = True
    par.alignment = 1
    doc.save(f'{file_name} - {o}.docx')


def check_ignore(path, file_lists, ignore, hr, file_name, number):
    newlist = []
    if ignore == None:
        entry(path, file_lists, hr, file_name)
    else:
        ignore = ignore.split()
        for file_list in file_lists:
            found = False
            for j in ignore:
                if f'\\{j}\\' in file_list:
                    found = j
                    break
            if found:
                continue
            else:
                newlist.append(file_list)
        file_lists = newlist
        count = len(file_lists)
        number = breaking(count, number)
        crutch(path, file_lists, hr, file_name, count, number)


it = 0


def crutch(path, filelist, hr, file_name, count, number):
    it = count
    bar = IncrementalBar('Loading...', max=it, suffix=f' %(index).d/%(max).d - %(percent).1f%% - %(elapsed).ds')
    while it > 0:
        try:
            it = it - 1
            entry(path, filelist, hr, file_name, count, bar, number)
        except:
            continue
    bar.finish()


def breaking(count, number):
    number = math.ceil((count / int(number)))
    return number


j = 0


def entry(path, filelist, hr, file_name, count, bar, number):
    global j
    doc = docx.Document(f'{file_name} - {o}.docx')
    if len(filelist) == 0:
        return True
    for name in filelist:

        if j >= int(number):
            create_doc1(hr, file_name, path, filelist, count, bar, number)
        else:
            filelist.remove(name)
            j += 1
            name1 = name.replace(f'{path}\\', "")
            name = name.replace("\\", "\\\\")
            dock_formation(doc, name1, name, hr, bar)
            doc.save(f'{file_name} - {o}.docx')
            bar.next()


def create_doc1(hr, file_name, path, filelist, count, bar, number):
    global o
    global j
    o += 1
    j = 0
    doc = docx.Document()
    doc.save(f'{file_name} - {o}.docx')
    entry(path, filelist, hr, file_name, count, bar, number)


i = 1


def dock_formation(doc, name1, name, hr, bar):
    global i
    p = doc.add_paragraph(f'Листинг {hr}.{i} - {name1}')
    try:
        i += 1
        table = doc.add_table(rows=1, cols=1)
        table.style = 'Table Grid'
        info = open(name, encoding="utf8", errors='ignore').read().strip()
        table.cell(0, 0).text = info
        fmt = p.paragraph_format
        fmt.space_before = Mm(3)
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Times New Roman'
        font.size = Pt(14)
    except:
        delete_paragraph(p)
        i -= 1
        all_tables = doc.tables
        for active_table in all_tables:
            if active_table.cell(0, 0).paragraphs[0].text == '':
                active_table._element.getparent().remove(active_table._element)


def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None


if __name__ == '__main__':
    parser = ArgumentParser(description="Formation of the program listing")
    parser.add_argument("-td", "--target_dir", dest="path", required=True,
                        help="Путь к директории")
    parser.add_argument("-ig", "--ignore_dir", dest="ignore", default=".git",
                        help="Игнорируемые элементы")
    parser.add_argument("-hr", "--header", dest="hr", default="А",
                        help="Application number")
    parser.add_argument("-o", "--output", dest="file_name", required=True,
                        help="Название файла")
    parser.add_argument("-n", "--num", dest="number", required=True,
                        help="Количество файлов для разделения")
    args = parser.parse_args()

    list_files(args.path, args.ignore, args.hr, args.file_name, args.number)
