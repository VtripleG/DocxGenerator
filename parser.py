import pprint
from docx.enum.text import WD_UNDERLINE
import xmltodict
import json
from docx.document import Document

try:
    document = Document()
except TypeError:
    from docx import Document

    document = Document()


def SearchParagraph(paragraphs, str):
    count = 0
    for paragraph in paragraphs:
        count += 1
        if str in paragraph.text:
            print(count)


def __DeleteLastColumn(table_index, document):
    table = document.tables[table_index]
    grid = table._tbl.find("w:tblGrid", table._tbl.nsmap)
    for cell in table.column_cells(-1):
        cell._tc.getparent().remove(cell._tc)
    col_elem = grid[-1]
    grid.remove(col_elem)


def __DeleteRow(item):
    item._element.getparent().remove(item._element)


def __DeleteTable(table):
    table._element.getparent().remove(table._element)


def __DeleteParagraph(paragraph):
    paragraph._element.getparent().remove(paragraph._element)


def XmlToDict(fileName):
    with (open(fileName, 'r', encoding='utf16') as xml_file):
        xml_data = xmltodict.parse(xml_file.read())
    return xml_data


def GenerateDocxOch(dictInfOO: dict, doc: Document) -> Document:
    dictTimeOO = dictInfOO['Часы']

    kursRabFlagOO = False
    kursPrFlagOO = False
    konRabFlagOO = False

    listKursRabOO = list()
    listKursPrOO = list()

    zedOO = int(0)

    for key in dictTimeOO.keys():
        semestrDict = dictTimeOO[key]
        if 'Курсовая работа' in semestrDict.keys():
            kursRabFlagOO = True
            listKursRabOO.append(key)
        if 'Курсовой проект' in semestrDict.keys():
            kursPrFlagOO = True
            listKursPrOO.append(key)
        if 'Контрольная работа' in semestrDict.keys():
            konRabFlagOO = True
        if 'ЗЕТ' in semestrDict.keys():
            zedOO += int(semestrDict['ЗЕТ'])

    stringKursRabOO = str()
    if len(listKursRabOO) != 0:
        stringKursRabOO = f"{listKursRabOO[0]} семестре для очной формы обучения"
    if len(listKursRabOO) > 1:
        for index in range(1, len(listKursRabOO)):
            stringKursRabOO += f", в {listKursRabOO[index]} семестре для очной формы обучения"

    stringKursPrOO = str()
    if len(listKursPrOO) != 0:
        stringKursPrOO = f"{listKursPrOO[0]} семестре для очной формы обучения"
    if len(listKursPrOO) > 1:
        for index in range(1, len(listKursPrOO)):
            stringKursPrOO += f", в {listKursPrOO[index]} семестре для очной формы обучения"

    for object in doc.paragraphs:
        if 'name' in object.text:
            for run in object.runs:
                run.text = run.text.replace('name', dictInfOO['Название'])
        if 'spec' in object.text:
            for run in object.runs:
                run.text = run.text.replace('spec', dictInfOO['Специальность'])
        if 'prof' in object.text:
            for run in object.runs:
                run.text = run.text.replace('prof', dictInfOO['Профиль'])
        if 'qual' in object.text:
            for run in object.runs:
                run.text = run.text.replace('qual', dictInfOO['Квалификация'])
        if 'period' in object.text:
            for run in object.runs:
                run.text = run.text.replace('period', dictInfOO['srok'])
        if 'form' in object.text:
            for run in object.runs:
                run.text = run.text.replace('form', 'Очная')
        if 'startYear' in object.text:
            for run in object.runs:
                run.text = run.text.replace('startYear', dictInfOO['startYear'])
        if 'B1O' in object.text:
            for run in object.runs:
                if 'B1O' in run.text:
                    if dictInfOO['B1'][3] == 'О':
                        run.text = run.text.replace('B1O', '')
                    else:
                        run.clear()
        if 'B1B' in object.text:
            for run in object.runs:
                if 'B1B' in run.text:
                    if dictInfOO['B1'][3] != 'О':
                        run.text = run.text.replace('B1B)', '')
                    else:
                        run.clear()
        if 'zed' in object.text:
            for run in object.runs:
                run.text = run.text.replace('zed', str(zedOO))
        if 'compList' in object.text:
            compStr = str()
            for key in dictInfOO['Компетенции'].keys():
                compStr += key + ' - ' + dictInfOO['Компетенции'][key] + '\n'
            object.text = object.text.replace('compList', compStr)
        if 'kursPList' in object.text:
            for run in object.runs:
                run.text = run.text.replace('kursPList', stringKursPrOO)
        if 'kursRList' in object.text:
            for run in object.runs:
                run.text = run.text.replace('kursRList', stringKursRabOO)
        if ('KPrY' in object.text):
            for run in object.runs:
                run.text = run.text.replace('KPrY', '')
            if (kursPrFlagOO == False):
                __DeleteParagraph(object)
        if ('KRabY' in object.text):
            for run in object.runs:
                run.text = run.text.replace('KRabY', '')
            if (kursRabFlagOO == False):
                __DeleteParagraph(object)
        if ('KPY' in object.text):
            for run in object.runs:
                run.text = run.text.replace('KPY', '')
            if (kursPrFlagOO == False and kursRabFlagOO == False):
                __DeleteParagraph(object)
        if ('KPN' in object.text):
            for run in object.runs:
                run.text = run.text.replace('KPN', '')
            if (kursPrFlagOO == True or kursRabFlagOO == True):
                __DeleteParagraph(object)
        if ('KRY' in object.text):
            for run in object.runs:
                run.text = run.text.replace('KRY', '')
            if (konRabFlagOO == False):
                __DeleteParagraph(object)
        if ('KRN' in object.text):
            for run in object.runs:
                run.text = run.text.replace('KRN', '')
            if (konRabFlagOO == True):
                __DeleteParagraph(object)
        if 'timeTableZ' in object.text:
            __DeleteParagraph(object)
        if 'allTimeTableZ' in object.text:
            __DeleteParagraph(object)

    for _ in range(4 - len(dictInfOO['Компетенции'])):
        for _ in range(3):
            __DeleteRow(doc.tables[1].rows[-1])

    compiKeysOO = list()
    for key in dictInfOO['Компетенции'].keys():
        compiKeysOO.append(key)
    for cell in doc.tables[1].columns[0].cells:
        if 'comp' in cell.text:
            cell.text = cell.text.replace('comp', compiKeysOO[0])
            compiKeysOO.pop(0)

    for semestrKey in dictTimeOO.keys():
        dictTimeOO[semestrKey]['Аудиторные занятия'] = str('0')
        for timeKey in dictTimeOO[semestrKey].keys():
            if timeKey == 'Практические занятия' or timeKey == 'Лабораторные занятия' or timeKey == 'Лекционные занятия':
                dictTimeOO[semestrKey]['Аудиторные занятия'] = str(
                    int(dictTimeOO[semestrKey]['Аудиторные занятия']) + int(dictTimeOO[semestrKey][timeKey]))

    timeKeysOO = list()
    for key in dictTimeOO.keys():
        timeKeysOO.append(key)

    dictAllTimeOO = dict()
    dictAllTimeOO['Аудиторные занятия'] = str(0)
    dictAllTimeOO['allTime'] = str(0)
    for keySemestr in dictTimeOO.keys():
        for timeKey in dictTimeOO[keySemestr].keys():
            if timeKey == 'Практические занятия' or timeKey == 'Лабораторные занятия' or timeKey == 'Лекционные занятия':
                dictAllTimeOO['Аудиторные занятия'] = str(
                    int(dictAllTimeOO['Аудиторные занятия']) + int(dictTimeOO[keySemestr][timeKey]))
                try:
                    dictAllTimeOO[timeKey] = str(int(dictAllTimeOO[timeKey]) + int(dictTimeOO[keySemestr][timeKey]))
                except:
                    dictAllTimeOO[timeKey] = dictTimeOO[keySemestr][timeKey]
            if timeKey == 'Самостоятельная работа' or timeKey == 'Итого часов':
                try:
                    dictAllTimeOO[timeKey] = str(int(dictAllTimeOO[timeKey]) + int(dictTimeOO[keySemestr][timeKey]))
                except:
                    dictAllTimeOO[timeKey] = dictTimeOO[keySemestr][timeKey]
    if 'Самостоятельная работа' in dictAllTimeOO.keys():
        dictAllTimeOO['allTime'] = dictAllTimeOO['Аудиторные занятия'] + dictAllTimeOO['Самостоятельная работа']
    else:
        dictAllTimeOO['allTime'] = dictAllTimeOO['Аудиторные занятия']

    timeTableOO = doc.tables[2]
    match len(dictTimeOO):
        case 1:
            for _ in range(3 * 13):
                __DeleteRow(timeTableOO.rows[-1])
        case 2:
            for _ in range(13):
                __DeleteRow(timeTableOO.rows[0])
            for _ in range(2 * 13):
                __DeleteRow(timeTableOO.rows[-1])
        case 3:
            for _ in range(2 * 13):
                __DeleteRow(timeTableOO.rows[0])
            for _ in range(13):
                __DeleteRow(timeTableOO.rows[-1])
        case 4:
            for _ in range(3 * 13):
                __DeleteRow(timeTableOO.rows[0])

    for colum in timeTableOO.columns:
        for cell in colum.cells:
            if 'adTimeAll' in cell.text:
                if 'Аудиторные занятия' in dictAllTimeOO.keys():
                    cell.text = cell.text.replace('adTimeAll', dictAllTimeOO['Аудиторные занятия'])
                else:
                    cell.text = ''
            if 'lcTimeAll' in cell.text:
                if 'Лекционные занятия' in dictAllTimeOO.keys():
                    cell.text = cell.text.replace('lcTimeAll', dictAllTimeOO['Лекционные занятия'])
                else:
                    cell.text = ''
            if 'prctTimeAll' in cell.text:
                if 'Практические занятия' in dictAllTimeOO.keys():
                    cell.text = cell.text.replace('prctTimeAll', dictAllTimeOO['Практические занятия'])
                else:
                    cell.text = ''
            if 'lbTimeAll' in cell.text:
                if 'Лабораторные занятия' in dictAllTimeOO.keys():
                    cell.text = cell.text.replace('lbTimeAll', dictAllTimeOO['Лабораторные занятия'])
                else:
                    cell.text = ''
            if 'smTimeAll' in cell.text:
                if 'Самостоятельная работа' in dictAllTimeOO.keys():
                    cell.text = cell.text.replace('smTimeAll', dictAllTimeOO['Самостоятельная работа'])
                else:
                    cell.text = ''
            if 'allTime' in cell.text:
                if 'Итого часов' in dictAllTimeOO.keys():
                    cell.text = cell.text.replace('allTime', dictAllTimeOO['Итого часов'])
                else:
                    cell.text = ''
            if 'allZed' in cell.text:
                cell.text = str(zedOO)
            if 'semestr' in cell.text:
                cell.text = cell.text.replace('semestr', str(timeKeysOO[0]))
            if 'audTime' in cell.text:
                if 'Аудиторные занятия' in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = cell.text.replace('audTime', dictTimeOO[timeKeysOO[0]]['Аудиторные занятия'])
                else:
                    cell.text = ''
            if 'practTime' in cell.text:
                if 'Практические занятия' in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = cell.text.replace('practTime', dictTimeOO[timeKeysOO[0]]['Практические занятия'])
                else:
                    cell.text = ''
            if 'labTime' in cell.text:
                if 'Лабораторные занятия' in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = cell.text.replace('labTime', dictTimeOO[timeKeysOO[0]]['Лабораторные занятия'])
                else:
                    cell.text = ''
            if 'lectTime' in cell.text:
                if 'Лекционные занятия' in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = cell.text.replace('lectTime', dictTimeOO[timeKeysOO[0]]['Лекционные занятия'])
                else:
                    cell.text = ''
            if 'samTime' in cell.text:
                if 'Самостоятельная работа' in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = cell.text.replace('samTime', dictTimeOO[timeKeysOO[0]]['Самостоятельная работа'])
                else:
                    cell.text = ''
            if 'kurs' in cell.text:
                if 'Курсовой проект' in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = '+'
                else:
                    cell.text = '-'
            if 'kr' in cell.text:
                if 'Контрольная работа' in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = '+'
                else:
                    cell.text = '-'
            if 'att' in cell.text:
                if 'Зачет' in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = cell.text.replace('att', 'Зачет')
                elif 'Зачет с оценкой' in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = cell.text.replace('att', 'Зачет с оценкой')
                elif 'Экзамен' in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = cell.text.replace('att', 'Экзамен')
            if 'fullTime' in cell.text:
                if 'Итого часов' in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = cell.text.replace('fullTime', dictTimeOO[timeKeysOO[0]]['Итого часов'])
                else:
                    cell.text = ''
            if 'fullZed' in cell.text:
                if 'ЗЕТ' in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = cell.text.replace('fullZed', dictTimeOO[timeKeysOO[0]]['ЗЕТ'])
                else:
                    cell.text = ''
                timeKeysOO.pop(0)

    allTimeTableOO = doc.tables[4]
    for colum in allTimeTableOO.columns:
        for cell in colum.cells:
            if 'allLect' in cell.text:
                if 'Лекционные занятия' in dictAllTimeOO.keys():
                    cell.text = cell.text.replace('allLect', dictAllTimeOO['Лекционные занятия'])
                else:
                    cell.text = ''
            if 'allPract' in cell.text:
                if 'Практические занятия' in dictAllTimeOO.keys():
                    cell.text = cell.text.replace('allPract', dictAllTimeOO['Практические занятия'])
                else:
                    cell.text = ''
            if 'allLab' in cell.text:
                if 'Лабораторные занятия' in dictAllTimeOO.keys():
                    cell.text = cell.text.replace('allLab', dictAllTimeOO['Лабораторные занятия'])
                else:
                    cell.text = ''
            if 'allSam' in cell.text:
                if 'Самостоятельная работа' in dictAllTimeOO.keys():
                    cell.text = cell.text.replace('allSam', dictAllTimeOO['Самостоятельная работа'])
                else:
                    cell.text = ''
            if 'allTime' in cell.text:
                if 'Аудиторные занятия' in dictAllTimeOO.keys():
                    if 'Самостоятельная работа' in dictAllTimeOO.keys():
                        cell.text = cell.text.replace('allTime', str(int(dictAllTimeOO['Аудиторные занятия']) + int(
                            dictAllTimeOO['Самостоятельная работа'])))
                    else:
                        cell.text = ''

    for _ in range(4 - len(dictInfOO['Компетенции'])):
        for _ in range(3):
            __DeleteRow(doc.tables[7].rows[-1])

    compiKeysOO = list()
    for key in dictInfOO['Компетенции'].keys():
        compiKeysOO.append(key)
    for cell in doc.tables[7].columns[0].cells:
        if 'comp' in cell.text:
            cell.text = cell.text.replace('comp', compiKeysOO[0])
            compiKeysOO.pop(0)

    for _ in range(4 - len(dictInfOO['Компетенции'])):
        for _ in range(3):
            __DeleteRow(doc.tables[8].rows[-1])

    compiKeysOO = list()
    for key in dictInfOO['Компетенции'].keys():
        compiKeysOO.append(key)
    for cell in doc.tables[8].columns[0].cells:
        if 'comp' in cell.text:
            cell.text = cell.text.replace('comp', compiKeysOO[0])
            compiKeysOO.pop(0)

    __DeleteTable(doc.tables[5])
    __DeleteTable(doc.tables[3])

    return doc


def GenerateDocxOchZ(dictInfOO: dict, dictInfZO: dict, doc: Document) -> Document:
    dictTimeOO = dictInfOO['Часы']
    dictTimeZO = dictInfZO['Часы']

    kursRabFlagOO = False
    kursPrFlagOO = False
    konRabFlagOO = False

    kursRabFlagZO = False
    kursPrFlagZO = False
    konRabFlagZO = False

    listKursRabOO = list()
    listKursRabZO = list()
    listKursPrOO = list()
    listKursPrZO = list()

    zedOO = int(0)
    zedZO = int(0)

    for key in dictTimeOO.keys():
        semestrDict = dictTimeOO[key]
        if 'Курсовая работа' in semestrDict.keys():
            kursRabFlagOO = True
            listKursRabOO.append(key)
        if 'Курсовой проект' in semestrDict.keys():
            kursPrFlagOO = True
            listKursPrOO.append(key)
        if 'Контрольная работа' in semestrDict.keys():
            konRabFlagOO = True
        if 'ЗЕТ' in semestrDict.keys():
            zedOO += int(semestrDict['ЗЕТ'])

    for key in dictTimeZO.keys():
        semestrDict = dictTimeZO[key]
        if 'Курсовая работа' in semestrDict.keys():
            kursRabFlagZO = True
            listKursRabOO.append(key)
        if 'Курсовой проект' in semestrDict.keys():
            kursPrFlagZO = True
            listKursPrZO.append(key)
        if 'Контрольная работа' in semestrDict.keys():
            konRabFlagZO = True
        if 'ЗЕТ' in semestrDict.keys():
            zedZO += int(semestrDict['ЗЕТ'])

    stringKursRabOO = str()
    if len(listKursRabOO) != 0:
        stringKursRabOO = f"{listKursRabOO[0]} семестре для очной формы обучения"
    if len(listKursRabOO) > 1:
        for index in range(1, len(listKursRabOO)):
            stringKursRabOO += f", в {listKursRabOO[index]} семестре для очной формы обучения"

    stringKursRabZO = str()
    if len(listKursRabZO) != 0:
        if stringKursRabZO != '':
            for index in range(0, len(listKursRabZO)):
                stringKursRabZO += f", в {listKursRabZO[index]} семестре для заочной формы обучения"
        else:
            stringKursRabZO = f"{listKursRabZO[0]} семестре для заочной формы обучения"
            if len(listKursRabZO) > 1:
                for index in range(1, len(listKursRabZO)):
                    stringKursRabZO += f", в {listKursRabZO[index]} семестре для заочной формы обучения"

    stringKursRab = stringKursRabOO + stringKursRabZO

    stringKursPrOO = str()
    if len(listKursPrOO) != 0:
        stringKursPrOO = f"{listKursPrOO[0]} семестре для очной формы обучения"
    if len(listKursPrOO) > 1:
        for index in range(1, len(listKursPrOO)):
            stringKursPrOO += f", в {listKursPrOO[index]} семестре для очной формы обучения"
    stringKursPrZO = str()
    if len(listKursPrZO) != 0:
        if stringKursPrOO != '':
            for index in range(0, len(listKursPrZO)):
                stringKursPrZO += f", в {listKursPrZO[index]} семестре для заочной формы обучения"
        else:
            stringKursPrZO = f"{listKursPrZO[0]} семестре для заочной формы обучения"
            if len(listKursPrZO) > 1:
                for index in range(1, len(listKursPrZO)):
                    stringKursPrZO += f", в {listKursPrZO[index]} семестре для заочной формы обучения"

    stringKursPr = stringKursPrOO + stringKursPrZO

    for object in doc.paragraphs:
        if 'name' in object.text:
            for run in object.runs:
                run.text = run.text.replace('name', dictInfOO['Название'])
        if 'spec' in object.text:
            for run in object.runs:
                run.text = run.text.replace('spec', dictInfOO['Специальность'])
        if 'prof' in object.text:
            for run in object.runs:
                run.text = run.text.replace('prof', dictInfOO['Профиль'])
        if 'qual' in object.text:
            for run in object.runs:
                run.text = run.text.replace('qual', dictInfOO['Квалификация'])
        if 'period' in object.text:
            for run in object.runs:
                run.text = run.text.replace('period', f"{dictInfOO['srok']} / {dictInfZO['srok']}")
        if 'form' in object.text:
            for run in object.runs:
                run.text = run.text.replace('form', 'Очная/заочная')
        if 'startYear' in object.text:
            for run in object.runs:
                run.text = run.text.replace('startYear', dictInfOO['startYear'])
        if 'B1O' in object.text:
            for run in object.runs:
                if 'B1O' in run.text:
                    if dictInfOO['B1'][3] == 'О':
                        run.text = run.text.replace('B1O', '')
                    else:
                        run.clear()
        if 'B1B' in object.text:
            for run in object.runs:
                if 'B1B' in run.text:
                    if dictInfOO['B1'][3] != 'О':
                        run.text = run.text.replace('B1B)', '')
                    else:
                        run.clear()
        if 'zed' in object.text:
            for run in object.runs:
                run.text = run.text.replace('zed', f"{zedOO} / {zedZO}")
        if 'compList' in object.text:
            compStr = str()
            for key in dictInfOO['Компетенции'].keys():
                compStr += key + ' - ' + dictInfOO['Компетенции'][key] + '\n'
            object.text = object.text.replace('compList', compStr)
        if 'kursPList' in object.text:
            for run in object.runs:
                run.text = run.text.replace('kursPList', stringKursPr)
        if 'kursRList' in object.text:
            for run in object.runs:
                run.text = run.text.replace('kursRList', stringKursRab)
        if ('KPrY' in object.text):
            for run in object.runs:
                run.text = run.text.replace('KPrY', '')
            if (kursPrFlagOO == False and kursPrFlagZO == False):
                __DeleteParagraph(object)
        if ('KRabY' in object.text):
            for run in object.runs:
                run.text = run.text.replace('KRabY', '')
            if (kursRabFlagZO == False and kursRabFlagOO == False):
                __DeleteParagraph(object)
        if ('KPY' in object.text):
            for run in object.runs:
                run.text = run.text.replace('KPY', '')
            if (kursPrFlagOO == False and kursPrFlagZO == False and kursRabFlagOO == False and kursRabFlagZO == False):
                __DeleteParagraph(object)
        if ('KPN' in object.text):
            for run in object.runs:
                run.text = run.text.replace('KPN', '')
            if (kursPrFlagOO == True or kursPrFlagZO == True or kursRabFlagOO == True):
                __DeleteParagraph(object)
        if ('KRY' in object.text):
            for run in object.runs:
                run.text = run.text.replace('KRY', '')
            if (konRabFlagOO == False and konRabFlagZO == False):
                __DeleteParagraph(object)
        if ('KRN' in object.text):
            for run in object.runs:
                run.text = run.text.replace('KRN', '')
            if (konRabFlagOO == True or konRabFlagZO == True):
                __DeleteParagraph(object)

        if 'timeTableZ' in object.text:
            object.runs[0].text = object.runs[0].text.replace('timeTableZ', '')
        if 'allTimeTableZ' in object.text:
            object.runs[0].text = object.runs[0].text.replace('allTimeTableZ', '')

    for _ in range(4 - len(dictInfOO['Компетенции'])):
        for _ in range(3):
            __DeleteRow(doc.tables[1].rows[-1])

    compiKeysOO = list()
    for key in dictInfOO['Компетенции'].keys():
        compiKeysOO.append(key)
    for cell in doc.tables[1].columns[0].cells:
        if 'comp' in cell.text:
            cell.text = cell.text.replace('comp', compiKeysOO[0])
            compiKeysOO.pop(0)

    for semestrKey in dictTimeOO.keys():
        dictTimeOO[semestrKey]['Аудиторные занятия'] = str('0')
        for timeKey in dictTimeOO[semestrKey].keys():
            if timeKey == 'Практические занятия' or timeKey == 'Лабораторные занятия' or timeKey == 'Лекционные занятия':
                dictTimeOO[semestrKey]['Аудиторные занятия'] = str(
                    int(dictTimeOO[semestrKey]['Аудиторные занятия']) + int(dictTimeOO[semestrKey][timeKey]))

    timeKeysOO = list()
    for key in dictTimeOO.keys():
        timeKeysOO.append(key)

    dictAllTimeOO = dict()
    dictAllTimeOO['Аудиторные занятия'] = str(0)
    dictAllTimeOO['allTime'] = str(0)
    for keySemestr in dictTimeOO.keys():
        for timeKey in dictTimeOO[keySemestr].keys():
            if timeKey == 'Практические занятия' or timeKey == 'Лабораторные занятия' or timeKey == 'Лекционные занятия':
                dictAllTimeOO['Аудиторные занятия'] = str(
                    int(dictAllTimeOO['Аудиторные занятия']) + int(dictTimeOO[keySemestr][timeKey]))
                try:
                    dictAllTimeOO[timeKey] = str(int(dictAllTimeOO[timeKey]) + int(dictTimeOO[keySemestr][timeKey]))
                except:
                    dictAllTimeOO[timeKey] = dictTimeOO[keySemestr][timeKey]
            if timeKey == 'Самостоятельная работа' or timeKey == 'Итого часов':
                try:
                    dictAllTimeOO[timeKey] = str(int(dictAllTimeOO[timeKey]) + int(dictTimeOO[keySemestr][timeKey]))
                except:
                    dictAllTimeOO[timeKey] = dictTimeOO[keySemestr][timeKey]
    if 'Самостоятельная работа' in dictAllTimeOO.keys():
        dictAllTimeOO['allTime'] = dictAllTimeOO['Аудиторные занятия'] + dictAllTimeOO['Самостоятельная работа']
    else:
        dictAllTimeOO['allTime'] = dictAllTimeOO['Аудиторные занятия']

    timeTableOO = doc.tables[2]
    match len(dictTimeOO):
        case 1:
            for _ in range(3 * 13):
                __DeleteRow(timeTableOO.rows[-1])
        case 2:
            for _ in range(13):
                __DeleteRow(timeTableOO.rows[0])
            for _ in range(2 * 13):
                __DeleteRow(timeTableOO.rows[-1])
        case 3:
            for _ in range(2 * 13):
                __DeleteRow(timeTableOO.rows[0])
            for _ in range(13):
                __DeleteRow(timeTableOO.rows[-1])
        case 4:
            for _ in range(3 * 13):
                __DeleteRow(timeTableOO.rows[0])

    for colum in timeTableOO.columns:
        for cell in colum.cells:
            if 'adTimeAll' in cell.text:
                if 'Аудиторные занятия' in dictAllTimeOO.keys():
                    cell.text = cell.text.replace('adTimeAll', dictAllTimeOO['Аудиторные занятия'])
                else:
                    cell.text = ''
            if 'lcTimeAll' in cell.text:
                if 'Лекционные занятия' in dictAllTimeOO.keys():
                    cell.text = cell.text.replace('lcTimeAll', dictAllTimeOO['Лекционные занятия'])
                else:
                    cell.text = ''
            if 'prctTimeAll' in cell.text:
                if 'Практические занятия' in dictAllTimeOO.keys():
                    cell.text = cell.text.replace('prctTimeAll', dictAllTimeOO['Практические занятия'])
                else:
                    cell.text = ''
            if 'lbTimeAll' in cell.text:
                if 'Лабораторные занятия' in dictAllTimeOO.keys():
                    cell.text = cell.text.replace('lbTimeAll', dictAllTimeOO['Лабораторные занятия'])
                else:
                    cell.text = ''
            if 'smTimeAll' in cell.text:
                if 'Самостоятельная работа' in dictAllTimeOO.keys():
                    cell.text = cell.text.replace('smTimeAll', dictAllTimeOO['Самостоятельная работа'])
                else:
                    cell.text = ''
            if 'allTime' in cell.text:
                if 'Итого часов' in dictAllTimeOO.keys():
                    cell.text = cell.text.replace('allTime', dictAllTimeOO['Итого часов'])
                else:
                    cell.text = ''
            if 'allZed' in cell.text:
                cell.text = str(zedOO)
            if 'semestr' in cell.text:
                cell.text = cell.text.replace('semestr', str(timeKeysOO[0]))
            if 'audTime' in cell.text:
                if 'Аудиторные занятия' in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = cell.text.replace('audTime', dictTimeOO[timeKeysOO[0]]['Аудиторные занятия'])
                else:
                    cell.text = ''
            if 'practTime' in cell.text:
                if 'Практические занятия' in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = cell.text.replace('practTime', dictTimeOO[timeKeysOO[0]]['Практические занятия'])
                else:
                    cell.text = ''
            if 'labTime' in cell.text:
                if 'Лабораторные занятия' in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = cell.text.replace('labTime', dictTimeOO[timeKeysOO[0]]['Лабораторные занятия'])
                else:
                    cell.text = ''
            if 'lectTime' in cell.text:
                if 'Лекционные занятия' in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = cell.text.replace('lectTime', dictTimeOO[timeKeysOO[0]]['Лекционные занятия'])
                else:
                    cell.text = ''
            if 'samTime' in cell.text:
                if 'Самостоятельная работа' in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = cell.text.replace('samTime', dictTimeOO[timeKeysOO[0]]['Самостоятельная работа'])
                else:
                    cell.text = ''
            if 'kurs' in cell.text:
                if ('Курсовой проект' or 'Курсовая работа') in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = '+'
                else:
                    cell.text = '-'
            if 'kr' in cell.text:
                if 'Контрольная работа' in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = '+'
                else:
                    cell.text = '-'
            if 'att' in cell.text:
                if 'Зачет' in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = cell.text.replace('att', 'Зачет')
                elif 'Зачет с оценкой' in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = cell.text.replace('att', 'Зачет с оценкой')
                elif 'Экзамен' in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = cell.text.replace('att', 'Экзамен')
            if 'fullTime' in cell.text:
                if 'Итого часов' in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = cell.text.replace('fullTime', dictTimeOO[timeKeysOO[0]]['Итого часов'])
                else:
                    cell.text = ''
            if 'fullZed' in cell.text:
                if 'ЗЕТ' in dictTimeOO[timeKeysOO[0]].keys():
                    cell.text = cell.text.replace('fullZed', dictTimeOO[timeKeysOO[0]]['ЗЕТ'])
                else:
                    cell.text = ''
                timeKeysOO.pop(0)

    allTimeTableOO = doc.tables[4]
    for colum in allTimeTableOO.columns:
        for cell in colum.cells:
            if 'allLect' in cell.text:
                if 'Лекционные занятия' in dictAllTimeOO.keys():
                    cell.text = cell.text.replace('allLect', dictAllTimeOO['Лекционные занятия'])
                else:
                    cell.text = ''
            if 'allPract' in cell.text:
                if 'Практические занятия' in dictAllTimeOO.keys():
                    cell.text = cell.text.replace('allPract', dictAllTimeOO['Практические занятия'])
                else:
                    cell.text = ''
            if 'allLab' in cell.text:
                if 'Лабораторные занятия' in dictAllTimeOO.keys():
                    cell.text = cell.text.replace('allLab', dictAllTimeOO['Лабораторные занятия'])
                else:
                    cell.text = ''
            if 'allSam' in cell.text:
                if 'Самостоятельная работа' in dictAllTimeOO.keys():
                    cell.text = cell.text.replace('allSam', dictAllTimeOO['Самостоятельная работа'])
                else:
                    cell.text = ''
            if 'allTime' in cell.text:
                if 'Аудиторные занятия' in dictAllTimeOO.keys():
                    if 'Самостоятельная работа' in dictAllTimeOO.keys():
                        cell.text = cell.text.replace('allTime', str(int(dictAllTimeOO['Аудиторные занятия']) + int(
                            dictAllTimeOO['Самостоятельная работа'])))
                    else:
                        cell.text = ''

    for semestrKey in dictTimeZO.keys():
        dictTimeZO[semestrKey]['Аудиторные занятия'] = str('0')
        for timeKey in dictTimeZO[semestrKey].keys():
            if timeKey == 'Практические занятия' or timeKey == 'Лабораторные занятия' or timeKey == 'Лекционные занятия':
                dictTimeZO[semestrKey]['Аудиторные занятия'] = str(
                    int(dictTimeZO[semestrKey]['Аудиторные занятия']) + int(dictTimeZO[semestrKey][timeKey]))

    timeKeysZO = list()
    for key in dictTimeZO.keys():
        timeKeysZO.append(key)

    dictAllTimeZO = dict()
    dictAllTimeZO['Аудиторные занятия'] = str(0)
    dictAllTimeZO['allTime'] = str(0)
    for keySemestr in dictTimeZO.keys():
        for timeKey in dictTimeZO[keySemestr].keys():
            if timeKey == 'Практические занятия' or timeKey == 'Лабораторные занятия' or timeKey == 'Лекционные занятия':
                dictAllTimeZO['Аудиторные занятия'] = str(
                    int(dictAllTimeZO['Аудиторные занятия']) + int(dictTimeZO[keySemestr][timeKey]))
                try:
                    dictAllTimeZO[timeKey] = str(int(dictAllTimeZO[timeKey]) + int(dictTimeZO[keySemestr][timeKey]))
                except:
                    dictAllTimeZO[timeKey] = dictTimeZO[keySemestr][timeKey]
            if timeKey == 'Самостоятельная работа' or timeKey == 'Итого часов':
                try:
                    dictAllTimeZO[timeKey] = str(int(dictAllTimeZO[timeKey]) + int(dictTimeZO[keySemestr][timeKey]))
                except:
                    dictAllTimeZO[timeKey] = dictTimeZO[keySemestr][timeKey]
    if 'Самостоятельная работа' in dictAllTimeZO.keys():
        dictAllTimeZO['allTime'] = dictAllTimeZO['Аудиторные занятия'] + dictAllTimeZO['Самостоятельная работа']
    else:
        dictAllTimeZO['allTime'] = dictAllTimeZO['Аудиторные занятия']

    timeTableZO = doc.tables[3]
    match len(dictTimeZO):
        case 1:
            for _ in range(3 * 13):
                __DeleteRow(timeTableZO.rows[-1])
        case 2:
            for _ in range(13):
                __DeleteRow(timeTableZO.rows[0])
            for _ in range(2 * 13):
                __DeleteRow(timeTableZO.rows[-1])
        case 3:
            for _ in range(2 * 13):
                __DeleteRow(timeTableZO.rows[0])
            for _ in range(13):
                __DeleteRow(timeTableZO.rows[-1])
        case 4:
            for _ in range(3 * 13):
                __DeleteRow(timeTableZO.rows[0])

    for colum in timeTableZO.columns:
        for cell in colum.cells:
            if 'adTimeAll' in cell.text:
                if 'Аудиторные занятия' in dictAllTimeZO.keys():
                    cell.text = cell.text.replace('adTimeAll', dictAllTimeZO['Аудиторные занятия'])
                else:
                    cell.text = ''
            if 'lcTimeAll' in cell.text:
                if 'Лекционные занятия' in dictAllTimeZO.keys():
                    cell.text = cell.text.replace('lcTimeAll', dictAllTimeZO['Лекционные занятия'])
                else:
                    cell.text = ''
            if 'prctTimeAll' in cell.text:
                if 'Практические занятия' in dictAllTimeZO.keys():
                    cell.text = cell.text.replace('prctTimeAll', dictAllTimeZO['Практические занятия'])
                else:
                    cell.text = ''
            if 'lbTimeAll' in cell.text:
                if 'Лабораторные занятия' in dictAllTimeZO.keys():
                    cell.text = cell.text.replace('lbTimeAll', dictAllTimeZO['Лабораторные занятия'])
                else:
                    cell.text = ''
            if 'smTimeAll' in cell.text:
                if 'Самостоятельная работа' in dictAllTimeZO.keys():
                    cell.text = cell.text.replace('smTimeAll', dictAllTimeZO['Самостоятельная работа'])
                else:
                    cell.text = ''
            if 'allTime' in cell.text:
                if 'Итого часов' in dictAllTimeZO.keys():
                    cell.text = cell.text.replace('allTime', dictAllTimeZO['Итого часов'])
                else:
                    cell.text = ''
            if 'allZed' in cell.text:
                cell.text = str(zedZO)
            if 'semestr' in cell.text:
                cell.text = cell.text.replace('semestr', str(timeKeysZO[0]))
            if 'audTime' in cell.text:
                if 'Аудиторные занятия' in dictTimeZO[timeKeysZO[0]].keys():
                    cell.text = cell.text.replace('audTime', dictTimeZO[timeKeysZO[0]]['Аудиторные занятия'])
                else:
                    cell.text = ''
            if 'practTime' in cell.text:
                if 'Практические занятия' in dictTimeZO[timeKeysZO[0]].keys():
                    cell.text = cell.text.replace('practTime', dictTimeZO[timeKeysZO[0]]['Практические занятия'])
                else:
                    cell.text = ''
            if 'labTime' in cell.text:
                if 'Лабораторные занятия' in dictTimeZO[timeKeysZO[0]].keys():
                    cell.text = cell.text.replace('labTime', dictTimeZO[timeKeysZO[0]]['Лабораторные занятия'])
                else:
                    cell.text = ''
            if 'lectTime' in cell.text:
                if 'Лекционные занятия' in dictTimeZO[timeKeysZO[0]].keys():
                    cell.text = cell.text.replace('lectTime', dictTimeZO[timeKeysZO[0]]['Лекционные занятия'])
                else:
                    cell.text = ''
            if 'samTime' in cell.text:
                if 'Самостоятельная работа' in dictTimeZO[timeKeysZO[0]].keys():
                    cell.text = cell.text.replace('samTime', dictTimeZO[timeKeysZO[0]]['Самостоятельная работа'])
                else:
                    cell.text = ''
            if 'kurs' in cell.text:
                if ('Курсовой проект' or 'Курсовая работа') in dictTimeZO[timeKeysZO[0]].keys():
                    cell.text = '+'
                else:
                    cell.text = '-'
            if 'kr' in cell.text:
                if 'Контрольная работа' in dictTimeZO[timeKeysZO[0]].keys():
                    cell.text = '+'
                else:
                    cell.text = '-'
            if 'att' in cell.text:
                if 'Зачет' in dictTimeZO[timeKeysZO[0]].keys():
                    cell.text = cell.text.replace('att', 'Зачет')
                elif 'Зачет с оценкой' in dictTimeZO[timeKeysZO[0]].keys():
                    cell.text = cell.text.replace('att', 'Зачет с оценкой')
                elif 'Экзамен' in dictTimeZO[timeKeysZO[0]].keys():
                    cell.text = cell.text.replace('att', 'Экзамен')
            if 'fullTime' in cell.text:
                if 'Итого часов' in dictTimeZO[timeKeysZO[0]].keys():
                    cell.text = cell.text.replace('fullTime', dictTimeZO[timeKeysZO[0]]['Итого часов'])
                else:
                    cell.text = ''
            if 'fullZed' in cell.text:
                if 'ЗЕТ' in dictTimeZO[timeKeysZO[0]].keys():
                    cell.text = cell.text.replace('fullZed', dictTimeZO[timeKeysZO[0]]['ЗЕТ'])
                else:
                    cell.text = ''
                timeKeysZO.pop(0)

    allTimeTableZO = doc.tables[5]
    for colum in allTimeTableZO.columns:
        for cell in colum.cells:
            if 'allLect' in cell.text:
                if 'Лекционные занятия' in dictAllTimeZO.keys():
                    cell.text = cell.text.replace('allLect', dictAllTimeZO['Лекционные занятия'])
                else:
                    cell.text = ''
            if 'allPract' in cell.text:
                if 'Практические занятия' in dictAllTimeZO.keys():
                    cell.text = cell.text.replace('allPract', dictAllTimeZO['Практические занятия'])
                else:
                    cell.text = ''
            if 'allLab' in cell.text:
                if 'Лабораторные занятия' in dictAllTimeZO.keys():
                    cell.text = cell.text.replace('allLab', dictAllTimeZO['Лабораторные занятия'])
                else:
                    cell.text = ''
            if 'allSam' in cell.text:
                if 'Самостоятельная работа' in dictAllTimeZO.keys():
                    cell.text = cell.text.replace('allSam', dictAllTimeZO['Самостоятельная работа'])
                else:
                    cell.text = ''
            if 'allTime' in cell.text:
                if 'Аудиторные занятия' in dictAllTimeZO.keys():
                    if 'Самостоятельная работа' in dictAllTimeZO.keys():
                        cell.text = cell.text.replace('allTime', str(int(dictAllTimeZO['Аудиторные занятия']) + int(
                            dictAllTimeZO['Самостоятельная работа'])))
                    else:
                        cell.text = ''

    for _ in range(4 - len(dictInfOO['Компетенции'])):
        for _ in range(3):
            __DeleteRow(doc.tables[7].rows[-1])

    compiKeysOO = list()
    for key in dictInfOO['Компетенции'].keys():
        compiKeysOO.append(key)
    for cell in doc.tables[7].columns[0].cells:
        if 'comp' in cell.text:
            cell.text = cell.text.replace('comp', compiKeysOO[0])
            compiKeysOO.pop(0)

    for _ in range(4 - len(dictInfOO['Компетенции'])):
        for _ in range(3):
            __DeleteRow(doc.tables[8].rows[-1])

    compiKeysOO = list()
    for key in dictInfOO['Компетенции'].keys():
        compiKeysOO.append(key)
    for cell in doc.tables[8].columns[0].cells:
        if 'comp' in cell.text:
            cell.text = cell.text.replace('comp', compiKeysOO[0])
            compiKeysOO.pop(0)

    return doc


def GetDisciplineList(plxData: dict) -> dict:
    list = {}
    for object in plxData['Документ']['diffgr:diffgram']['dsMMISDB']['ПланыСтроки']:
        if '@КодКафедры' in object.keys() and object['@КодКафедры'] == '82':
            list[object['@Код']] = object['@Дисциплина']
    return list


def __SearchCompetenciesByDisciplineCode(disciplineCode: str, plxData: dict) -> dict:
    compCodeList = []
    for object in plxData['Документ']['diffgr:diffgram']['dsMMISDB']['ПланыКомпетенцииДисциплины']:
        if object['@КодСтроки'] == disciplineCode:
            compCodeList.append(object['@КодКомпетенции'])
    dict = {}
    for object in plxData['Документ']['diffgr:diffgram']['dsMMISDB']['ПланыКомпетенции']:
        if compCodeList.__contains__(object['@Код']):
            dict[object['@ШифрКомпетенции']] = object['@Наименование']
    return dict


def __SearchHoursBySemesterNumber(semesterNumber: int, disciplineCode: str, plxData: dict) -> dict:
    codeList = []
    hoursList = []
    for object in plxData['Документ']['diffgr:diffgram']['dsMMISDB']['ПланыНовыеЧасы']:
        if (object['@КодОбъекта'] == disciplineCode) and (
                int(object['@Курс']) * 2 - 1 + int(object['@Семестр']) - 1 == semesterNumber or int(
            object['@Курс']) * 2 - 1 + ((int(object['@Сессия']) - 1) // 2) == semesterNumber):
            if codeList.__contains__(object['@КодВидаРаботы']) == False:
                codeList.append(object['@КодВидаРаботы'])
                hoursList.append(object['@Количество'])
    nameList = []
    dict = {}
    for key in codeList:
        for object in plxData['Документ']['diffgr:diffgram']['dsMMISDB']['СправочникВидыРабот']:
            if object['@Код'] == key:
                nameList.append(object['@Название'])
    for i in range(nameList.__len__()):
        dict[nameList[i]] = hoursList[i]

    return dict


def __SearchHours(disciplineCode: str, plxData: dict) -> dict:
    dict = {}
    semesterNumberList = []
    for object in plxData['Документ']['diffgr:diffgram']['dsMMISDB']['ПланыНовыеЧасы']:

        if object['@КодОбъекта'] == disciplineCode:
            if object['@Семестр'] != '0':
                num = int(object['@Курс']) * 2 - 1 + int(object['@Семестр']) - 1
            else:
                num = int(object['@Курс']) * 2 - 1 + ((int(object['@Сессия']) - 1) // 2)
            if semesterNumberList.__contains__(num) == False:
                semesterNumberList.append(num)
    for i in range(semesterNumberList.__len__()):
        dict[semesterNumberList[i]] = __SearchHoursBySemesterNumber(semesterNumberList[i], disciplineCode, plxData)
    return dict


def KeyFromVal(dict: dict, val):
    for key, value in dict.items():
        if value == val:
            return key


def ReadDocxTemplate(filePath: str) -> Document:
    doc = Document(filePath)
    return doc


def SaveDocx(doc: Document, fileName: str, path: str):
    fullPath = str(path) + str(fileName) + '.docx'
    doc.save(fullPath)


def __GetSpecialization(plxData: dict) -> str:
    str = plxData['Документ']['diffgr:diffgram']['dsMMISDB']['ООП']['@Название'] + ' / ' + \
          plxData['Документ']['diffgr:diffgram']['dsMMISDB']['ООП']['@Шифр']
    return str


def __GetB1(disciplineCode: str, plxData: dict) -> str:
    string = str()
    for object in plxData['Документ']['diffgr:diffgram']['dsMMISDB']['ПланыСтроки']:
        if object['@Код'] == disciplineCode:
            string = object['@ДисциплинаКод']
    return string


def __GetProfile(plxData: dict) -> str:
    string = plxData['Документ']['diffgr:diffgram']['dsMMISDB']['ООП']['ООП']['@Название']
    return string


def __GetQualification(plxData: dict) -> str:
    code = plxData['Документ']['diffgr:diffgram']['dsMMISDB']['ООП']['@Квалификация']
    for object in plxData['Документ']['diffgr:diffgram']['dsMMISDB']['Уровень_образования']:
        if object['@Код_записи'] == code:
            return object['@ВидПлана']


def __GetStartYear(plxData: dict) -> str:
    string = plxData['Документ']['diffgr:diffgram']['dsMMISDB']['Планы']['@ГодНачалаПодготовки']
    return string


def __GetSrok(plxData: dict) -> str:
    string = plxData['Документ']['diffgr:diffgram']['dsMMISDB']['Планы']['@СрокОбучения'] + ' года '
    if plxData['Документ']['diffgr:diffgram']['dsMMISDB']['Планы']['@СрокОбученияМесяцев'] != '0':
        string += plxData['Документ']['diffgr:diffgram']['dsMMISDB']['Планы']['@СрокОбученияМесяцев'] + ' месяцев'
    return string


def GetFullInf(disciplineName: str, disciplineCode: str, plxData: dict) -> dict:
    dictInf = {}
    dictInf['Название'] = disciplineName
    dictInf['Специальность'] = __GetSpecialization(plxData)
    dictInf['Профиль'] = __GetProfile(plxData)
    dictInf['Квалификация'] = __GetQualification(plxData)
    dictInf['Компетенции'] = __SearchCompetenciesByDisciplineCode(disciplineCode, plxData)
    dictInf['Часы'] = __SearchHours(disciplineCode, plxData)
    dictInf['B1'] = __GetB1(disciplineCode, plxData)
    dictInf['startYear'] = __GetStartYear(plxData)
    dictInf['srok'] = __GetSrok(plxData)

    return dictInf

# _________________________________________________________________________________________________________________________________________Test code

# doc = ReadDocxTemplate('./examples/RPD.docx')
# fileDataOO = XmlToDict('./data/ochnoe.plx')
# discListOO = GetDisciplineList(fileDataOO)
# dictInfOO = GetFullInf('Информатика', KeyFromVal(discListOO, 'Информатика'), fileDataOO)
#
# fileDataZO = XmlToDict('./data/zaoch.plx')
# discListZO = GetDisciplineList(fileDataZO)
# dictInfZO = GetFullInf('Информатика', KeyFromVal(discListZO, 'Информатика'), fileDataZO)
#
# doc = GenerateDocxOchZ(dictInfOO, dictInfZO, doc)
# SaveDocx(doc, 'test_ZO', './files/')
#
# doc = ReadDocxTemplate('./examples/RPD.docx')
# fileData = XmlToDict('./data/ochnoe.plx')
# discList = GetDisciplineList(fileData)
# dictInf = GetFullInf('Информатика', KeyFromVal(discList, 'Информатика'), fileData)
# doc = GenerateDocxOch(dictInf, doc)
# SaveDocx(doc, 'test_O', './files/')

# Базы данных
# Информатика
