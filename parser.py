import pprint
import xml.etree.ElementTree as ET
# from docx import Document
# from docx.document import Document
from docx.enum.text import WD_UNDERLINE
# from docx.shared import Inches
import json
import xmltodict

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

def GenerateDocx(dictInf: dict, doc: Document):
    doc.tables[0].cell(0, 1).paragraphs[1].runs[1].text = 'ФИТКБ'  # Hаименование факультета
    doc.tables[0].cell(0, 1).paragraphs[1].runs[1].underline = WD_UNDERLINE.SINGLE
    doc.tables[0].cell(0, 1).paragraphs[3].runs[1].text = 'А.В. Бредихин'  # Декан факультета
    doc.tables[0].cell(0, 1).paragraphs[1].runs[2].underline = WD_UNDERLINE.SINGLE
    doc.paragraphs[8].runs[1].text = dictInf['Название']  # наименование дисциплины
    doc.paragraphs[8].runs[1].underline = WD_UNDERLINE.SINGLE
    doc.paragraphs[12].runs[3].text = dictInf['Специальность']  # Направление подготовки
    doc.paragraphs[12].runs[3].underline = WD_UNDERLINE.SINGLE
    doc.paragraphs[14].runs[2].text = dictInf['Профиль']  # Профиль
    doc.paragraphs[14].runs[2].underline = WD_UNDERLINE.SINGLE
    doc.paragraphs[16].runs[1].text = dictInf['Квалификация']  # Квалификация выпускника
    doc.paragraphs[16].runs[1].underline = WD_UNDERLINE.SINGLE
    doc.paragraphs[18].runs[1].text = '2022-2023' + '/' + '2022-2023'  # Нормативный период обучения
    doc.paragraphs[18].runs[1].underline = WD_UNDERLINE.SINGLE
    doc.paragraphs[18].runs[2].text = ''
    doc.paragraphs[18].runs[3].text = ''
    doc.paragraphs[18].runs[4].text = ''
    doc.paragraphs[18].runs[5].text = ''
    doc.paragraphs[18].runs[6].text = ''
    doc.paragraphs[49].runs[2].text = dictInf['Название']  # Дисциплина (модуль)
    doc.paragraphs[49].runs[2].underline = WD_UNDERLINE.SINGLE
    doc.paragraphs[54].runs[1].text = dictInf['Название']  # Процесс изучения дисциплины
    doc.paragraphs[54].runs[1].underline = WD_UNDERLINE.SINGLE
    startRow = 2
    for object in dictInf['Компетенции'].keys():
        doc.tables[1].cell(startRow, 0).text = object
        startRow += 2
    for _ in range(4 - len(dictInf['Компетенции'])):
        for _ in range(3):
            __DeleteRow(doc.tables[1].rows[-1])
    doc.paragraphs[64].runs[1].text = dictInf['Название']  # Общая трудоемкость дисциплины
    doc.paragraphs[64].runs[1].underline = WD_UNDERLINE.SINGLE
    dictTime = dictInf['Часы']
    match len(dictTime):
        case 1:
            startCol = 4
            allCol = 3
            for _ in range(3 * 13):
                __DeleteRow(doc.tables[2].rows[-1])
        case 2:
            startCol = 8
            allCol = 3
            for _ in range(13):
                __DeleteRow(doc.tables[2].rows[0])
            for _ in range(2 * 13):
                __DeleteRow(doc.tables[2].rows[-1])
        case 3:
            startCol = 5
            allCol = 3
            for _ in range(2 * 13):
                __DeleteRow(doc.tables[2].rows[0])
            for _ in range(13):
                __DeleteRow(doc.tables[2].rows[-1])
        case 4:
            startCol = 8
            allCol = 3
            for _ in range(3 * 13):
                __DeleteRow(doc.tables[2].rows[-1])
    doc.save(f"./data/{dictInf['Название']}.docx")
    doc = ReadDocxTemplate(f"./data/{dictInf['Название']}.docx")
    allTimeDict = dict()
    for key in dictTime.keys():
        if 'Практические занятия' in dictTime.get(key).keys():
            try:
                allTimeDict['Практические занятия'] += int(dictTime.get(key)['Практические занятия'])
            except:
                allTimeDict['Практические занятия'] = int(dictTime.get(key)['Практические занятия'])
        if 'Лабораторные занятия' in dictTime.get(key).keys():
            try:
                allTimeDict['Лабораторные занятия'] += int(dictTime.get(key)['Лабораторные занятия'])
            except:
                allTimeDict['Лабораторные занятия'] = int(dictTime.get(key)['Лабораторные занятия'])
        if 'Самостоятельная работа' in dictTime.get(key).keys():
            try:
                allTimeDict['Самостоятельная работа'] += int(dictTime.get(key)['Самостоятельная работа'])
            except:
                allTimeDict['Самостоятельная работа'] = int(dictTime.get(key)['Самостоятельная работа'])
        if 'Лекционные занятия' in dictTime.get(key).keys():
            try:
                allTimeDict['Лекционные занятия'] += int(dictTime.get(key)['Лекционные занятия'])
            except:
                allTimeDict['Лекционные занятия'] = int(dictTime.get(key)['Лекционные занятия'])
        if 'Итого часов' in dictTime.get(key).keys():
            try:
                allTimeDict['Итого часов'] += int(dictTime.get(key)['Итого часов'])
            except:
                allTimeDict['Итого часов'] = int(dictTime.get(key)['Итого часов'])
        if 'ЗЕТ' in dictTime.get(key).keys():
            try:
                allTimeDict['ЗЕТ'] += int(dictTime.get(key)['ЗЕТ'])
            except:
                allTimeDict['ЗЕТ'] = int(dictTime.get(key)['ЗЕТ'])
        if 'Курсовой проект' in dictTime.get(key).keys():
            allTimeDict['Курсовой проект'] = int(dictTime.get(key)['Курсовой проект'])
        if 'Контрольная работа' in dictTime.get(key).keys():
            allTimeDict['Контрольная работа'] = int(dictTime.get(key)['Контрольная работа'])
    for key in dictTime.keys():
        doc.tables[2].cell(1, startCol).paragraphs[0].runs[0].text = str(key)
        if 'Лекционные занятия' in dictTime.get(key).keys():
            doc.tables[2].cell(4, startCol).paragraphs[0].runs[0].text = dictTime.get(key)['Лекционные занятия']
            doc.tables[2].cell(4, allCol).paragraphs[0].runs[0].text = str(allTimeDict['Лекционные занятия'])
        if 'Практические занятия' in dictTime.get(key).keys():
            doc.tables[2].cell(5, startCol).paragraphs[0].runs[0].text = dictTime.get(key)['Практические занятия']
            doc.tables[2].cell(5, allCol).paragraphs[0].runs[0].text = str(allTimeDict['Практические занятия'])
        if 'Лабораторные занятия' in dictTime.get(key).keys():
            doc.tables[2].cell(6, startCol).paragraphs[0].runs[0].text = dictTime.get(key)['Лабораторные занятия']
            doc.tables[2].cell(6, allCol).paragraphs[0].runs[0].text = str(allTimeDict['Лабораторные занятия'])
        if 'Самостоятельная работа' in dictTime.get(key).keys():
            doc.tables[2].cell(7, startCol).paragraphs[0].runs[0].text = dictTime.get(key)['Самостоятельная работа']
            doc.tables[2].cell(7, allCol).paragraphs[0].runs[0].text = str(allTimeDict['Самостоятельная работа'])
        if 'Курсовой проект' in dictTime.get(key).keys():
            doc.tables[2].cell(8, startCol).paragraphs[0].runs[0].text = '+'
        else:
            doc.tables[2].cell(8, startCol).paragraphs[0].runs[0].text = '-'
        if 'Контрольная работа' in dictTime.get(key).keys():
            doc.tables[2].cell(9, startCol).paragraphs[0].runs[0].text = '+'
        else:
            doc.tables[2].cell(9, startCol).paragraphs[0].runs[0].text = '-'
        if 'Зачет' in dictTime.get(key).keys():
            doc.tables[2].cell(10, startCol).paragraphs[0].runs[0].text = 'Зачет'
        elif 'Зачет с оценкой' in dictTime.get(key).keys():
            doc.tables[2].cell(10, startCol).paragraphs[0].runs[0].text = 'Зачет с оценкой'
        elif 'Экзамен' in dictTime.get(key).keys():
            doc.tables[2].cell(10, startCol).paragraphs[0].runs[0].text = 'Экзамен'
        if 'Итого часов' in dictTime.get(key).keys():
            doc.tables[2].cell(11, startCol).paragraphs[0].runs[0].text = dictTime.get(key)['Итого часов']
            doc.tables[2].cell(11, allCol).paragraphs[0].runs[0].text = str(allTimeDict['Итого часов'])
        if 'ЗЕТ' in dictTime.get(key).keys():
            doc.tables[2].cell(12, startCol).paragraphs[0].runs[0].text = dictTime.get(key)['ЗЕТ']
            doc.tables[2].cell(12, allCol).paragraphs[0].runs[0].text = str(allTimeDict['ЗЕТ'])
        startCol += 3
    doc.paragraphs[64].runs[3].text = str(allTimeDict['ЗЕТ'])  # Общая трудоемкость дисциплины
    doc.paragraphs[64].runs[3].underline = WD_UNDERLINE.SINGLE
    for key in allTimeDict:
        match key:
            case 'Лекционные занятия':
                doc.tables[4].cell(3, 3).text = str(allTimeDict[key])
            case 'Практические занятия':
                doc.tables[4].cell(3, 4).text = str(allTimeDict[key])
            case 'Лабораторные занятия':
                doc.tables[4].cell(3, 5).text = str(allTimeDict[key])
            case 'Самостоятельная работа':
                doc.tables[4].cell(3, 6).text = str(allTimeDict[key])
            case 'Итого часов':
                doc.tables[4].cell(3, 7).text = str(allTimeDict[key])
    doc.paragraphs[158].runs[1].text = dictInf['Название']  # По дисциплине
    doc.paragraphs[158].runs[1].underline = WD_UNDERLINE.SINGLE
    startRow = 2
    for key in dictInf['Компетенции'].keys():
        doc.tables[7].cell(startRow, 0).paragraphs[0].runs[0].text = key
        startRow += 2
    for _ in range(4 - len(dictInf['Компетенции'])):
        for _ in range(3):
            __DeleteRow(doc.tables[7].rows[-1])
    startRow = 2
    for key in dictInf['Компетенции'].keys():
        doc.tables[8].cell(startRow, 0).paragraphs[0].runs[0].text = key
        startRow += 2
    for _ in range(4 - len(dictInf['Компетенции'])):
        for _ in range(3):
            __DeleteRow(doc.tables[8].rows[-1])
    # DELETE UNUSED
    if 'Контрольная работа' in allTimeDict.keys():
        __DeleteParagraph(doc.paragraphs[98])
    else:
        for _ in range(8):
            __DeleteParagraph(doc.paragraphs[99])
    if 'Курсовой проект' in allTimeDict.keys():
        __DeleteParagraph(doc.paragraphs[87])
    else:
        for _ in range(8):
            __DeleteParagraph(doc.paragraphs[88])
    __DeleteParagraph(doc.paragraphs[58])
    fullQualStr = ''
    for key in dictInf['Компетенции'].keys():
        fullQualStr+= f"{key} - {dictInf['Компетенции'][key]}\n"
    doc.paragraphs[55].text = f"{doc.paragraphs[55].text}\n {fullQualStr}"
    return doc


def GetDisciplineList(jsonData):
    list = {}
    for object in jsonData['Документ']['diffgr:diffgram']['dsMMISDB']['ПланыСтроки']:
        list[object['@Код']] = object['@Дисциплина']
    return list


def __SearchCompetenciesByDisciplineCode(disciplineCode, jsonData):
    compCodeList = []
    for object in jsonData['Документ']['diffgr:diffgram']['dsMMISDB']['ПланыКомпетенцииДисциплины']:
        if object['@КодСтроки'] == disciplineCode:
            compCodeList.append(object['@КодКомпетенции'])
    dict = {}
    for object in jsonData['Документ']['diffgr:diffgram']['dsMMISDB']['ПланыКомпетенции']:
        if compCodeList.__contains__(object['@Код']):
            dict[object['@ШифрКомпетенции']] = object['@Наименование']
    return dict


def __SearchHoursBySemesterNumber(semesterNumber, disciplineCode, jsonData):
    codeList = []
    hoursList = []
    for object in jsonData['Документ']['diffgr:diffgram']['dsMMISDB']['ПланыНовыеЧасы']:
        if (object['@КодОбъекта'] == disciplineCode) and (
                int(object['@Курс']) * 2 - 1 + int(object['@Семестр']) - 1 == semesterNumber):
            codeList.append(object['@КодВидаРаботы'])
            hoursList.append(object['@Количество'])
    nameList = []
    dict = {}
    for key in codeList:
        for object in jsonData['Документ']['diffgr:diffgram']['dsMMISDB']['СправочникВидыРабот']:
            if object['@Код'] == key:
                nameList.append(object['@Название'])
    for i in range(nameList.__len__()):
        dict[nameList[i]] = hoursList[i]

    return dict


def __SearchHours(disciplineCode, jsonData):
    dict = {}
    semesterNumberList = []
    for object in jsonData['Документ']['diffgr:diffgram']['dsMMISDB']['ПланыНовыеЧасы']:
        if object['@КодОбъекта'] == disciplineCode:
            num = int(object['@Курс']) * 2 - 1 + int(object['@Семестр']) - 1
            if semesterNumberList.__contains__(num) == False:
                semesterNumberList.append(num)
    for i in range(semesterNumberList.__len__()):
        dict[semesterNumberList[i]] = __SearchHoursBySemesterNumber(semesterNumberList[i], disciplineCode, jsonData)
    return dict


def KeyFromVal(dict, val):
    for key, value in dict.items():
        if value == val:
            return key


def ReadDocxTemplate(filePath):
    doc = Document(filePath)
    return doc


def SaveDocx(doc, fileName, path: str):
    fullPath = str(path) + str(fileName) + '.docx'
    doc.save(fullPath)


def __GetSpecialization(jsonData) -> str:
    str = jsonData['Документ']['diffgr:diffgram']['dsMMISDB']['ООП']['@Название'] + ' / ' + \
          jsonData['Документ']['diffgr:diffgram']['dsMMISDB']['ООП']['@Шифр']
    return str


def __GetProfile(jsonData) -> str:
    str = jsonData['Документ']['diffgr:diffgram']['dsMMISDB']['ООП']['ООП']['@Название']
    return str


def __GetQualification(jsonData) -> str:
    code = jsonData['Документ']['diffgr:diffgram']['dsMMISDB']['ООП']['@Квалификация']
    for object in jsonData['Документ']['diffgr:diffgram']['dsMMISDB']['Уровень_образования']:
        if object['@Код_записи'] == code:
            return object['@ВидПлана']


def GetFullInf(disciplineName: str, disciplineCode: str, plxData: dict) -> dict:
    dictInf = {}
    dictInf['Название'] = disciplineName
    dictInf['Специальность'] = __GetSpecialization(plxData)
    dictInf['Профиль'] = __GetProfile(plxData)
    dictInf['Квалификация'] = __GetQualification(plxData)
    dictInf['Компетенции'] = __SearchCompetenciesByDisciplineCode(disciplineCode, plxData)
    print(dictInf['Компетенции'])
    dictInf['Часы'] = __SearchHours(disciplineCode, plxData)
    return dictInf
