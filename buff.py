def NGG(dictInfOO: dict, dictInfZO: dict, doc: Document):
    dictTimeOO = dictInfOO['Часы']
    kursRabFlagOO = False
    konRabFlagOO = False
    zed = int(0)
    for key in dictTimeOO.keys():
        semestrDict = dictTimeOO[key]
        if 'Курсовой проект' in semestrDict.keys():
            kursRabFlagOO = True
        if 'Контрольная работа' in semestrDict.keys():
            konRabFlagOO = True
        if 'ЗЕТ' in semestrDict.keys():
            zed += int(semestrDict['ЗЕТ'])

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
        if 'zed' in object.text:
            for run in object.runs:
                run.text = run.text.replace('zed', str(zed))
        if 'compList' in object.text:
            compStr = str()
            for key in dictInfOO['Компетенции'].keys():
                compStr += key + ' - ' + dictInfOO['Компетенции'][key] + '\n'
            object.text = object.text.replace('compList', compStr)
        if ('KPY' in object.text):
            for run in object.runs:
                run.text = run.text.replace('KPY', '')
            if (kursRabFlagOO == False):
                __DeleteParagraph(object)
        if ('KPN' in object.text):
            for run in object.runs:
                run.text = run.text.replace('KPN', '')
            if (kursRabFlagOO == True):
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
                dictTimeOO[semestrKey]['Аудиторные занятия'] = str(int(dictTimeOO[semestrKey]['Аудиторные занятия']) + int(dictTimeOO[semestrKey][timeKey]))

    timeKeysOO = list()
    for key in dictTimeOO.keys():
        timeKeysOO.append(key)

    dictAllTimeOO = dict()
    dictAllTimeOO['Аудиторные занятия'] = str(0)
    dictAllTimeOO['allTime'] = str(0)
    for keySemestr in dictTimeOO.keys():
        for timeKey in dictTimeOO[keySemestr].keys():
            if timeKey == 'Практические занятия' or timeKey == 'Лабораторные занятия' or timeKey == 'Лекционные занятия':
                dictAllTimeOO['Аудиторные занятия'] = str(int(dictAllTimeOO['Аудиторные занятия']) + int(dictTimeOO[keySemestr][timeKey]))
                try:
                    dictAllTimeOO[timeKey] = str(int(dictAllTimeOO[timeKey]) + int(dictTimeOO[keySemestr][timeKey]))
                except:
                    dictAllTimeOO[timeKey] = dictTimeOO[keySemestr][timeKey]
            if timeKey == 'Самостоятельная работа' or timeKey == 'Итого часов':
                try:
                    dictAllTimeOO[timeKey] = str(int(dictAllTimeOO[timeKey]) + int(dictTimeOO[keySemestr][timeKey]))
                except:
                    dictAllTimeOO[timeKey] = dictTimeOO[keySemestr][timeKey]
    dictAllTimeOO['allTime'] = dictAllTimeOO['Аудиторные занятия'] + dictAllTimeOO['Самостоятельная работа']



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
                __DeleteRow(timeTableOO.rows[-1])

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
                cell.text = str(zed)
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
                        cell.text = cell.text.replace('allTime', str(int(dictAllTimeOO['Аудиторные занятия']) + int(dictAllTimeOO['Самостоятельная работа'])))
                    else:
                        cell.text = ''

    dictTimeZO = dictInfOO['Часы']
    kursRabFlagZO = False
    konRabFlagZO = False

    for key in dictTimeZO.keys():
        semestrDict = dictTimeZO[key]
        if 'Курсовой проект' in semestrDict.keys():
            kursRabFlagZO = True
        if 'Контрольная работа' in semestrDict.keys():
            konRabFlagZO = True

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
    dictAllTimeZO['allTime'] = dictAllTimeZO['Аудиторные занятия'] + dictAllTimeZO['Самостоятельная работа']

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
                __DeleteRow(timeTableZO.rows[-1])

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
                cell.text = str(zed)
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
                if 'Курсовой проект' in dictTimeZO[timeKeysZO[0]].keys():
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

    return doc

def GenerateDocxOchZ(dictInfO: dict, dictInfZ: dict, doc: Document):
    doc.tables[0].cell(0, 1).paragraphs[1].runs[1].text = 'ФИТКБ'  # Hаименование факультета
    doc.tables[0].cell(0, 1).paragraphs[1].runs[1].underline = WD_UNDERLINE.SINGLE
    doc.tables[0].cell(0, 1).paragraphs[3].runs[1].text = 'А.В. Бредихин'  # Декан факультета
    doc.paragraphs[8].runs[0].text = f"{dictInfO['Название']}"  # наименование дисциплины
    doc.paragraphs[8].runs[0].underline = WD_UNDERLINE.SINGLE
    doc.paragraphs[12].runs[3].text = dictInfO['Специальность']  # Направление подготовки
    doc.paragraphs[12].runs[3].underline = WD_UNDERLINE.SINGLE
    doc.paragraphs[14].runs[2].text = dictInfO['Профиль']  # Профиль
    doc.paragraphs[14].runs[2].underline = WD_UNDERLINE.SINGLE
    doc.paragraphs[16].runs[1].text = dictInfO['Квалификация']  # Квалификация выпускника
    doc.paragraphs[16].runs[1].underline = WD_UNDERLINE.SINGLE
    doc.paragraphs[18].runs[1].text = '2022-2023' + '/' + '2022-2023'  # Нормативный период обучения
    doc.paragraphs[18].runs[1].underline = WD_UNDERLINE.SINGLE
    doc.paragraphs[49].runs[1].text = f"{dictInfO['Название']}"  # Дисциплина (модуль)
    doc.paragraphs[49].runs[1].underline = WD_UNDERLINE.SINGLE
    doc.paragraphs[54].runs[1].text = f"{dictInfO['Название']}"  # Процесс изучения дисциплины
    doc.paragraphs[54].runs[1].underline = WD_UNDERLINE.SINGLE
    startRow = 2
    for object in dictInfO['Компетенции'].keys():
        doc.tables[1].cell(startRow, 0).text = object
        startRow += 2
    for _ in range(4 - len(dictInfO['Компетенции'])):
        for _ in range(3):
            __DeleteRow(doc.tables[1].rows[-1])
    doc.paragraphs[64].runs[1].text = f"{dictInfO['Название']}"  # Общая трудоемкость дисциплины
    doc.paragraphs[64].runs[1].underline = WD_UNDERLINE.SINGLE
    dictTimeOch = dictInfO['Часы']
    match len(dictTimeOch):
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
                __DeleteRow(doc.tables[2].rows[0])
    allTimeDictOch = dict()
    for key in dictTimeOch.keys():
        if 'Практические занятия' in dictTimeOch.get(key).keys():
            try:
                allTimeDictOch['Практические занятия'] += int(dictTimeOch.get(key)['Практические занятия'])
            except:
                allTimeDictOch['Практические занятия'] = int(dictTimeOch.get(key)['Практические занятия'])
        if 'Лабораторные занятия' in dictTimeOch.get(key).keys():
            try:
                allTimeDictOch['Лабораторные занятия'] += int(dictTimeOch.get(key)['Лабораторные занятия'])
            except:
                allTimeDictOch['Лабораторные занятия'] = int(dictTimeOch.get(key)['Лабораторные занятия'])
        if 'Самостоятельная работа' in dictTimeOch.get(key).keys():
            try:
                allTimeDictOch['Самостоятельная работа'] += int(dictTimeOch.get(key)['Самостоятельная работа'])
            except:
                allTimeDictOch['Самостоятельная работа'] = int(dictTimeOch.get(key)['Самостоятельная работа'])
        if 'Лекционные занятия' in dictTimeOch.get(key).keys():
            try:
                allTimeDictOch['Лекционные занятия'] += int(dictTimeOch.get(key)['Лекционные занятия'])
            except:
                allTimeDictOch['Лекционные занятия'] = int(dictTimeOch.get(key)['Лекционные занятия'])
        if 'Итого часов' in dictTimeOch.get(key).keys():
            try:
                allTimeDictOch['Итого часов'] += int(dictTimeOch.get(key)['Итого часов'])
            except:
                allTimeDictOch['Итого часов'] = int(dictTimeOch.get(key)['Итого часов'])
        if 'ЗЕТ' in dictTimeOch.get(key).keys():
            try:
                allTimeDictOch['ЗЕТ'] += int(dictTimeOch.get(key)['ЗЕТ'])
            except:
                allTimeDictOch['ЗЕТ'] = float(dictTimeOch.get(key)['ЗЕТ'])
        if 'Курсовой проект' in dictTimeOch.get(key).keys():
            allTimeDictOch['Курсовой проект'] = int(dictTimeOch.get(key)['Курсовой проект'])
        if 'Контрольная работа' in dictTimeOch.get(key).keys():
            allTimeDictOch['Контрольная работа'] = int(dictTimeOch.get(key)['Контрольная работа'])
    for key in dictTimeOch.keys():
        doc.tables[2].cell(1, startCol).paragraphs[0].runs[0].text = str(key)
        if 'Лекционные занятия' in dictTimeOch.get(key).keys():
            doc.tables[2].cell(4, startCol).paragraphs[0].runs[0].text = dictTimeOch.get(key)['Лекционные занятия']
            doc.tables[2].cell(4, allCol).paragraphs[0].runs[0].text = str(allTimeDictOch['Лекционные занятия'])
        if 'Практические занятия' in dictTimeOch.get(key).keys():
            doc.tables[2].cell(5, startCol).paragraphs[0].runs[0].text = dictTimeOch.get(key)['Практические занятия']
            doc.tables[2].cell(5, allCol).paragraphs[0].runs[0].text = str(allTimeDictOch['Практические занятия'])
        if 'Лабораторные занятия' in dictTimeOch.get(key).keys():
            doc.tables[2].cell(6, startCol).paragraphs[0].runs[0].text = dictTimeOch.get(key)['Лабораторные занятия']
            doc.tables[2].cell(6, allCol).paragraphs[0].runs[0].text = str(allTimeDictOch['Лабораторные занятия'])
        if 'Самостоятельная работа' in dictTimeOch.get(key).keys():
            doc.tables[2].cell(7, startCol).paragraphs[0].runs[0].text = dictTimeOch.get(key)['Самостоятельная работа']
            doc.tables[2].cell(7, allCol).paragraphs[0].runs[0].text = str(allTimeDictOch['Самостоятельная работа'])
        if 'Курсовой проект' in dictTimeOch.get(key).keys():
            doc.tables[2].cell(8, startCol).paragraphs[0].runs[0].text = '+'
        else:
            doc.tables[2].cell(8, startCol).paragraphs[0].runs[0].text = '-'
        if 'Контрольная работа' in dictTimeOch.get(key).keys():
            doc.tables[2].cell(9, startCol).paragraphs[0].runs[0].text = '+'
        else:
            doc.tables[2].cell(9, startCol).paragraphs[0].runs[0].text = '-'
        if 'Зачет' in dictTimeOch.get(key).keys():
            doc.tables[2].cell(10, startCol).paragraphs[0].runs[0].text = 'Зачет'
        elif 'Зачет с оценкой' in dictTimeOch.get(key).keys():
            doc.tables[2].cell(10, startCol).paragraphs[0].runs[0].text = 'Зачет с оценкой'
        elif 'Экзамен' in dictTimeOch.get(key).keys():
            doc.tables[2].cell(10, startCol).paragraphs[0].runs[0].text = 'Экзамен'
        if 'Итого часов' in dictTimeOch.get(key).keys():
            doc.tables[2].cell(11, startCol).paragraphs[0].runs[0].text = dictTimeOch.get(key)['Итого часов']
            doc.tables[2].cell(11, allCol).paragraphs[0].runs[0].text = str(allTimeDictOch['Итого часов'])
        if 'ЗЕТ' in dictTimeOch.get(key).keys():
            doc.tables[2].cell(12, startCol).paragraphs[0].runs[0].text = dictTimeOch.get(key)['ЗЕТ']
            doc.tables[2].cell(12, allCol).paragraphs[0].runs[0].text = str(allTimeDictOch['ЗЕТ'])
        startCol += 3
    doc.paragraphs[64].runs[5].text = str(allTimeDictOch['ЗЕТ'])  # Общая трудоемкость дисциплины
    doc.paragraphs[64].runs[5].underline = WD_UNDERLINE.SINGLE
    for key in allTimeDictOch:
        match key:
            case 'Лекционные занятия':
                doc.tables[4].cell(3, 3).text = str(allTimeDictOch[key])
            case 'Практические занятия':
                doc.tables[4].cell(3, 4).text = str(allTimeDictOch[key])
            case 'Лабораторные занятия':
                doc.tables[4].cell(3, 5).text = str(allTimeDictOch[key])
            case 'Самостоятельная работа':
                doc.tables[4].cell(3, 6).text = str(allTimeDictOch[key])
            case 'Итого часов':
                doc.tables[4].cell(3, 7).text = str(allTimeDictOch[key])
    # ___________________________________________________________________
    dictTimeZ = dictInfZ['Часы']
    match len(dictTimeZ):
        case 1:
            startCol = 4
            allCol = 3
            for _ in range(3 * 13):
                __DeleteRow(doc.tables[3].rows[-1])
        case 2:
            startCol = 8
            allCol = 3
            for _ in range(13):
                __DeleteRow(doc.tables[3].rows[0])
            for _ in range(2 * 13):
                __DeleteRow(doc.tables[3].rows[-1])
        case 3:
            startCol = 5
            allCol = 3
            for _ in range(2 * 13):
                __DeleteRow(doc.tables[3].rows[0])
            for _ in range(13):
                __DeleteRow(doc.tables[3].rows[-1])
        case 4:
            startCol = 6
            allCol = 3
            for _ in range(13):
                __DeleteRow(doc.tables[3].rows[0])
            for _ in range(13):
                __DeleteRow(doc.tables[3].rows[13])
            for _ in range(13):
                __DeleteRow(doc.tables[3].rows[0])
    allTimeDictZ = dict()
    for key in dictTimeZ.keys():
        if 'Практические занятия' in dictTimeZ.get(key).keys():
            try:
                allTimeDictZ['Практические занятия'] += int(dictTimeZ.get(key)['Практические занятия'])
            except:
                allTimeDictZ['Практические занятия'] = int(dictTimeZ.get(key)['Практические занятия'])
        if 'Лабораторные занятия' in dictTimeZ.get(key).keys():
            try:
                allTimeDictZ['Лабораторные занятия'] += int(dictTimeZ.get(key)['Лабораторные занятия'])
            except:
                allTimeDictZ['Лабораторные занятия'] = int(dictTimeZ.get(key)['Лабораторные занятия'])
        if 'Самостоятельная работа' in dictTimeZ.get(key).keys():
            try:
                allTimeDictZ['Самостоятельная работа'] += int(dictTimeZ.get(key)['Самостоятельная работа'])
            except:
                allTimeDictZ['Самостоятельная работа'] = int(dictTimeZ.get(key)['Самостоятельная работа'])
        if 'Лекционные занятия' in dictTimeZ.get(key).keys():
            try:
                allTimeDictZ['Лекционные занятия'] += int(dictTimeZ.get(key)['Лекционные занятия'])
            except:
                allTimeDictZ['Лекционные занятия'] = int(dictTimeZ.get(key)['Лекционные занятия'])
        if 'Итого часов' in dictTimeZ.get(key).keys():
            try:
                allTimeDictZ['Итого часов'] += int(dictTimeZ.get(key)['Итого часов'])
            except:
                allTimeDictZ['Итого часов'] = int(dictTimeZ.get(key)['Итого часов'])
        if 'ЗЕТ' in dictTimeZ.get(key).keys():
            try:
                allTimeDictZ['ЗЕТ'] += int(dictTimeZ.get(key)['ЗЕТ'])
            except:
                allTimeDictZ['ЗЕТ'] = int(dictTimeZ.get(key)['ЗЕТ'])
        if 'Курсовой проект' in dictTimeZ.get(key).keys():
            allTimeDictZ['Курсовой проект'] = int(dictTimeZ.get(key)['Курсовой проект'])
        if 'Контрольная работа' in dictTimeZ.get(key).keys():
            allTimeDictZ['Контрольная работа'] = int(dictTimeZ.get(key)['Контрольная работа'])
    for key in dictTimeZ.keys():
        doc.tables[3].cell(1, startCol).paragraphs[0].runs[0].text = str(key)
        if 'Лекционные занятия' in dictTimeZ.get(key).keys():
            doc.tables[3].cell(4, startCol).paragraphs[0].runs[0].text = dictTimeZ.get(key)['Лекционные занятия']
            doc.tables[3].cell(4, allCol).paragraphs[0].runs[0].text = str(allTimeDictZ['Лекционные занятия'])
        if 'Практические занятия' in dictTimeZ.get(key).keys():
            doc.tables[3].cell(5, startCol).paragraphs[0].runs[0].text = dictTimeZ.get(key)['Практические занятия']
            doc.tables[3].cell(5, allCol).paragraphs[0].runs[0].text = str(allTimeDictZ['Практические занятия'])
        if 'Лабораторные занятия' in dictTimeZ.get(key).keys():
            doc.tables[3].cell(6, startCol).paragraphs[0].runs[0].text = dictTimeZ.get(key)['Лабораторные занятия']
            doc.tables[3].cell(6, allCol).paragraphs[0].runs[0].text = str(allTimeDictZ['Лабораторные занятия'])
        if 'Самостоятельная работа' in dictTimeZ.get(key).keys():
            doc.tables[3].cell(7, startCol).paragraphs[0].runs[0].text = dictTimeZ.get(key)['Самостоятельная работа']
            doc.tables[3].cell(7, allCol).paragraphs[0].runs[0].text = str(allTimeDictZ['Самостоятельная работа'])
        if 'Курсовой проект' in dictTimeZ.get(key).keys():
            doc.tables[3].cell(8, startCol).paragraphs[0].runs[0].text = '+'
        else:
            doc.tables[3].cell(8, startCol).paragraphs[0].runs[0].text = '-'
        if 'Контрольная работа' in dictTimeZ.get(key).keys():
            doc.tables[3].cell(9, startCol).paragraphs[0].runs[0].text = '+'
        else:
            doc.tables[3].cell(9, startCol).paragraphs[0].runs[0].text = '-'
        if 'Зачет' in dictTimeZ.get(key).keys():
            doc.tables[3].cell(10, startCol).paragraphs[0].runs[0].text = 'Зачет'
        elif 'Зачет с оценкой' in dictTimeZ.get(key).keys():
            doc.tables[3].cell(10, startCol).paragraphs[0].runs[0].text = 'Зачет с оценкой'
        elif 'Экзамен' in dictTimeZ.get(key).keys():
            doc.tables[3].cell(10, startCol).paragraphs[0].runs[0].text = 'Экзамен'
        if 'Итого часов' in dictTimeZ.get(key).keys():
            doc.tables[3].cell(11, startCol).paragraphs[0].runs[0].text = dictTimeZ.get(key)['Итого часов']
            doc.tables[3].cell(11, allCol).paragraphs[0].runs[0].text = str(allTimeDictZ['Итого часов'])
        # if 'ЗЕТ' in dictTimeZ.get(key).keys():
        #     doc.tables[3].cell(12, startCol).paragraphs[0].runs[0].text = dictTimeZ.get(key)['ЗЕТ']
        #     doc.tables[3].cell(12, allCol).paragraphs[0].runs[0].text = str(allTimeDictZ['ЗЕТ'])
        startCol += 2 + startCol % 3
    for key in allTimeDictZ:
        match key:
            case 'Лекционные занятия':
                doc.tables[5].cell(3, 3).text = str(allTimeDictZ[key])
            case 'Практические занятия':
                doc.tables[5].cell(3, 4).text = str(allTimeDictZ[key])
            case 'Лабораторные занятия':
                doc.tables[5].cell(3, 5).text = str(allTimeDictZ[key])
            case 'Самостоятельная работа':
                doc.tables[5].cell(3, 6).text = str(allTimeDictZ[key])
            case 'Итого часов':
                doc.tables[5].cell(3, 7).text = str(allTimeDictZ[key])
    # ________________________________________________________________
    doc.paragraphs[158].runs[3].text = f"{dictInfO['Название']}"  # По дисциплине
    doc.paragraphs[158].runs[3].underline = WD_UNDERLINE.SINGLE
    startRow = 2
    for key in dictInfO['Компетенции'].keys():
        doc.tables[7].cell(startRow, 0).paragraphs[0].runs[0].text = key
        startRow += 2
    for _ in range(4 - len(dictInfO['Компетенции'])):
        for _ in range(3):
            __DeleteRow(doc.tables[7].rows[-1])
    startRow = 2
    for key in dictInfO['Компетенции'].keys():
        doc.tables[8].cell(startRow, 0).paragraphs[0].runs[0].text = key
        startRow += 2
    for _ in range(4 - len(dictInfO['Компетенции'])):
        for _ in range(3):
            __DeleteRow(doc.tables[8].rows[-1])
    # DELETE UNUSED
    if 'Контрольная работа' in allTimeDictOch.keys():
        __DeleteParagraph(doc.paragraphs[98])
    else:
        for _ in range(8):
            __DeleteParagraph(doc.paragraphs[99])
    if 'Курсовой проект' in allTimeDictOch.keys():
        __DeleteParagraph(doc.paragraphs[87])
    else:
        for _ in range(8):
            __DeleteParagraph(doc.paragraphs[88])
    __DeleteParagraph(doc.paragraphs[58])
    fullQualStr = ''
    for key in dictInfO['Компетенции'].keys():
        fullQualStr += f"{key} - {dictInfO['Компетенции'][key]}\n"
    doc.paragraphs[55].text = f"{doc.paragraphs[55].text}\n {fullQualStr}"
    return doc

    stringKursPrOO = str()
    if len(listKursPrOO) != 0:
        stringKursPrOO = f"{listKursPrOO[0]} семестре для очной формы обучения"
    if len(listKursPrOO) > 1:
        for index in range(1, len(listKursPrOO)):
            stringKursPrOO += f", в {listKursPrOO[index]} семестре для очной формы обучения"
    stringKursPrZO = str()
    if len(listKursPrZO) != 0:
        stringKursPrZO = f"{listKursPrZO[0]} семестре для очной формы обучения"
    if len(listKursPrZO) > 1:
        for index in range(1, len(listKursPrZO)):
            stringKursPrZO += f", в {listKursPrZO[index]} семестре для очной формы обучения"

    stringKursPr = stringKursPrOO + stringKursPrZO