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
                dictTimeZO[semestrKey]['Аудиторные занятия'] = str(int(dictTimeZO[semestrKey]['Аудиторные занятия']) + int(dictTimeZO[semestrKey][timeKey]))

    timeKeysZO = list()
    for key in dictTimeZO.keys():
        timeKeysZO.append(key)

    dictAllTimeZO = dict()
    dictAllTimeZO['Аудиторные занятия'] = str(0)
    dictAllTimeZO['allTime'] = str(0)
    for keySemestr in dictTimeZO.keys():
        for timeKey in dictTimeZO[keySemestr].keys():
            if timeKey == 'Практические занятия' or timeKey == 'Лабораторные занятия' or timeKey == 'Лекционные занятия':
                dictAllTimeZO['Аудиторные занятия'] = str(int(dictAllTimeZO['Аудиторные занятия']) + int(dictTimeZO[keySemestr][timeKey]))
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
                        cell.text = cell.text.replace('allTime', str(int(dictAllTimeZO['Аудиторные занятия']) + int(dictAllTimeZO['Самостоятельная работа'])))
                    else:
                        cell.text = ''