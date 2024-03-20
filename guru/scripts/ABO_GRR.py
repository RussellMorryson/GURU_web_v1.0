# -*- coding: utf-8 -*-

import openpyxl
import warnings
#import collections

# Проверка и корректировка чисел из ячеек
def Checking_Of_Value(word):
    if word == '' or word ==' ' or word =='(-)' or word =='(-)2' or word =='0' or word == None or word == 'None':
        return 0
    else:
        count = 0
        num_w = ''
        for i in word:
            if i == '(' or i == '-':
                count+=1
            if i != ' ' and i != '(' and i != '-' and i != ')':
                num_w  += i
        number = int(num_w)
        if count > 0:
            number *= -1  
        return int(number)

def Find_Reporting_Year(word):
    #На 31 декабря 2022 г.
    year = ''
    for i in range(len(word)):
        if i == 14:
            for j in range(i, i+4):
                year += word[j]
    return int(year)

#RESERVE
def analysisOfAccountingStatements(file_path:str):
    warnings.simplefilter("ignore")
    wookbook = openpyxl.load_workbook(file_path) # Открытие книги
    worksheet_1 = wookbook['Organization Info'] # Открытие первого листа книги
    alf = ['B','C','D','E','F','G','H','J','K','I','L','M','N','O','P','R','S','T','U','V']
    col = ''

    print(worksheet_1['A1'].value)
    for i in range(1, worksheet_1.max_row):
        if str(worksheet_1['A' + str(i)].value) == 'Дата формирования информации':
            for j in alf:
                # print(len(alf))
                # print(alf[2])
                # print(worksheet_1['H' + str(i)].value)
                #print(worksheet_1[j + str(i)].value)
                if str(worksheet_1[j + str(i)].value) != '' and \
                    str(worksheet_1[j + str(i)].value) != None and \
                    str(worksheet_1[j + str(i)].value) != 'None':
                    date_of_unloading =  str(worksheet_1[j + str(i)].value) # Дата выгрузки
                    #print(date_of_unloading)
                    col = j
                    break
        if str(worksheet_1['A' + str(i)].value) == 'Полное наименование юридического лица':
            full_name = str(worksheet_1[col + str(i)].value) # Полное название юридического лица
        if str(worksheet_1['A' + str(i)].value) == 'ИНН':
            i_number = str(worksheet_1[col + str(i)].value) # ИНН
        if str(worksheet_1['A' + str(i)].value) == 'Местонахождение (адрес)':
            address = str(worksheet_1[col + str(i)].value) # Местонахождение (адрес)
        if str(worksheet_1['A' + str(i)].value) == 'ОГРН/ОГРНИП':
            ogrn = str(worksheet_1[col + str(i)].value) # ОГРН/ОГРНИП

    # Обработка бухгалтерского баланса
    worksheet_2 = wookbook['Balance'] # Открытие второго листа книги

   # balance_let = 'HIJKLMNOPQRSTU' # Строка с именами столбцов    

    bal_key = ''
    bal_col1 = '' 
    y1 = ''
    bal_col2 = ''
    y2 = ''
    bal_col3 = ''
    y3 = ''
    #val = ''

    for i in range(1, 100):
        for j in alf:
            #val = str(worksheet_2[alf[j] + str(i)].value)
            if str(worksheet_2[j + str(i)].value) == 'Код строки': 
                bal_key = j
            if 'На 31 декабря' in str(worksheet_2[j + str(i)].value) and bal_col1 =='': 
                bal_col1 = j
                y1 = str(worksheet_2[j + str(i)].value)
            if 'На 31 декабря' in str(worksheet_2[j + str(i)].value) and bal_col2 =='':
                bal_col2 = j
                y2 = str(worksheet_2[j + str(i)].value)
            if 'На 31 декабря' in str(worksheet_2[j + str(i)].value) and bal_col3 =='': 
                bal_col3 = j
                y3 = str(worksheet_2[j + str(i)].value)

    # Инициализация словарей для хранения показателей
    bd_1 = {'year': 0, '1110': 0, '1120': 0, '1130': 0, '1140': 0, '1150': 0, '1160': 0, '1170': 0, '1180': 0, '1190': 0,
        '1210': 0, '1220': 0, '1230': 0, '1240': 0, '1250': 0, '1260': 0, '1310': 0, '1320': 0, '1330': 0, '1340': 0, '1350': 0, '1360': 0, '1370': 0,
        '1410': 0, '1420': 0, '1430': 0, '1440': 0, '1450': 0, '1510': 0, '1520': 0, '1530': 0, '1540': 0, '1550': 0,
        '1600': 0, '1700': 0, '2110': 0, '2120': 0, '2210': 0, '2220': 0, '2310': 0, '2320': 0, '2330': 0, '2340': 0, '2350': 0,
        '2410': 0, '2411': 0, '2412': 0, '2460': 0, '2465': 0, '2510': 0, '2520': 0, '2530': 0, '2900': 0, '2910': 0,
        '1100': 0, '1200': 0, '1300': 0, '1400': 0, '1500': 0, '2100': 0, '2200': 0, '2300': 0, '2400': 0, '2500': 0}

    bd_2 = {'year': 0, '1110': 0, '1120': 0, '1130': 0, '1140': 0, '1150': 0, '1160': 0, '1170': 0, '1180': 0, '1190': 0,
        '1210': 0, '1220': 0, '1230': 0, '1240': 0, '1250': 0, '1260': 0, '1310': 0, '1320': 0, '1330': 0, '1340': 0, '1350': 0, '1360': 0, '1370': 0,
        '1410': 0, '1420': 0, '1430': 0, '1440': 0, '1450': 0, '1510': 0, '1520': 0, '1530': 0, '1540': 0, '1550': 0,
        '1600': 0, '1700': 0, '2110': 0, '2120': 0, '2210': 0, '2220': 0, '2310': 0, '2320': 0, '2330': 0, '2340': 0, '2350': 0,
        '2410': 0, '2411': 0, '2412': 0, '2460': 0, '2465': 0, '2510': 0, '2520': 0, '2530': 0, '2900': 0, '2910': 0,
        '1100': 0, '1200': 0, '1300': 0, '1400': 0, '1500': 0, '2100': 0, '2200': 0, '2300': 0, '2400': 0, '2500': 0}

    bd_3 = {'year': 0, '1110': 0, '1120': 0, '1130': 0, '1140': 0, '1150': 0, '1160': 0, '1170': 0, '1180': 0, '1190': 0,
        '1210': 0, '1220': 0, '1230': 0, '1240': 0, '1250': 0, '1260': 0, '1310': 0, '1320': 0, '1330': 0, '1340': 0, '1350': 0, '1360': 0, '1370': 0,
        '1410': 0, '1420': 0, '1430': 0, '1440': 0, '1450': 0, '1510': 0, '1520': 0, '1530': 0, '1540': 0, '1550': 0,
        '1600': 0, '1700': 0, '2110': 0, '2120': 0, '2210': 0, '2220': 0, '2310': 0, '2320': 0, '2330': 0, '2340': 0, '2350': 0,
        '2410': 0, '2411': 0, '2412': 0, '2460': 0, '2465': 0, '2510': 0, '2520': 0, '2530': 0, '2900': 0, '2910': 0,
        '1100': 0, '1200': 0, '1300': 0, '1400': 0, '1500': 0, '2100': 0, '2200': 0, '2300': 0, '2400': 0, '2500': 0}

    # Запись годов отчетности в словари
    bd_1['year'] = int(Find_Reporting_Year(y1))
    bd_2['year'] = int(Find_Reporting_Year(y2))
    bd_3['year'] = int(Find_Reporting_Year(y3))

    # Запись информации из бухгалтерского баланса
    number = ''
    for row in range(1, 100):
        for key in bd_1:
            if key == str(worksheet_2[bal_key + str(row)].value):
                # Заполнение словаря № 1
                if str(worksheet_2[bal_col1 + str(row)].value) != '-':
                    number = str(worksheet_2[bal_col1 + str(row)].value)
                    bd_1[key] = Checking_Of_Value(number)
                # Заполнение словаря № 2
                if str(worksheet_2[bal_col2 + str(row)].value) != '-':
                    number = str(worksheet_2[bal_col2 + str(row)].value)
                    bd_2[key] = Checking_Of_Value(number)
                # Заполнение словаря № 3
                if str(worksheet_2[bal_col3 + str(row)].value) != '-':
                    number = str(worksheet_2[bal_col3 + str(row)].value)
                    bd_3[key] = Checking_Of_Value(number)

# Пересчет показателей, поскольку некоторые выгрузки не содержат информации о кодах
# Заполнение словаря № 1
    bd_1['1100'] = bd_1['1110'] + bd_1['1120'] + bd_1['1130'] + bd_1['1140'] + bd_1['1150'] + bd_1['1160'] + bd_1['1170'] + bd_1['1180'] + bd_1['1190']
    bd_1['1200'] = bd_1['1210'] + bd_1['1220'] + bd_1['1230'] + bd_1['1240'] + bd_1['1250'] + bd_1['1260']
    bd_1['1300'] = bd_1['1310'] + bd_1['1320'] + bd_1['1330'] + bd_1['1340'] + bd_1['1350'] + bd_1['1360'] + bd_1['1370']
    bd_1['1400'] = bd_1['1410'] + bd_1['1420'] + bd_1['1430'] + bd_1['1440'] + bd_1['1450']
    bd_1['1500'] = bd_1['1510'] + bd_1['1520'] + bd_1['1530'] + bd_1['1540'] + bd_1['1550']

# Заполнение словаря № 2
    bd_2['1100'] = bd_2['1110'] + bd_2['1120'] + bd_2['1130'] + bd_2['1140'] + bd_2['1150'] + bd_2['1160'] + bd_2['1170'] + bd_2['1180'] + bd_2['1190']
    bd_2['1200'] = bd_2['1210'] + bd_2['1220'] + bd_2['1230'] + bd_2['1240'] + bd_2['1250'] + bd_2['1260']
    bd_2['1300'] = bd_2['1310'] + bd_2['1320'] + bd_2['1330'] + bd_2['1340'] + bd_2['1350'] + bd_2['1360'] + bd_2['1370']
    bd_2['1400'] = bd_2['1410'] + bd_2['1420'] + bd_2['1430'] + bd_2['1440'] + bd_2['1450']
    bd_2['1500'] = bd_2['1510'] + bd_2['1520'] + bd_2['1530'] + bd_2['1540'] + bd_2['1550']

# Заполнение словаря № 3
    bd_3['1100'] = bd_3['1110'] + bd_3['1120'] + bd_3['1130'] + bd_3['1140'] + bd_3['1150'] + bd_3['1160'] + bd_3['1170'] + bd_3['1180'] + bd_3['1190']
    bd_3['1200'] = bd_3['1210'] + bd_3['1220'] + bd_3['1230'] + bd_3['1240'] + bd_3['1250'] + bd_3['1260']
    bd_3['1300'] = bd_3['1310'] + bd_3['1320'] + bd_3['1330'] + bd_3['1340'] + bd_3['1350'] + bd_3['1360'] + bd_3['1370']
    bd_3['1400'] = bd_3['1410'] + bd_3['1420'] + bd_3['1430'] + bd_3['1440'] + bd_3['1450']
    bd_3['1500'] = bd_3['1510'] + bd_3['1520'] + bd_3['1530'] + bd_3['1540'] + bd_3['1550']

###########################################################################################
# Обработка отчета о финансовых результатах
    worksheet_3 = wookbook['Financial Result'] # Открытие третьего листа книги

    #finance_let = 'HIJKLMNOPQRSTU' # Строка с именами столбцов    
    #word_num = '' 

    # Запись информации из отчета о прибылях и убытках
    fin_key = ''
    fin_col1 = ''
    fin_col2 = ''

    for i in range(1, 100):
        for j in alf:
            if str(worksheet_3[j + str(i)].value) == 'Код строки' and fin_key == '':
                fin_key = j
            if 'За ' in str(worksheet_3[j + str(i)].value) and fin_col1 == '':
                fin_col1 = j
            if 'За ' in str(worksheet_3[j + str(i)].value) and fin_col2 == '':
                fin_col2 = j

    for row in range(1, 100):
        for key in bd_1:
            if key == str(worksheet_3[fin_key + str(row)].value):
                # Заполнение словаря № 1
                if str(worksheet_3[fin_col1 + str(row)].value) != '-':
                    number = str(worksheet_3[fin_col1 + str(row)].value)
                    bd_1[key] = Checking_Of_Value(number)
                # Заполнение словаря № 2
                if str(worksheet_3[fin_col2 + str(row)].value) != '-':
                    number = str(worksheet_3[fin_col2 + str(row)].value)
                    bd_2[key] = Checking_Of_Value(number)

# Инициализация словарей с показателями
    indi_1 = {'year': 0, 'SP': 0.0, 'EP': 0.0, 'AP': 0.0, 'MP': 0.0, 'TermLR': 0.0, 'CLR': 0.0, 'TotalLR': 0.0, 'ALR': 0.0, 'CR': 0.0,
          'SRC3': 0.0, 'SRC6': 0.0, 'SRC9': 0.0, 'SRC12': 0.0, 'TROAR': 0.0, 'TROAP': 0.0, 'TPOAR': 0.0, 'TPOAP': 0.0,
          'CPF': 0.0, 'RARAP': 0.0, 'STDR': 0.0, 'FR': 0.0, 'BFC0': 0.0, 'BFC': 0.0, 'FSC': 0.0, 'FIC': 0.0, 'FDR': 0.0, 'ASF': 0.0}

    indi_2 = {'year': 0, 'SP': 0.0, 'EP': 0.0, 'AP': 0.0, 'MP': 0.0, 'TermLR': 0.0, 'CLR': 0.0, 'TotalLR': 0.0, 'ALR': 0.0, 'CR': 0.0,
          'SRC3': 0.0, 'SRC6': 0.0, 'SRC9': 0.0, 'SRC12': 0.0, 'TROAR': 0.0, 'TROAP': 0.0, 'TPOAR': 0.0, 'TPOAP': 0.0,
          'CPF': 0.0, 'RARAP': 0.0, 'STDR': 0.0, 'FR': 0.0, 'BFC0': 0.0, 'BFC': 0.0, 'FSC': 0.0, 'FIC': 0.0, 'FDR': 0.0, 'ASF': 0.0}

    indi_3 = {'year': 0, 'SP': 0.0, 'EP': 0.0, 'AP': 0.0, 'MP': 0.0, 'TermLR': 0.0, 'CLR': 0.0, 'TotalLR': 0.0, 'ALR': 0.0, 'CR': 0.0,
          'SRC3': 0.0, 'SRC6': 0.0, 'SRC9': 0.0, 'SRC12': 0.0, 'TROAR': 0.0, 'TROAP': 0.0, 'TPOAR': 0.0, 'TPOAP': 0.0,
          'CPF': 0.0, 'RARAP': 0.0, 'STDR': 0.0, 'FR': 0.0, 'BFC0': 0.0, 'BFC': 0.0, 'FSC': 0.0, 'FIC': 0.0, 'FDR': 0.0, 'ASF': 0.0}

# Инициализация словаря с расшифровкой
    decoding_dic = {'year': 'Год', 
                'SP': 'Рентабельность продаж', 
                'EP': 'Экономическая рентабельность', 
                'AP': 'Рентабельность активов', 
                'MP': 'Валовая рентабельность', 
                'TermLR': 'Коэффициент срочной ликвидности', 
                'CLR': 'Коэффициент текущей ликвидности', 
                'TotalLR': 'Коэффициент общей ликвидности', 
                'ALR': 'Коэффициент абсолютной ликвидности', 
                'CR': 'Коэффициент покрытия',
                'SRC3': 'Коэффициент восстановления платежеспособности 3м', 
                'SRC6': 'Коэффициент восстановления платежеспособности 6м', 
                'SRC9': 'Коэффициент восстановления платежеспособности 9м', 
                'SRC12':'Коэффициент восстановления платежеспособности 12м', 
                'TROAR': 'Коэффициент оборачиваемости дебиторской задолженности', 
                'TROAP': 'Коэффициент оборачиваемости кредиторской задолженности', 
                'TPOAR': 'Срок оборачиваемости дебиторской задолженности', 
                'TPOAP': 'Срок оборачиваемости кредиторской задолженности',
                'CPF': 'Коэффициент обеспеченности собственными средствами', 
                'RARAP': 'Соотношение дебиторской и кредиторской задолженностей', 
                'STDR': 'Коэффициент краткосрочной задолженности', 
                'FR': 'Коэффициент привлечения кредств', 
                'BFC0':'Коэффициент прогноза банкротства (0% НДС)', 
                'BFC': 'Коэффициент прогноза банкротства', 
                'FSC': 'Коэффициент финансовой устойчивости', 
                'FIC': 'Коэффициент финансовой независимости', 
                'FDR': 'Коэффициент финансовой зависимости', 
                'ASF': 'Коэффициент обеспеченности собст. ист. финансирования'}

# Заполнение словарей с показателями
    indi_1['year'] = bd_1['year']
    indi_2['year'] = bd_2['year']
    indi_3['year'] = bd_3['year']

# 1.1 Рентабельность  
# 1.1.2 Рентабельность продаж 'SP'  
    try: indi_1['SP'] = bd_1['2200'] / bd_1['2110']
    except ZeroDivisionError: indi_1['SP'] = 0
    
  # 1.1.2 Экономическая рентабельность 'EP'    
    try: indi_1['EP'] = bd_1['2300'] / ((bd_1['1600'] + bd_2['1600']) / 2)
    except ZeroDivisionError: indi_1['EP'] = 0
      
# 1.1.3 Рентабельность активов 'AP'   
    try: indi_1['AP'] = bd_1['2400'] / ((bd_1['1600'] + bd_2['1600']) / 2)
    except ZeroDivisionError: indi_1['AP'] = 0
    
# 1.1.4 Валовая рентабельность 'MP'
    try: indi_1['MP'] = bd_1['2100'] / bd_1['2110']
    except ZeroDivisionError: indi_1['MP'] = 0                    

# 1.2 Коэффициенты ликвидности
# 1.2.1 Коэффициент срочной ликвидности 'TermLR' 
    try: indi_1['TermLR'] = (bd_1['1240'] + bd_1['1250'] + bd_1['1260']) / (bd_1['1500'] - bd_1['1530'] - bd_1['1540'])
    except ZeroDivisionError: indi_1['TermLR'] = 0
    
# 1.2.2 Коэффициент текущей ликвидности 'CLR'
    try: indi_1['CLR'] = bd_1['1200'] / (bd_1['1500'] - bd_1['1530'] - bd_1['1540'])
    except ZeroDivisionError: indi_1['CLR'] = 0
  
# 1.2.3 Коэффициент общей ликвидности 'TotalLR'
    try: indi_1['TotalLR'] = bd_1['1200'] / (bd_1['1400'] + bd_1['1500'] - bd_1['1530'] - bd_1['1540'])                   
    except ZeroDivisionError: indi_1['TotalLR'] = 0
  
# 1.2.4 Коэффициент абсолютной ликвидности 'ALR'
    try: indi_1['ALR'] = (bd_1['1240'] + bd_1['1250']) / (bd_1['1500'] - bd_1['1530'] - bd_1['1540'])                     
    except ZeroDivisionError: indi_1['ALR'] = 0
  
# 1.2.5 Коэффициент покрытия 'CR'
    try: indi_1['CR'] = (bd_1['1200'] + bd_1['1170']) / (bd_1['1500'] - bd_1['1530'] - bd_1['1540'])                       
    except ZeroDivisionError: indi_1['CR'] = 0

# 1.3 Коэффициенты оборачиваемости
# 1.3.1 Коэффициент оборачиваемости дебиторской задолженности 'TROAR'
    try: indi_1['TROAR'] = bd_1['2110'] / ((bd_1['1230'] + bd_2['1230']) / 2) 
    except ZeroDivisionError: indi_1['TROAR'] = 0
      
# 1.3.2 Коэффициент оборачиваемости кредиторской задолженности 'TROAP'    
    try: indi_1['TROAP'] = bd_1['2110'] / ((bd_1['1520'] + bd_2['1520']) / 2)
    except ZeroDivisionError: indi_1['TROAP'] = 0
      
# 1.3.3 Срок оборачиваемости дебиторской задолженности 'TPOAR'    
    try: indi_1['TPOAR'] = (365 * 0.5 * (bd_1['1230'] + bd_2['1230'])) / bd_1['2110']
    except ZeroDivisionError: indi_1['TPOAR'] = 0
      
# 1.3.4 Срок оборачиваемости кредиторской задолженности 'TPOAP'
    try: indi_1['TPOAP'] = 365 / indi_1['TROAP']     
    except ZeroDivisionError: indi_1['TPOAP'] = 0                         

# 1.4 Коэффициенты рыночной стоимости
# 1.4.1 Коэффициент обеспеченности собственными средствами 'CPF'
    try: indi_1['CPF'] = (bd_1['1300'] - bd_1['1100']) / bd_1['1200']
    except ZeroDivisionError: indi_1['CPF'] = 0
    
# 1.4.2 Соотношение дебиторской и кредиторской задолженностей 'RARAP'
    try: indi_1['RARAP'] = bd_1['1230'] / bd_1['1520']
    except ZeroDivisionError: indi_1['RARAP'] = 0
    
# 1.4.3 Коэффициент краткосрочной задолженности 'STDR'
    try: indi_1['STDR'] = bd_1['1500'] / (bd_1['1400'] + bd_1['1500'])
    except ZeroDivisionError: indi_1['STDR'] = 0
    
# 1.4.4 Коэффициент привлечения кредств 'FR'
    try: indi_1['FR'] = bd_1['1400'] / (bd_1['1400'] + bd_1['1300'])
    except ZeroDivisionError: indi_1['FR'] = 0
    
# 1.4.5 Коэффициент прогноза банкротства (0% НДС) 'BFC0'   
    try: indi_1['BFC0'] = (bd_1['1200'] - bd_1['1500']) / bd_1['1700']
    except ZeroDivisionError: indi_1['BFC0'] = 0
    
# 1.4.6 Коэффициент прогноза банкротства 'BFC'
    try: indi_1['BFC'] = (bd_1['1200'] + bd_1['1180'] - bd_1['1500']) / bd_1['1700']     
    except ZeroDivisionError: indi_1['BFC'] = 0
    
# 1.5 Коэффициенты финансовой устойчивости
# 1.5.1 Коэффициент финансовой устойчивости 'FSC'
    try: indi_1['FSC'] = (bd_1['1300'] + bd_1['1400']) / bd_1['1700']
    except ZeroDivisionError: indi_1['FSC'] = 0
    
# 1.5.2 Коэффициент финансовой независимости 'FIC'
    try: indi_1['FIC'] = bd_1['1300'] / bd_1['1700']
    except ZeroDivisionError: indi_1['FIC'] = 0
      
# 1.5.3 Коэффициент финансовой зависимости 'FDR'    
    try:  indi_1['FDR'] = bd_1['1700'] / bd_1['1300']             
    except ZeroDivisionError: indi_1['FDR'] = 0
    
# 1.5.4 Коэффициент обеспеченности собст. ист. финансирования 'ASF'
    try: indi_1['ASF'] = (bd_1['1300'] - bd_1['1100']) / bd_1['1200']   
    except ZeroDivisionError: indi_1['ASF'] = 0

###########################################
# 2 Вычисление индикаторов поза-прошлого года
# 2.1 Рентабельность  
# 2.1.1 Рентабельность продаж 'SP'    
    try: indi_2['SP'] = bd_2['2200'] / bd_2['2110']                            
    except ZeroDivisionError: indi_2['SP'] = 0
    
 # 2.1.2 Экономическая рентабельность 'EP'    
    try: indi_2['EP'] = bd_2['2300'] / ((bd_2['1600'] + bd_3['1600']) / 2)    
    except ZeroDivisionError: indi_2['EP'] = 0
     
# 2.1.3 Рентабельность активов 'AP'    
    try: indi_2['AP'] = bd_2['2400'] / ((bd_2['1600'] + bd_3['1600']) / 2)     
    except ZeroDivisionError: indi_2['AP'] = 0
    
# 2.1.4 Валовая рентабельность 'MP' 
    try: indi_2['MP'] = bd_2['2100'] / bd_2['2110']        
    except ZeroDivisionError: indi_2['MP'] = 0        

# 2.2 Коэффициенты ликвидности
# 2.2.1 Коэффициент срочной ликвидности 'TermLR'    
    try: indi_2['TermLR'] = (bd_2['1240'] + bd_2['1250'] + bd_2['1260']) / (bd_2['1500'] - bd_2['1530'] - bd_2['1540'])  
    except ZeroDivisionError: indi_2['TermLR'] = 0
    
# 2.2.2 Коэффициент текущей ликвидности 'CLR'    
    try: indi_2['CLR'] = bd_2['1200'] / (bd_2['1500'] - bd_2['1530'] - bd_2['1540'])                                      
    except ZeroDivisionError: indi_2['CLR'] = 0

# 2.2.3 Коэффициент общей ликвидности 'TotalLR'    
    try: indi_2['TotalLR'] = bd_2['1200'] / (bd_2['1400'] + bd_2['1500'] - bd_2['1530'] - bd_2['1540'])
    except ZeroDivisionError: indi_2['TotalLR'] = 0

# 2.2.4 Коэффициент абсолютной ликвидности 'ALR'    
    try: indi_2['ALR'] = (bd_2['1240'] + bd_2['1250']) / (bd_2['1500'] - bd_2['1530'] - bd_2['1540'])
    except ZeroDivisionError: indi_2['ALR'] = 0

# 2.2.5 Коэффициент покрытия 'CR'
    try: indi_2['CR'] = (bd_2['1200'] + bd_2['1170']) / (bd_2['1500'] - bd_2['1530'] - bd_2['1540'])
    except ZeroDivisionError: indi_2['CR'] = 0

# 2.3 Коэффициенты оборачиваемости    
# 2.3.1 Коэффициент оборачиваемости дебиторской задолженности 'TROAR'    
    try: indi_2['TROAR'] = bd_2['2110'] / ((bd_2['1230'] + bd_3['1230']) / 2)  
    except ZeroDivisionError: indi_2['TROAR'] = 0
      
# 2.3.2 Коэффициент оборачиваемости кредиторской задолженности 'TROAP'    
    try:  indi_2['TROAP'] = bd_2['2110'] / ((bd_2['1520'] + bd_3['1520']) / 2)       
    except ZeroDivisionError: indi_2['TROAP'] = 0
           
# 2.3.3 Срок оборачиваемости дебиторской задолженности 'TPOAR'    
    try: indi_2['TPOAR'] = (365 * 0.5 * (bd_2['1230'] + bd_3['1230'])) / bd_2['2110']    
    except ZeroDivisionError: indi_2['TPOAR'] = 0
      
# 2.3.4 Срок оборачиваемости кредиторской задолженности 'TPOAP'
    try: indi_2['TPOAP'] = 365 / indi_2['TROAP']      
    except ZeroDivisionError: indi_2['TPOAP'] = 0                                   

# 2.4 Коэффициенты рыночной стоимости    
# 2.4.1 Коэффициент обеспеченности собственными средствами 'CPF'    
    try: indi_2['CPF'] = (bd_2['1300'] - bd_2['1100']) / bd_2['1200']            
    except ZeroDivisionError: indi_2['CPF'] = 0
      
# 2.4.2 Соотношение дебиторской и кредиторской задолженностей 'RARAP'    
    try: indi_2['RARAP'] = bd_2['1230'] / bd_2['1520']    
    except ZeroDivisionError: indi_2['RARAP'] = 0
                                   
# 2.4.3 Коэффициент краткосрочной задолженности 'STDR'    
    try: indi_2['STDR'] = bd_2['1500'] / (bd_2['1400'] + bd_2['1500'])          
    except ZeroDivisionError: indi_2['STDR'] = 0
            
# 2.4.4 Коэффициент привлечения кредств 'FR'    
    try: indi_2['FR'] = bd_2['1400'] / (bd_2['1400'] + bd_2['1300'])       
    except ZeroDivisionError: indi_2['FR'] = 0
                  
# 2.4.5 Коэффициент прогноза банкротства (0% НДС) 'BFC0'    
    try: indi_2['BFC0'] = (bd_2['1200'] - bd_2['1500']) / bd_2['1700']          
    except ZeroDivisionError: indi_2['BFC0'] = 0
             
# 2.4.6 Коэффициент прогноза банкротства 'BFC'
    try: indi_2['BFC'] = (bd_2['1200'] + bd_2['1180'] - bd_2['1500']) / bd_2['1700']   
    except ZeroDivisionError: indi_2['BFC'] = 0

# 2.5 Коэффициенты финансовой устойчивости    
# 2.5.1 Коэффициент финансовой устойчивости 'FSC'
    try: indi_2['FSC'] = (bd_2['1300'] + bd_2['1400']) / bd_2['1700']       
    except ZeroDivisionError: indi_2['FSC'] = 0
        
# 2.5.1 Коэффициент финансовой независимости 'FIC'    
    try: indi_2['FIC'] = bd_2['1300'] / bd_2['1700']        
    except ZeroDivisionError: indi_2['FIC'] = 0
                     
# 2.5.2 Коэффициент финансовой зависимости 'FDR'    
    try: indi_2['FDR'] = bd_2['1700'] / bd_2['1300']     
    except ZeroDivisionError: indi_2['FDR'] = 0
                        
# 2.5.3 Коэффициент обеспеченности собст. ист. финансирования 'ASF'
    try: indi_2['ASF'] = (bd_2['1300'] - bd_2['1100']) / bd_2['1200']    
    except ZeroDivisionError: indi_2['ASF'] = 0

###########################################
# 3 Вычисление индикаторов поза-поза-прошлого года
# 3.1 Рентабельность    
# 3.1.1 Рентабельность продаж 'SP'  
    try: indi_3['SP'] = bd_3['2200'] / bd_3['2110']   
    except ZeroDivisionError: indi_3['SP'] = 0
          
# 3.1.2 Валовая рентабельность 'MP'
    try: indi_3['MP'] = bd_3['2100'] / bd_3['2110']    
    except ZeroDivisionError: indi_3['MP'] = 0

# 3.2 Коэффициенты ликвидности
# 3.2.1 Коэффициент срочной ликвидности 'TermLR'    
    try: indi_3['TermLR'] = (bd_3['1240'] + bd_3['1250'] + bd_3['1260']) / (bd_3['1500'] - bd_3['1530'] - bd_3['1540'])    
    except ZeroDivisionError: indi_3['TermLR'] = 0
      
# 3.2.2 Коэффициент текущей ликвидности 'CLR'    
    try: indi_3['CLR'] = bd_3['1200'] / (bd_3['1500'] - bd_3['1530'] - bd_3['1540'])  
    except ZeroDivisionError: indi_3['CLR'] = 0
                                           
# 3.2.3 Коэффициент общей ликвидности 'TotalLR'    
    try: indi_3['TotalLR'] = bd_3['1200'] / (bd_3['1400'] + bd_3['1500'] - bd_3['1530'] - bd_3['1540'])    
    except ZeroDivisionError: indi_3['TotalLR'] = 0
                     
# 3.2.4 Коэффициент абсолютной ликвидности 'ALR'    
    try: indi_3['ALR'] = (bd_3['1240'] + bd_3['1250']) / (bd_3['1500'] - bd_3['1530'] - bd_3['1540'])   
    except ZeroDivisionError: indi_3['ALR'] = 0
                         
# 3.2.5 Коэффициент покрытия 'CR'
    try: indi_3['CR'] = (bd_3['1200'] + bd_3['1170']) / (bd_3['1500'] - bd_3['1530'] - bd_3['1540'])  
    except ZeroDivisionError: indi_3['CR'] = 0
                          
# 3.3 Коэффициенты рыночной стоимости    
# 3.3.1 Коэффициент обеспеченности собственными средствами 'CPF'    
    try: indi_3['CPF'] = (bd_3['1300'] - bd_3['1100']) / bd_3['1200'] 
    except ZeroDivisionError: indi_3['CPF'] = 0
                         
# 3.3.2 Соотношение дебиторской и кредиторской задолженностей 'RARAP'    
    try: indi_3['RARAP'] = bd_3['1230'] / bd_3['1520']         
    except ZeroDivisionError: indi_3['RARAP'] = 0
                                
# 3.3.3 Коэффициент краткосрочной задолженности 'STDR'    
    try: indi_3['STDR'] = bd_3['1500'] / (bd_3['1400'] + bd_3['1500'])       
    except ZeroDivisionError: indi_3['STDR'] = 0
                  
# 3.3.4 Коэффициент привлечения кредств 'FR'    
    try: indi_3['FR'] = bd_3['1400'] / (bd_3['1400'] + bd_3['1300'])      
    except ZeroDivisionError: indi_3['FR'] = 0
                     
# 3.3.5 Коэффициент прогноза банкротства (0% НДС) 'BFC0'    
    try: indi_3['BFC0'] = (bd_3['1200'] - bd_3['1500']) / bd_3['1700']    
    except ZeroDivisionError: indi_3['BFC0'] = 0
                    
# 3.3.6 Коэффициент прогноза банкротства 'BFC'
    try: indi_3['BFC'] = (bd_3['1200'] + bd_3['1180'] - bd_3['1500']) / bd_3['1700']     
    except ZeroDivisionError: indi_3['BFC'] = 0

# 3.4 Коэффициенты финансовой устойчивости  
# 3.4.1 Коэффициент финансовой устойчивости 'FSC'   
    try: indi_3['FSC'] = (bd_3['1300'] + bd_3['1400']) / bd_3['1700']    
    except ZeroDivisionError: indi_3['FSC'] = 0
       
# 3.4.2 Коэффициент финансовой независимости 'FIC'    
    try: indi_3['FIC'] = bd_3['1300'] / bd_3['1700']    
    except ZeroDivisionError: indi_3['FIC'] = 0
                           
# 3.4.3 Коэффициент финансовой зависимости 'FDR'  
    try: indi_3['FDR'] = bd_3['1700'] / bd_3['1300']      
    except ZeroDivisionError: indi_3['FDR'] = 0
                          
# 3.4.4 Коэффициент обеспеченности собст. ист. финансирования 'ASF'
    try: indi_3['ASF'] = (bd_3['1300'] - bd_3['1100']) / bd_3['1200']     
    except ZeroDivisionError: indi_3['ASF'] = 0

# Проведение анализа показателей
    text_result = '' # явная инициализация строки результата
    text_result += 'Дата выгрузки: ' + date_of_unloading + '\n'
    text_result += 'Полное наименование юридического лица: ' + full_name + '\n'
    text_result += 'Местонахождение (адрес): ' + address + '\n'
    text_result += 'ИНН: ' + i_number + '\n'
    text_result += 'ОГРН/ОГРНИП: ' + ogrn + '\n'
    text_result += '\tАнализируя бухгалтерскую отчетность из открытых источников, за' + str(bd_1['year']) + ' отчетный период, \nобъем выручки составил '
    text_result += str(bd_1['2110']) + ' тыс. рублей, из которой себестоимость продаж \nсоставила ' + str(bd_1['2120']*(-1)) + ' тыс. рублей, '
    if bd_1['2300'] > 0: text_result += 'прибыль до налогообложения ' + str(bd_1['2300']) + ' тыс. рублей, \n'
    if bd_1['2400'] > 0: text_result += 'а чистая прибыль в размере ' + str(bd_1['2400']) + ' тыс. рублей. '
    else: text_result += 'а убыток составил ' + str(bd_1['2400']) + ' тыс. рублей. \n'
    text_result += 'Объем выручки в сравнении с прошлым отчетным периодом '

    if bd_1['2110'] > bd_2['2110']: 
        text_result += 'увеличился на ' + str(bd_1['2110'] - bd_2['2110']) + ' тыс. рублей или на \n'
        try: 
            text_result += str("{:.2f}".format((bd_1['2110'] / bd_2['2110'])*100)) + ' процентов. \n'
        except ZeroDivisionError: 
            text_result += '100.00 процентов. '
    else:   
        text_result += 'уменьшился на ' + str(bd_2['2110'] - bd_1['2110']) + ' тыс. рублей или на '
        try: 
            text_result += str("{:.2f}".format((bd_2['2110'] - bd_1['2110'])/bd_2['2110']*100)) + ' процентов. \n'
        except ZeroDivisionError:
            text_result += '100.00 процентов. '

    text_result += 'Размер чистой прибыли '
    if bd_1['2400'] > bd_2['2400']: 
        text_result += 'увеличился на ' + str(bd_1['2400'] - bd_2['2400']) + ' тыс. рублей или на '
        try: 
            text_result += str("{:.2f}".format(((bd_1['2400'] - bd_2['2400']) / bd_1['2400'])*100)) + ' процентов. \n'
        except ZeroDivisionError: 
            text_result += ' 100.00 процентов. '          
    else: 
        text_result += 'уменьшился на ' + str(bd_2['2400'] - bd_1['2400']) + ' тыс. рублей или на '
        try:
            text_result += str("{:.2f}".format(((bd_2['2400'] - bd_1['2400']) / bd_2['2400'])*100)) + ' процентов. \n'
        except ZeroDivisionError:
            text_result += ' 100.00 процентов. '
      
    if indi_1['EP'] + indi_2['EP'] >= 0.02*2:
        text_result += 'У организации хороший показатель экономической рентабельности. \n'
    else:
        text_result += 'У организации низкий показатель экономической рентабельности. \n' 

    text_result += f'\n\tКоэффициент текущей ликвидности за последний отчетный период составляет ' + str("{:.2f}".format(indi_1['CLR'])) + '. '    
    if indi_1['CLR'] + indi_2['CLR'] > 3*2:
        text_result += 'Организация оплачивает текущие счета. Структура капитала распределена нерационально. '
    elif indi_1['CLR'] + indi_2['CLR'] > 1*2:
        text_result += 'Организация оплачивает текущие счета. Платежеспособность в норме. '
    else:
        text_result += 'Организация не оплачивает текущие счета. Структура капитала распределена нерационально. '

    text_result += f'Показатели абсолютной ликвидности за последние два отчетных периода составили ' + str("{:.2f}".format(indi_1['ALR'])) + ' и ' + str("{:.2f}".format(indi_2['ALR'])) + ', '   
    if indi_1['ALR'] + indi_2['ALR'] > 0.5*2:
        text_result += 'что говорит о высоком уровне ликвидности. Возможное нерациональное распределение активов. \n \
            Объем активов покрывает краткосрочные обязательства в размере коэффициента. '
    elif indi_1['ALR'] + indi_2['ALR'] > 0.2*2:
        text_result += 'что говорит о нормальном уровне ликвидности. Объем активов покрывает краткосрочные \n \
            обязательства в размере коэффициента. '
    else:
        text_result += 'что говорит о низком уровне ликвидности, где объем активов покрывает краткосрочные \n \
            обязательства в размере коэффициента. '

    text_result += '\n\tПоказатель оборачиваемости дебиторской задолженности, демонстрирует '
    if indi_1['TROAR'] + indi_2['TROAR'] > 2.5*2:
        text_result += 'хорошую платежную дисциплину контрагентов. Производится своевременное погашение платежей дебиторами. '
    else:
        text_result += 'низкую платежную дисциплину контрагентов. Образуется отсрочка поступлений. '

    text_result += f'Коэффициент оборачиваемости кредиторской задолженности на уровне ' + str("{:.2f}".format(indi_1['TROAP'])) + ' говорит о '  
    if indi_1['TROAP'] + indi_2['TROAP'] > 2.5*2:
        text_result += 'своевременном погашении обязательств. '
    else:
        text_result += 'постепенном образовании отсрочки платежей. '

    text_result += 'Соотношение дебиторской и кредиторской задолженности находится '  
    if indi_1['RARAP'] + indi_2['RARAP'] > 1.2*2:
        text_result += 'на высоком уровне. У организации отвлечены средства из хозяйственного оборота. '     
    elif indi_1['RARAP'] + indi_2['RARAP'] < 1.2*2 and indi_1['RARAP'] + indi_2['RARAP'] > 1*2:
        text_result += 'в пределах нормы. Организация финансово устойчива. '
    else:
        text_result += ' на низком уровне. Присутствует угроза финансовой устойчивости организации. '

    text_result += '\n\tПоказатель обеспеченности собственными средствами находится на '  
    if indi_1['CPF'] + indi_2['CPF'] > 0.1 * 2:
        text_result += 'хорошем уровне. Структура оборотного капитала в норме. '
    elif indi_1['CPF'] + indi_2['CPF'] < 0.1 * 2 and indi_1['CPF'] + indi_2['CPF'] > 0:
        text_result += 'низком уровне. Структура баланса является неудовлетворительной. '      
    else:
        text_result += 'низком уровне. Оборотные средства сформированы за счет заемных средств. '
        
    text_result += 'Расчетный коэффициент банкротства организации, рассчитанный путем отношения суммы \n \
        оборотных активов и отложенных налоговых обязательств за минусом краткосрочных обязательств к валюте \
            пассива баланса, показывает '  
    if indi_1['BFC'] + indi_2['BFC'] > 0.5 * 2:
        text_result += 'низкую вероятность банкротства в течении 6 месяцев. '
    elif indi_1['BFC'] + indi_2['BFC'] < 0.5 * 2 and indi_1['BFC'] + indi_2['BFC'] > 0:
        text_result += 'среднюю вероятность банкротства в течении 6 месяцев. '
    else:
        text_result += 'высокую вероятность банкротства в течении 6 месяцев. '   

    text_result += 'Показатель финансовой устойчивости находится на '
    if indi_1['FSC'] + indi_2['FSC'] > 0.75 * 2:
        text_result += 'хорошем уровне, при котором организация финансово устойчива, присутствует тенденция к росту. \n'
    else:
        text_result += 'низком уровне, при котором возможно ухудшение финансовой устойчивости организации. \n'
  
    text_result += 'Показатель финансовой независимости организации находится на '
    if indi_1['FIC'] + indi_2['FIC'] > 1 * 2:
        text_result += 'высоком уровне, при этом происходит сдерживание темпов развития предприятия. '
    elif indi_1['FIC'] + indi_2['FIC'] < 1 * 2 and indi_1['FIC'] + indi_2['FIC'] > 0.25 * 2:
        text_result += 'хорошем уровне. '
    else:
        text_result += 'низком уровне, что говорит о возможном ухудшении финансового состояния. '

# Сортировка
    b1 = list(bd_1.keys())
    b1.sort()
    bs_1 = {i: bd_1[i] for i in b1}

    b2 = list(bd_2.keys())
    b2.sort()
    bs_2 = {i: bd_2[i] for i in b2}

    b3 = list(bd_3.keys())
    b3.sort()
    bs_3 = {i: bd_3[i] for i in b3}
    
# Сохранение показателей
    log_str = ''
    log_str += 'Полное наименование юридического лица: ' + str(full_name) + '\n'
    log_str += 'ИНН: ' + str(i_number) + '\n'
    log_str += 'Дата выгрузки: ' + str(date_of_unloading) + '\n\n'    
    log_str += 'Год\t\t : \t\t ' + str(bs_1['year']) + ' \t\t ' + str(bs_2['year']) + ' \t\t ' + str(bs_3['year']) + '\n'
    for key in bs_1:
        if key != 'year':
            log_str += str(key) + ' \t\t : \t\t ' + str(bs_1[key]) + ' \t\t ' + str(bs_2[key]) + ' \t\t ' + str(bs_3[key]) + '\n'
    
    log_str += '\n ' + str(decoding_dic['year']) + ' \t\t : \t\t ' + str(bs_1['year']) + ' \t\t ' + str(bs_2['year']) + ' \t\t ' + str(bs_3['year']) + '\n'
    
    for key in decoding_dic:
        if key != 'year':
            log_str += str(decoding_dic[key]) + ' \t\t : \t\t ' + str(indi_1[key]) + ' \t\t ' + str(indi_2[key]) + ' \t\t ' + str(indi_3[key]) + '\n'

    #text_result += '\n\n LOGGER: \n ' + str(log_str)
    # new_path = str(file_path)
    # arr_path = []

    # for i in new_path:
    #     arr_path.append(i)

    # new_path = ''
    # for i in range(len(arr_path)-1, 0):
    #     if arr_path[i] == '\\':
    #         for j in range(i):
    #             new_path += arr_path[j]
    #         break

    # res_file = open(new_path + '\\result\\' + full_name + '\\Result_' + i_number + '.txt', 'w')
    # res_file.write(text_result)
    # res_file.close()

    # log_file = open(new_path + '\\result\\' + full_name + '\\Log_' + i_number + '.txt', 'w')
    # log_file.write(log_str)
    # log_file.close()

    # for i in range(len(text_result)):
    #     if text_result[i] == '\\' and text_result[i+1] == 'n':
    #         text_result[i] = '<br'
    #         text_result[i+1] = '>'

    # for i in range(len(log_str)):
    #     if log_str[i] == '\\' and log_str[i+1] == 'n':
    #         log_str[i] = '<br'
    #         log_str[i+1] = '>'

    # text_result += log_str
    # text_result = 'Задача выполнена!<br><br>' + text_result
    return text_result

