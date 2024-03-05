import openpyxl
import re
import math
import shutil

#sprawdz czy jest ogolne zestawienie
#pobierz dane całego słupa (4 wiersze, nieskończenie wiele w bok)
#ustaw pierwszy jako glowny m
    #oblicz kat miedzy a b
    #dla pierwszego kabla magistralnego ustaw zawsze kat 0 stopni
    #dla pozostałych w drugiej czesci obliczenia
    #dla abonenckich normalne wyliczenia

#formatuj nazwe
    #rozdziel + i - (czy ten - potrzebny? może sam +?)
    #formatuj adss na ADSS + spacja + reszta
    #formatuj elektryke - 
        #jak al to al. + spacja + reszta

#wstaw dane
#skopiuj i wstaw dane

#LG1 - NN
    #E44 E45 E46 E47
#LG1 OPTO
    #E48 E49 E50 E51

#LG2 
    #E52 E53 E54 E55
#LG2 OPTO
    #E56 E57 E58 E59

def countDeq():
    def parsTuple(str):
        return tuple(map(float, re.findall(r'\d+\.\d+', str)))

    excelFile = openpyxl.load_workbook("C:\\Users\\BFS\\Documents\\polesData.xlsx")
    sheet = excelFile.active

    #get data from cells
    tuple_a1 = parsTuple(sheet['B3'].value)
    tuple_a2 = parsTuple(sheet['B4'].value)
    tuple_b1 = parsTuple(sheet['C3'].value)
    tuple_b2 = parsTuple(sheet['C4'].value)

    #count vectors
    vector1 = tuple(b - a for a, b in zip(tuple_a1, tuple_a2))
    vector2 = tuple(b - a for a, b in zip(tuple_b1, tuple_b2))

    dotProduct = sum(a * b for a, b in zip(vector1, vector2))
    vectorLength1 = math.sqrt(sum(a**2 for a in vector1))
    vectorLength2 = math.sqrt(sum(b**2 for b in vector2))

    degCos = dotProduct / (vectorLength1 * vectorLength2)
    deg = math.degrees(math.acos(degCos))

    print(f"Kąt między wektorami wynosi: {deg} stopni")
    excelFile.save(r"C:\Users\BFS\Documents\polesData_wyniki.xlsx")
def fileSave(sourcePath, finalPath):
    #Copy excel
    shutil.copy(sourcePath, finalPath)

    # Load copied excel
    finalExcel = openpyxl.load_workbook(finalPath)

    #Set the active sheet to calculator
    finalExcel.active = finalExcel["KALKULATOR"];

    #how to access cells
        #currentSheet = finalExcel.active
        # currentSheet['E46'] = 'test'
    
    finalExcel.save(finalPath)
    finalExcel.close()
# fileSave("C:\\Users\\BFS\\Documents\\kalkulator.xlsx", 'C:\\Users\\BFS\\Documents\\test101.xlsx')
def formatString(inputString):
    #trim string
    trimmed_string = inputString.strip()

    #separate satring after +
    parts = trimmed_string.split('+')
    parts = trimmed_string.split('-')

    #format string
    for i in range(len(parts)):
        part = parts[i]

        if 'adss' in part:
            last_s_index = part.rfind('s')
            parts[i] = part[:last_s_index + 1] + ' ' + part[last_s_index + 1:]

        if 'al' in part:
            if 'l.' not in part:
                parts[i] = part.replace('l', 'l. ')
            else:
                parts[i] = part.replace('l.', 'l. ')

        if 'asxsn' in part:
            parts[i] = part.replace('n', 'n ')

    return parts
def getPoleFromData(data):
    formattedString = data.strip('()')

    separatedStrings = formattedString.split()

    pole, function, number = separatedStrings
    return {"pole":pole, "function": function, "number": number}
def ReadExcelData(file_path):
    # Load the Excel file
    workbook = openpyxl.load_workbook(file_path)
    
    # Get the active sheet
    sheet = workbook.active
    
    # Get the number of rows and columns
    rows = sheet.max_row
    cols = sheet.max_column
    
    # Initialize the data table
    data_table = []

    # Iterate through rows, starting from row 1 and incrementing by 4
    for row_start in range(1, rows + 1, 4):
        # Initialize an object for the current range
        data_object = []
        
        # Iterate through columns
        for col in range(1, cols + 1):
            # Get values from the current range
            cell_value = [sheet.cell(row=row, column=col).value for row in range(row_start, row_start + 4)]
            
            # Add values to the object
            data_object.append(cell_value)
        
        # Add the object to the data table
        data_table.append(data_object)

    return data_table
#table of tables -> one table one pole with all equipment
result_table = ReadExcelData("C:\\Users\\BFS\\Documents\\polesData_wyniki.xlsx")
# # Display the result
# for i, data_object in enumerate(result_table, start=1):
#     print(f"{data_object}")


#each table is another pole
#if length is 0 or 1 it's an error!
#if 2 it's automaticly K and deg should be 90 up front
#if more it's should check
    #if 1m and 1a it's still K
    #if not count as normal

def handle_P(data, excel):
    indicator = data[0]
    #{"pole":pole, "function": function, "number": number}
    poleData = getPoleFromData(data[1])
    cords = data[3]
  
    excel['C78'] = poleData.number;
    excel['G78'] = poleData.pole;
    excel['J78'] = poleData.function.upper();
    excel['G88'] = 15; #mufa
def handle_M(data, excel, isVectorA):
    indicator = data[0]
    cables = formatString(data[1])
    cord1 = data[2]
    cord2 = data[3]

    # print(excel["E53"].value) #None = empty
    rangeVectorA = ['E44', 'E45', 'E46', 'E47', 'E48', 'E49', 'E50', 'E51']
    rangeVectorB = ['E52', 'E53', 'E54', 'E55', 'E56', 'E57', 'E58', 'E59']
    
    #not working 😒
    def putValueToCell(cable, rangeVector):
        for i in range(len(rangeVector)):
            cell = rangeVector[i]
            excelCell = excel[cell]
            excelValue = excelCell.value
            if (excelValue == None):
                print(excelValue)
                excelCell = cable
                break

    for i in range(len(cables)):
        cable = cables[i]
        if (isVectorA):
            putValueToCell(cable, rangeVectorA)
        else:
            putValueToCell(cable, rangeVectorB)
                
def handleData(data):
    #copy / create all excel summary data
    #####

    for i in range(len(data)):
        poleData = data[i]

        if (len(poleData) <=1): 
            print("error", poleData)
            return any
        #excel to copy
        sourceExcel = "C:\\Users\\BFS\\Documents\\kalkulator.xlsx"
        #final pole calculator excel
        newExcel = f'C:\\Users\\BFS\\Documents\\test{i}.xlsx'

        

        fileSave(sourceExcel, newExcel)
        excel = openpyxl.load_workbook(newExcel)

        excel.active = excel["KALKULATOR"];
        sheet = excel.active;

        isVectorA = True;
        for j in range(len(poleData)):
            table = poleData[j]
            #for each table handle p, m, a
            indicator = table[0]
            if (indicator.upper() == 'P'):
                # handle_P(table, sheet)
                next
            if (indicator.upper() == 'M'):
                handle_M(table, sheet, isVectorA)
                isVectorA = None
            if (indicator.upper() == 'A'):
                next

        #copy all cables data
        #depending on pole function, copy data
            
handleData(result_table)

#kolejność
    #ReadExcelData
    #fileSave => nowy kalkulator
    #handleData
        #dla każdego przyupadku handle...m, handle_p itd
        #każdy przypadek wysyła instrukcję co robić do dataPutter
    #zapisz plik

