import openpyxl
import json
import os

def main(path):
    def getListOfFiles(dirName):
        listOfFile = os.listdir(dirName)
        allFiles = list()

        for file in listOfFile:
            fullPath = os.path.join(dirName, file)
            if os.path.isdir(fullPath):
                allFiles = allFiles + getListOfFiles(fullPath)
            else:
                allFiles.append(fullPath)

        print(allFiles)

        return allFiles


    def find_instrument_model_cell(currentSheet):
        for row in range(1, 50):
            for column in "ABCDEFGHIJKLMNOPQRSTUV":  # Here you can add or reduce the columns
                modelCell = "{}{}".format(column, row)
                if currentSheet[modelCell].value == instrumentModel:
                    print("the row is {0} and the column {1}" .format(row, column))
                    print(currentSheet[modelCell].value)
                    print(modelCell)

                    return [row, column, modelCell]

        return [0, 0, None]


    def find_instrument_sn_cell(instModelRow, instModelColumn):
        snRow = str(int(instModelRow) + 1)
        snCol = instModelColumn
        snCell = currentSheet[snCol + snRow]

        print('xxxxxxxxxxx')
        print(snRow)
        print(snCol)
        print(snCell.value)

        return snCell


    def find_instrument_efficiency(instModelRow, instModelColumn):
        effRow = str(int(instModelRow) + 3)
        effCol = chr(ord(instModelColumn) + 2)
        effCell = currentSheet[effCol + effRow]

        if type(effCell).__name__ == 'MergedCell':
            effCol = chr(ord(instModelColumn) + 3)
            effCell = currentSheet[effCol + effRow]

        print(effCell.value)
        print(effRow)
        print(effCol)

        return effCell


    def modify_efficiency(instSNcell, instEfficiencyCell):
        for inst in instrumentsData:
            if inst['sn'] == instSNcell.value:
                instEfficiencyCell.value = inst['betaEfficiency']

                return inst['sn']




    instrumentModel = '2360/43-93'
    #instrumentId = '227413/PR295918'

    filesWithNoMatchingSN = list()
    sheetsOfFilesWithNoMatchingSN = list()

    files = getListOfFiles(path)

    with open('package.json') as instruments_file:
        instrumentsData = json.load(instruments_file)

    for file in files:
        theFile = openpyxl.load_workbook(file)
        allSheetNames = theFile.sheetnames

        print("All sheet names {} ".format(theFile.sheetnames))

        for x in allSheetNames:
            print("Current sheet name is {}" .format(x))
            currentSheet = theFile[x]
            instModelRow = find_instrument_model_cell(currentSheet)[0]
            instModelColumn = find_instrument_model_cell(currentSheet)[1]
            instModelCell = find_instrument_model_cell(currentSheet)[2]

            print("The cell is {}, the row is {} and the column is {} ".format(instModelCell, instModelRow, instModelColumn))

            if instModelCell is None:
                continue
            instSNcell = find_instrument_sn_cell(instModelRow, instModelColumn)
            instEfficiencyCell = find_instrument_efficiency(instModelRow, instModelColumn)
            serialNumber = modify_efficiency(instSNcell, instEfficiencyCell)

            if serialNumber is None:
                filesWithNoMatchingSN.append(file)
                sheetsOfFilesWithNoMatchingSN.append(currentSheet)

        theFile.close()
        theFile.save(file)

    print("The files with no s/n are {}, the sheet is {}".format(filesWithNoMatchingSN, sheetsOfFilesWithNoMatchingSN))
