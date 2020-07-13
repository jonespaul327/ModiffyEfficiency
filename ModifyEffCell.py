import openpyxl

theFile = openpyxl.load_workbook('characterization.xlsx')
allSheetNames = theFile.sheetnames

instrumentModel = '2360/43-93'
instrumentId = '227413/PR295918'

print("All sheet names {} " .format(theFile.sheetnames))

def find_instrument_model_cell(currentSheet):
    for row in range(1, 50):
        for column in "ABCDEFGHIJKLMNOPQRSTUV":  # Here you can add or reduce the columns
            cell_name = "{}{}".format(column, row)
            if currentSheet[cell_name].value == instrumentModel:
                #print("{1} cell is located on {0}" .format(cell_name, currentSheet[cell_name].value))
                #print("cell position {} has value {}".format(cell_name, currentSheet[cell_name].value))
                print("the row is {0} and the column {1}" .format(cell_name[1], cell_name[0]))
                return cell_name

def find_instrument_sn_cell(instModelCell):

    refRow = str(int(instModelCell[1]) + 1)
    refCol = instModelCell[0]

    print('xxxxxxxxxxx')
    print(refRow)
    print(refCol)
    print(currentSheet[refCol + refRow].value)


for x in allSheetNames:
    print("Current sheet name is {}" .format(x))
    currentSheet = theFile[x]
    instModelCell = find_instrument_model_cell(currentSheet)
    if instModelCell is None:
        continue
    instSNcell = find_instrument_sn_cell(instModelCell)


    #print(currentSheet['L8'].value)

    # if currentSheet['L8'].value == instrumentId:
    #     currentSheet['N10'].value = 0.152




theFile.close()
theFile.save('characterization.xlsx')

"""
Purpose: This function is used to create a list of all files within a designated folder and then return the list. The 
            origin folder must be within the same folder as the program.
Parameter(s): The parameter is a folder name/folder path from the current directory of the program.
Return: Returns a list of all file names, including file path, within the folder as well as all subdirectories.

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
"""