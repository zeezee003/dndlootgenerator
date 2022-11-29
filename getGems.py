import openpyxl as xl

class GetGems(){



def getIndex():
    arrayIndex = lootSheet['M42'].value
    print(arrayIndex)

def getGems():
    wb = xl.load_workbook(filename='dndloot.xlsx')
    lootSheet = wb['Loot']

    pass
}