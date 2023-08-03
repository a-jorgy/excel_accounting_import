import csv
import sys
from enum import Enum
from datetime import date
import re
import openpyxl
from extendedopenpyxl import load_workbook, save_workbook
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import  CellIsRule, FormulaRule
from datetime import datetime
from openpyxl.chart import (
    PieChart,
    Reference
)

# -- Title --
SHEET_TITLES = ['Titre', 'Type', 'Catégorie', 'Compte', 'Valeur', 'Vérification', 'Date Prélèv.', 'Date Emission', 'Description']
SHEET_TITLE_STYLE = Font(color='FFFFFF')
SHEET_TITLE_FILL = PatternFill("solid", fgColor="A5A5A5")
SHEET_TITLE_BORDER = Border(bottom=Side(border_style="medium", color="000000"))

# -- Row --
SHEET_ROW_BASIC_COLOR = PatternFill(start_color="A9D08E", end_color="A9D08E", fill_type="solid")
SHEET_ROW_GREY_COLOR = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")
SHEET_ROW_RED_FONT_COLOR = Font(color="C65911")
SHEET_ROW_BLUE_FONT_COLOR = Font(color="4772C4")
SHEET_ROW_RED_COLOR = PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid")
SHEET_ROW_BLUE_COLOR = PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid")
SHEET_ROW_OK_COLOR = PatternFill(start_color="548235", end_color="548235", fill_type="solid")
SHEET_ROW_NOK_COLOR = PatternFill(start_color="C65911", end_color="C65911", fill_type="solid")
SHEET_ROW_START_BORDER = Border(left=Side(border_style="medium", color="000000"), bottom=Side(border_style="thin", color="000000"), right=Side(border_style="dotted", color="000000"))
SHEET_ROW_MIDDLE_BORDER = Border(left=Side(border_style="dotted", color="000000"), bottom=Side(border_style="thin", color="000000"), right=Side(border_style="dotted", color="000000"))
SHEET_ROW_END_BORDER = Border(left=Side(border_style="dotted", color="000000"), bottom=Side(border_style="thin", color="000000"), right=Side(border_style="medium", color="000000"))

# -- List type depense --
SHEET_SOMME_TYPE_DEPENSE = [
    ('=Data!C2', '=SUMIF($C$2:$C$150,K3,$E$2:$E$150)'),
    ('=Data!C3', '=SUMIF($C$2:$C$150,K4,$E$2:$E$150)'),
    ('=Data!C4', '=SUMIF($C$2:$C$150,K5,$E$2:$E$150)'),
    ('=Data!C5', '=SUMIF($C$2:$C$150,K6,$E$2:$E$150)'),
    ('=Data!C6', '=SUMIF($C$2:$C$150,K7,$E$2:$E$150)'),
    ('=Data!C7', '=SUMIF($C$2:$C$150,K8,$E$2:$E$150)'),
    ('=Data!C8', '=SUMIF($C$2:$C$150,K9,$E$2:$E$150)'),
    ('=Data!C9', '=SUMIF($C$2:$C$150,K10,$E$2:$E$150)')]

class typeEnum(Enum) :
    Entre = "Entré"
    Sortie = "Sortie"
    TransfereSortie = "Transfère Sortie"
    TransfereEntre = "Transfère Entré"
    Empty = ""

class monthEnum(Enum):
    Janvier= 1
    Février= 2
    Mars= 3
    Avril= 4
    Mai = 5
    Juin= 6
    Juillet = 7
    Aout = 8
    Septembre = 9
    Octobre = 10
    Novembre = 11
    Décembre = 12

class compteEnum(Enum) :
    LCL = "LCL"
    Bourso = "Boursorama"

class operation:
    def __init__(self, type, description, montant, date, compte):
        self.type = type
        self.description = description
        self.montant = montant
        self.date = date
        self.compte = compte
    def __repr__(self):
        return repr((self.type, self.description, self.montant, self.date, self.compte))

# Main function
def main():
    # Check arguments
    if(len(sys.argv) != 3) :
        print('Usage : python compte.py <lcl csv export> <bourso csv export>')
        return
    
    operations = []
    operations.extend(convertLCLFile(sys.argv[1]))
    operations.extend(convertBoursoFile(sys.argv[2]))
    exportOperations(operations)
    compteExcel(operations)

# Convert a csv export form lcl to use it in the program
def convertLCLFile(file):
    with open(file) as csvfile:
        lclLine = csv.reader(csvfile, delimiter=';')
        lclOperation = []
        for id, row in enumerate(lclLine) :
            if len(row) == 8:
                day, month, year = re.search("(\d{2})\/(\d{2})\/(\d{4})",row[0]).groups()
                dateEvent = date(int(year), int(month), int(day))
                montant = abs(float(row[1].replace(',','.')))
                description = row[4]+row[5]
                type = typeEnum.Empty if re.match("VIR", description) else typeEnum.Sortie if re.match("-", row[1]) else typeEnum.Entre
                lclOperation.append(operation(type, description, str(montant).replace('.', ','), dateEvent, compteEnum.LCL))
        return lclOperation

# Convert a csv export from boursorama to use in the program
def convertBoursoFile(file):
    with open(file) as csvfile:
        boursoline = csv.reader(csvfile, delimiter=';')
        boursoOperation = []
        for id, row in enumerate(boursoline):
            if id != 0 :
                year, month, day = re.search("(\d{4})-(\d{2})-(\d{2})", row[0]).groups()
                dateEvent = date(int(year), int(month), int(day))
                montant = abs(float(row[5].replace(',','.')))
                description = row[2]
                type = typeEnum.Empty if re.match("VIR", description) else typeEnum.Sortie if re.match("-", row[5]) else typeEnum.Entre
                boursoOperation.append(operation(type, description, str(montant).replace('.', ','), dateEvent, compteEnum.Bourso))
        return boursoOperation


def exportOperations(operations):
    sortedOperations = sorted(operations, key=lambda operation: operation.date)
    with open('./export.csv', 'w', newline='') as file:
        writer = csv.writer(file)
        for id, row in enumerate(sortedOperations):
            writer.writerow(["", row.type.value,"", row.compte.value, row.montant, "NOK", row.date, "", row.description])

# Update the excel file
def compteExcel(operations):
    # fileDep = pd.read_excel('Dépenses_ntest.xlsx', engine='openpyxl')
    # print(fileDep.head())

    # Load the excel file 
    workbook = openpyxl.load_workbook('Dépenses_test.xlsx', keep_links=False)
    createNewSheet(workbook, 5, 2023)
    addOperations(workbook, 5, 2023, operations)
    manageStyle(workbook, 5, 2023)

    # Save and close the new excel file
    workbook.save('Dépenses_test'+str(datetime.timestamp(datetime.now()))+".xlsx")
    # workbook.close()

# Find if a sheet exist depending of this month & year
def findSheet(workbook, month, year):
    sheetToFind =  monthEnum._value2member_map_[month].name + ' ' + str(year)
    return workbook.sheetnames.index(sheetToFind)

# Create a new sheet of data at the expected month & year
def createNewSheet(workbook, month, year):
    title = monthEnum._value2member_map_[month].name + ' ' + str(year)
    workbook.create_sheet(title)
    sheet = workbook[title]

    # Create sheet titles
    for id, val in enumerate(SHEET_TITLES):
        sheet[chr(65+id)+'1'].font = SHEET_TITLE_STYLE
        sheet[chr(65+id)+'1'].fill = SHEET_TITLE_FILL
        sheet[chr(65+id)+'1'].border = SHEET_TITLE_BORDER
        sheet[chr(65+id)+'1'] = val
    
    # Create sheet type de dépense
    for id, (name, value) in enumerate(SHEET_SOMME_TYPE_DEPENSE):
        sheet['K'+str(id+3)] = name
        sheet['L'+str(id+3)] = value
    
    # Create camembert
    pie = PieChart()
    labels = Reference(sheet, min_col=11, min_row=3, max_row=10)
    data = Reference(sheet, min_col=12, min_row=3, max_row=10)
    pie.add_data(data)
    pie.set_categories(labels)
    sheet.add_chart(pie, "N2")

    # Add row to statistiques
    statistiqueSheet = workbook['Statistiques']
    rowToAddIndex = (year - 2023) * 13 + 7 + month
    if month == 1 :
        statistiqueSheet.merge_cells(start_row = rowToAddIndex - 1, start_column = 1, end_row=rowToAddIndex-1, end_column=4)
        statistiqueSheet.cell(row=rowToAddIndex - 1, column=1).value = year
    statistiqueSheet.cell(row=rowToAddIndex, column = 1).value = monthEnum._value2member_map_[month].name
    statistiqueSheet.cell(row=rowToAddIndex, column = 2).value = "=SUMIF('"+title+"'!$B$2:$B$200, Data!$A$2, '"+title+"'!$E$2:$E$200)"
    statistiqueSheet.cell(row=rowToAddIndex, column = 3).value = "=SUMIF('"+title+"'!$B$2:$B$200, Data!$A$3, '"+title+"'!$E$2:$E$200)"
    statistiqueSheet.cell(row=rowToAddIndex, column = 4).value = "="+statistiqueSheet.cell(row=rowToAddIndex, column = 2).coordinate+"-"+statistiqueSheet.cell(row=rowToAddIndex, column = 3).coordinate

    # sheet['A1'].fill = PatternFill("solid", fgColor="ff0000")
    # my_fill = openpyxl.styles.fills.PatternFill(patternType='solid', bgColor='9C0006')
    # sheet.conditional_formatting.add('$B$2:$E$150', FormulaRule(formula=['$B2=Data!$A$3'], stopIfTrue=True, fill=my_fill))

# Add the information of the operations
def addOperations(workbook, month, year, operationList):
    print("add Operations")
    # Open the sheet to find
    sheetToFind =  monthEnum._value2member_map_[month].name + ' ' + str(year)
    sheet = workbook[sheetToFind]

    # Find the line where there in no data
    rowEmpty = 2 # begin at line 2, line 1 have the titles
    while (sheet.cell(row=rowEmpty, column=5).value !=  None): # See if the cell "Valeur" is empty
        rowEmpty += 1
    
    # Add the data
    for operation in operationList: # Loop throw each operation
        # Create the line
        sheet.cell(row=rowEmpty, column=2).value = operation.type.value
        sheet.cell(row=rowEmpty, column=4).value = operation.compte.value
        sheet.cell(row=rowEmpty, column=5).data_type = 'n'
        sheet.cell(row=rowEmpty, column=5).value = float(operation.montant.replace(',','.'))
        sheet.cell(row=rowEmpty, column=6).value = "NOK"
        sheet.cell(row=rowEmpty, column=7).value = operation.date.strftime("%d/%m/%Y")
        sheet.cell(row=rowEmpty, column=9).value = operation.description
        # Increment the line to implement
        rowEmpty += 1

def manageStyle(workbook, month, year):
    sheetToFind =  monthEnum._value2member_map_[month].name + ' ' + str(year)
    sheet = workbook[sheetToFind]

    # Manage column width
    sheet.column_dimensions['A'].width = 32
    sheet.column_dimensions['B'].width = 16
    sheet.column_dimensions['C'].width = 16
    sheet.column_dimensions['D'].width = 11
    sheet.column_dimensions['E'].width = 11
    sheet.column_dimensions['F'].width = 11
    sheet.column_dimensions['G'].width = 11
    sheet.column_dimensions['H'].width = 11
    sheet.column_dimensions['I'].width = 48

    # Manage line style
    # Entré/Sortie to Valeur
    sheet.conditional_formatting.add("B2:E200", FormulaRule(formula=['$B2=Data!$A$5'], stopIfTrue=True, fill=SHEET_ROW_GREY_COLOR, font=SHEET_ROW_BLUE_FONT_COLOR))
    sheet.conditional_formatting.add("B2:E200", FormulaRule(formula=['$B2=Data!$A$4'], stopIfTrue=True, fill=SHEET_ROW_GREY_COLOR, font=SHEET_ROW_RED_FONT_COLOR))
    sheet.conditional_formatting.add("B2:E200", FormulaRule(formula=['$B2=Data!$A$2'], stopIfTrue=True, fill=SHEET_ROW_BLUE_COLOR))
    sheet.conditional_formatting.add("B2:E200", FormulaRule(formula=['$B2=Data!$A$3'], stopIfTrue=True, fill=SHEET_ROW_RED_COLOR))
    sheet.conditional_formatting.add("B2:E200", FormulaRule(formula=['ISBLANK($B2)'], stopIfTrue=True, fill=SHEET_ROW_GREY_COLOR))
    
    # Vérification
    sheet.conditional_formatting.add("F2:F200", CellIsRule(operator='equal', formula=["Data!$F$3"], stopIfTrue=True, fill=SHEET_ROW_NOK_COLOR))
    sheet.conditional_formatting.add("F2:F200", CellIsRule(operator='equal', formula=["Data!$F$2"], stopIfTrue=True, fill=SHEET_ROW_OK_COLOR))
    sheet.conditional_formatting.add("F2:F200", FormulaRule(formula=['ISBLANK(F2)'], stopIfTrue=True, fill=SHEET_ROW_GREY_COLOR))

    # Other columns
    for rowIndex in range(2,201) :
        # Fill
        sheet.cell(row=rowIndex, column=1).fill = SHEET_ROW_BASIC_COLOR
        sheet.cell(row=rowIndex, column=7).fill = SHEET_ROW_BASIC_COLOR
        sheet.cell(row=rowIndex, column=8).fill = SHEET_ROW_BASIC_COLOR
        sheet.cell(row=rowIndex, column=9).fill = SHEET_ROW_BASIC_COLOR

        # Border
        sheet.cell(row=rowIndex, column=1).border = SHEET_ROW_START_BORDER
        for columnIndex in range(2,9):
            sheet.cell(row=rowIndex, column=columnIndex).border = SHEET_ROW_MIDDLE_BORDER
        sheet.cell(row=rowIndex, column=9).border = SHEET_ROW_END_BORDER

    # Manage row data validation
    # Type
    ruleType = DataValidation(type="list", formula1="Data!$A$2:$A$5")
    ruleType.add("B2:B200")
    sheet.add_data_validation(ruleType)
    # Catégorie
    ruleCategorie = DataValidation(type="list", formula1="IF($B2 = Data!$A$2,Data!$B$2:$B$5,IF($B2 = Data!$A$3,Data!$C$2:$C$9,))")
    ruleCategorie.add("C2:C200")
    sheet.add_data_validation(ruleCategorie)
    # Compte
    ruleCompte = DataValidation(type="list", formula1="Data!$D$2:$D$6")
    ruleCompte.add("D2:D200")
    sheet.add_data_validation(ruleCompte)
    # Vérification
    ruleVerification = DataValidation(type="list", formula1="Data!$F$2:$F$3")
    ruleVerification.add("F2:F200")
    sheet.add_data_validation(ruleVerification)

if __name__ == "__main__":
    main()