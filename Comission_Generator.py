import win32com.client
import os
import sys
import datetime
import pandas as pd
from os import walk
import tkinter as tk
from tkinter import filedialog
    
# ----- Variables -----

# Product Codes
byop_product_codes = ["CUSTDEVROG", "CUSTDEVFIDO", "MULTKITCHPKG"]
tablet_product_codes = ["TAB", "IPD", "IPAD"]
apple_watches_product_codes = ["AW"]

# Comission
byop_comission = 5
activation_comission = 10
hup_comission = 5
tablet_comission = 5
appleWatch_comission = 5
dp_comission = 5

# ----- Functions -----

def getFileName():
    tk.Tk().withdraw() # prevents an empty tkinter window from appearing
    filetypes = ( ('Excel Files', '*.XLSX'), ('All files', '*.*'),)
    filename = tk.filedialog.askopenfilename( title='Select a file...', filetypes=filetypes,)
    return filename

# Copy of original excel file
def createUncorruptedExcelFile (corrupted_file_path, new_file_name):
    o = win32com.client.Dispatch("Excel.Application")
    o.Visible = False
    filename = corrupted_file_path
    output = os.getcwd() + '/' + new_file_name
    wb = o.Workbooks.Open(filename)
    wb.ActiveSheet.SaveAs(output,51)

# Open new excel file and extract specific column
def readExcelColumn (file_name, column_name):
    info = pd.read_excel(file_name)
    return info.loc[:, column_name]

# Delete a file from a location
def deleteFile (filepath):
    def delete():
        os.system('taskkill /IM "' + "EXCEL.EXE" + '" /F')
        os.system("cls")
        os.remove(filepath)
    while True:
        try:
            delete()
            break # stop the loop if the function completes sucessfully
        except Exception as e:
            print("")

def getMonthAndYear(userOption):
    if userOption == "1":
        today = datetime.date.today()
        first = today.replace(day=1)
        last_month = first - datetime.timedelta(days=1)
        last_month = str(last_month.strftime("%b_%Y"))
        return last_month
    else:
        current_month_text = datetime.date.today().strftime('%h')
        current_year_full = datetime.date.today().strftime('%Y')
        date = current_month_text + "_" + current_year_full
        return date

def checkIfFolderAndCreate(Folder_Path):
    if not os.path.isdir(Folder_Path):
        os.makedirs(Folder_Path)

# ----- Main -----

currentDir = os.getcwd()

choose = input("Press ENTER and choose the exported excel file ...")

# Choose the excel file
createUncorruptedExcelFile(getFileName(),"xxcvyzxid.xlsx")

# Extract the Data
productCode = readExcelColumn("xxcvyzxid.xlsx", "Product Code")
productCodes = []
for i in range(0, len(productCode)-1,1): 
    if productCode[i] != "nan" and productCode[i] != "NAN": productCodes.append(productCode[i])

transactionTypes = readExcelColumn("xxcvyzxid.xlsx", "Transaction Type")
transactionType = []
for i in range(0, len(transactionTypes)-1,1): 
    if transactionTypes[i] != "nan" and transactionTypes[i] != "NAN": transactionType.append(transactionTypes[i])

invoice = readExcelColumn("xxcvyzxid.xlsx", "Reference")
invoiceN, invoiceNumber = [], []
for i in range(0, len(invoice)-1,1): 
    if invoice[i] != "nan" and invoice[i] != "NAN": invoiceN.append(invoice[i])
for i in range(0, len(invoiceN), 1):
    x = str(invoiceN[i])
    x = x.replace("Invoice#: ", "")
    invoiceNumber.append(x)

an = readExcelColumn("xxcvyzxid.xlsx", "BAN")
accountNumber, ban = [], []
for i in range(0, len(an)-1,1): 
    if an[i] != "nan" and an[i] != "NAN": accountNumber.append(an[i])

for i in range(0, len(accountNumber), 1):
    x = str(accountNumber[i])
    x = x.replace(".0", "")
    ban.append(x)

number = readExcelColumn("xxcvyzxid.xlsx", "#Cell phone")
customerPhoneNumber = []
for i in range(0, len(number)-1,1): 
    if number[i] != "nan" and number[i] != "NAN": customerPhoneNumber.append(number[i])

date = readExcelColumn("xxcvyzxid.xlsx", "Activation date")
ad, activationDate = [], []
for i in range(0, len(date)-1,1): 
    if date[i] != "nan" and date[i] != "NAN": ad.append(date[i])
for i in range(0,len(ad), 1):
    x = str(ad[i])
    x = x.replace("Timestamp('", "")
    x = x.replace(" 00:00:00", "")
    activationDate.append(x)

network = readExcelColumn("xxcvyzxid.xlsx", "Network")
company = []
for i in range(0, len(network)-1,1): 
    if network[i] != "nan" and network[i] != "NAN": company.append(network[i])

# Remove the copy and the original
deleteFile (currentDir + "/xxcvyzxid.xlsx")
#deleteFile (currentDir + "\Paste_your_Excel_file_here/" + originalExcelFile)

# Filter all the information
byop, tablets, apple_watches, activations, hup, dp, total = 0,0,0,0,0,0,0

dp = input("Enter the number of DP's sold: ")
if dp == "": dp = 0
else: dp = int(dp)

for i in range(0, len(productCodes), 1):
    
    info = str(productCodes[i])
    
    if info in byop_product_codes:
        byop += 1
    elif (any(substring in info for substring in tablet_product_codes)) == True:
        tablets += 1
    elif (any(substring in info for substring in apple_watches_product_codes)) == True:
        apple_watches += 1
    else:
        activations += 1
        
total = byop + tablets + apple_watches + activations

for i in range(0, len(transactionTypes),1):
    
    info = str(transactionTypes[i])
    if info == "Renewal":
        hup += 1

activations = activations - hup

os.system("cls")
print("---------------- Comission calculator by Harjot ---------------\n")
print("Product\t\t\t     Qty x Comission \t\tTotal\n")
print (f"BYOP: \t\t\t\t{byop} x {byop_comission}$ \t=\t {byop*byop_comission}$")
print(f"Activations with phones: \t{activations} x {activation_comission}$ \t=\t {activations*activation_comission}$")
print(f"Upgrades: \t\t\t{hup} x {hup_comission}$ \t\t=\t {hup*hup_comission}$")
print(f"Tablets: \t\t\t{tablets} x {tablet_comission}$ \t\t=\t {tablets*tablet_comission}$")
print(f"Apple Watches: \t\t\t{apple_watches} x {appleWatch_comission}$ \t\t=\t {apple_watches*appleWatch_comission}$")
print(f"Device Protection: \t\t{dp} x {dp_comission}$ \t\t=\t {dp*dp_comission}$")
print(f"\nTotal: \t\t\t   {total} products sold \t\t{dp*dp_comission + byop*byop_comission + tablets*tablet_comission + apple_watches*appleWatch_comission + activations*activation_comission + hup*hup_comission + dp*dp_comission}$")
print("\n--------------------------------------------------------------\n\n")

writeToFile = input("Do you want to write this information in a word file (y/n): ")
os.system("cls")

nameOfFile = ""

if writeToFile == "y" or writeToFile == "Y":
    nameOfEmployee = input("Please enter your name: ")
    nameOfEmployee = nameOfEmployee.capitalize()
    os.system("cls")
    
    whichMonth = input("\nIs the comission for the current month or the last month\n\n1 - Last month\n2 - Current month\n\n--> ")
    last_month = getMonthAndYear(whichMonth)
    
    nameOfFile = "[" + str(last_month) + "]_" + nameOfEmployee + ".docx"
    
    checkIfFolderAndCreate("C:\Comission_History")
    
    with open(os.path.join('C:\Comission_History',nameOfFile), 'w') as sys.stdout:
        print(f"\t[Employee name: {nameOfEmployee}]\t\t[Comission for {last_month}]\n\n")
        print("\t------------ Comission calculator by Harjot ------------\n")
        print("\tProduct\t\t\t  Qty x Comission \t\tTotal\n")
        print (f"\tBYOP: \t\t\t\t{byop} x {byop_comission}$ \t\t=\t {byop*byop_comission}$")
        print(f"\tActivations with phones: \t{activations} x {activation_comission}$ \t\t=\t {activations*activation_comission}$")
        print(f"\tUpgrades: \t\t\t\t{hup} x {hup_comission}$ \t\t=\t {hup*hup_comission}$")
        print(f"\tTablets: \t\t\t\t{tablets} x {tablet_comission}$ \t\t=\t {tablets*tablet_comission}$")
        print(f"\tApple Watches: \t\t\t{apple_watches} x {appleWatch_comission}$ \t\t=\t {apple_watches*appleWatch_comission}$")
        print(f"\tDevice Protection: \t\t{dp} x {dp_comission}$ \t\t=\t {dp*dp_comission}$")
        print(f"\n\tTotal: \t\t\t   {byop + activations + hup + tablets + apple_watches} products sold \t\t{ dp*dp_comission + byop*byop_comission + tablets*tablet_comission + apple_watches*appleWatch_comission + activations*activation_comission + hup*hup_comission + dp*dp_comission}$")
        print("\n\t--------------------------------------------------------\n")
        print("\n Print page 1 for summary or print all pages for detailed view.")
        print('\n All reports are saved in "C:\Comission_History" by default.')
        
        for i in range(17): print("\n")
        print("\t\t\t\t\tDetailed View\n")
        print("Invoice\tBAN\t\tNet\tDate\t\tComission\tProduct")
        print("-------\t---\t\t---\t----\t\t---------\t-------\n")
        for i in range(0, len(customerPhoneNumber),1):
            x = ''
            if company[i] == "Rogers": x = "Rog"
            else: x = "Fid"
            comission = "0"
            info = str(productCodes[i])
            if info in byop_product_codes:
                comission = " 5"
            elif (any(substring in info for substring in tablet_product_codes)) == True:
                comission = " 5"
            elif (any(substring in info for substring in apple_watches_product_codes)) == True:
                comission = " 5"
            else:
                comission = "10"
            
            if ban[i] == "nan":
                print(f"{invoiceNumber[i]}\t\tNo BAN\t\t{x}\t{activationDate[i]}\t\t{comission} $\t{productCodes[i]}")
            else:
                print(f"{invoiceNumber[i]}\t\t{ban[i]}\t{x}\t{activationDate[i]}\t\t{comission} $\t{productCodes[i]}")
            
    sys.stdout.close()
    
    command = 'start wordpad ' + 'C:\Comission_History\\' + nameOfFile
    os.system(command)
