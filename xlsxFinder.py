import os
import pyinputplus as pyip

def findXlsxFiles():
    # Scan the excel file in the folder
    files = []
    for file in os.listdir():
        if file.endswith(".xlsx"):
            files.append(file)

    print("Which spreadsheet is the one with all the instructors?")
    file1 = pyip.inputMenu(files, numbered=True)

    print("\nWhich spreadsheet contains the experienced instructors?")
    file2 = pyip.inputMenu(files, numbered=True)

    return file1, file2