from log import authLog
from docx import Document

import csv

import os
import traceback

def chooseCSV():
    while True:
        csvFile = input("Please enter the path to the CSV file: ")
        try:
            with open(csvFile, "r") as csvFileFinal:
                authLog.info(f"User chose  the CSV File path: {csvFile}")
                print(f"INFO: file successfully found.")
                csvReader = csv.reader(csvFileFinal)
                csvData = list(csvReader)
                if len(csvData):
                    row = csvData[1]
                    for row in row:
                        print(row)


        except FileNotFoundError:
            print("File not found. Please check the file path and try again.")
            authLog.error(f"File not found in path {csvFile}")
            authLog.error(traceback.format_exc())
            continue

        except Exception as error:
            print(f"ERROR: {error}\n", traceback.format_exc())
            authLog.error(f"Wasn't possible to choose the CSV file, error message: {error}\n", traceback.format_exc())

def chooseDocx():
    while True:
        wordFile = input("Please enter the path to the DOCX file: ")
        try:
            with open(wordFile, "r") as wordDOC:
                wordDOC = Document(wordFile)
                authLog.info(f"User chose  the DOCX File path: {wordFile}")
                print(f"INFO: file successfully found: {wordFile}.")
                replaceText = [
                    '',
                    '',
                    ''
                ]

        except FileNotFoundError:
            print("File not found. Please check the file path and try again.")
            authLog.error(f"File not found in path {wordFile}")
            authLog.error(traceback.format_exc())
            continue

        except Exception as error:
            print(f"ERROR: {error}\n", traceback.format_exc())
            authLog.error(f"Wasn't possible to choose the DOCX file, error message: {error}\n", traceback.format_exc())
