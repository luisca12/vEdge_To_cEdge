from log import authLog
from docx import Document
from docx.shared import RGBColor

import re
import os
import csv
import traceback

def chooseCSV():
    ignoreStrings = re.compile(r'(FALSE|TRUE|100000|biz-internet|private5|^ge0\/0$|^ge0\/1$|^ge0\/1$|^ge0\/2$|^ge0\/3$)')

    while True:
        csvFile = input("Please enter the path to the CSV file: ")
        try:
            with open(csvFile, "r") as csvFileFinal:
                authLog.info(f"User chose  the CSV File path: {csvFile}")
                print(f"INFO: file successfully found.")
                csvReader = csv.reader(csvFileFinal)
                csvData = list(csvReader)
                if csvData:
                    rowText = csvData[1]
                    filteredRowText = [row for row in rowText if not ignoreStrings.search(row)]
                    print("Found the following strings in the CSV file:")
                    for row in filteredRowText:
                        print(f"{row}")
                    authLog.info(f"Found the following strings in the CSV file:\n{filteredRowText}")
                    return filteredRowText
                else:
                    print(f"INFO: Cells not found under file: {csvFile}")
                    authLog.info(f"Cells not found under file: {csvFile}")
        except FileNotFoundError:
            print("File not found. Please check the file path and try again.")
            authLog.error(f"File not found in path {csvFile}")
            authLog.error(traceback.format_exc())
            continue

        except Exception as error:
            print(f"ERROR: {error}\n", traceback.format_exc())
            authLog.error(f"Wasn't possible to choose the CSV file, error message: {error}\n", traceback.format_exc())

def chooseDocx(rowText):
    cityRegex = re.compile(r'city', re.IGNORECASE)
    stateRegex = re.compile(r'state', re.IGNORECASE)
    mplsRegex = re.compile(r'mpls-speed', re.IGNORECASE)
    siteCodeRegex = re.compile(r'site-code', re.IGNORECASE)

    while True:
        wordFile = input("Please enter the path to the DOCX file: ")
        try:
            wordDOC = Document(wordFile)
            hostname = input("Please input the Hostname of the device: ")
            siteCode = input(f"Please input the Site Code: ")
            city = input("Please input the City: ")
            state = input("Please input the State: ")
            mplsSpeed = input("Please input the MPLS speed: ")
            
            authLog.info(f"User chose  the DOCX File path: {wordFile}")
            print(f"INFO: file successfully found: {wordFile}.")
            replaceText = [
                'cedge1-serial-no',
                'cedge-device-ip',
                'cEdge1-host',
                'snmp-location',
                'sw-host - sw-cEdge1-port',
                'cEdge1-rtr-ip',
                'cEdge1-loop',
                'cEdge-asn',
                'cEdge1-loop',
                'cEdge1-loop',
                'cEdge1-sw-ip',
                'sw-host-sw-cEdge1-port',
                'switch-asn',
                'mpls-pe-ip',
                'cEdge2-tloc3-ext-ip',
                'cedge2-host - gi0/0/3 - TLOC3',
                'cEdge1-tloc3-ip',
                'sw-host - sw-cEdge1-mpls-port - LUM - mpls-circuitid',
                'mpls-ce1-ip',
                'cEdge1-host',
                'latitude',
                'longitude',
                'cEdge1-loop',
                'site-no'
            ]

            manualReplacements = {
                siteCodeRegex : siteCode,
                cityRegex : city,
                stateRegex : state,
                mplsRegex : mplsSpeed
            }

            for para in wordDOC.paragraphs:
                for run in para.runs:
                    if run.font.color.rgb == RGBColor(255, 0, 0):
                        print(f"Found red text: {run.text}")
                        for wordString, csvString in zip(replaceText, rowText):
                            if re.search(r'\b{}\b'.format(re.escape(wordString)), run.text, re.IGNORECASE):
                                print(f"INFO: Replacing '{wordString}' with '{csvString}'")
                                authLog.info(f"Replacing '{wordString}' with '{csvString}'")
                                run.text = re.sub(r'\b{}\b'.format(re.escape(wordString)), csvString, run.text, flags=re.IGNORECASE)

                        for placeholder, replacement in manualReplacements.items():
                            if placeholder.search(run.text):
                                print(f"Replacing '{placeholder.pattern}' with '{replacement}'")
                                authLog.info(f"Replacing '{placeholder.pattern}' with '{replacement}'")
                                run.text = placeholder.sub(replacement, run.text)

            for table in wordDOC.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                if run.font.color.rgb == RGBColor(255, 0, 0):
                                    print(f"Found red text: {run.text}")
                                    for wordString, csvString in zip(replaceText, rowText):
                                        if re.search(r'\b{}\b'.format(re.escape(wordString)), run.text, re.IGNORECASE):
                                            print(f"INFO: Replacing '{wordString}' with '{csvString}'")
                                            authLog.info(f"Replacing in Table: '{wordString}' with '{csvString}'")
                                            run.text = re.sub(r'\b{}\b'.format(re.escape(wordString)), csvString, run.text, flags=re.IGNORECASE)

                                    for placeholder, replacement in manualReplacements.items():
                                        if placeholder.search(run.text):
                                            print(f"Replacing '{placeholder.pattern}' with '{replacement}'")
                                            authLog.info(f"Replacing in Table: '{placeholder.pattern}' with '{replacement}'")
                                            run.text = placeholder.sub(replacement, run.text)
                                            
            newWordDoc = f"{hostname}.docx"
            wordDOC.save(newWordDoc)
            authLog.info(f"Replacements made successfully in DOCX file and saved as: {newWordDoc}")
            print(f"INFO: Replacements made successfully in DOCX file and saved as: {newWordDoc}")
            break

        except FileNotFoundError:
            print("File not found. Please check the file path and try again.")
            authLog.error(f"File not found in path {wordFile}")
            authLog.error(traceback.format_exc())
            continue

        except Exception as error:
            print(f"ERROR: {error}\n", traceback.format_exc())
            authLog.error(f"Wasn't possible to choose the DOCX file, error message: {error}\n", traceback.format_exc())
