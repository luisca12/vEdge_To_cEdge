from log import authLog
from docx import Document
from docx.shared import RGBColor

import re
import os
import csv
import traceback

def chooseCSV():
    ignoreStrings = re.compile(r'(FALSE|TRUE|100000|biz-internet|private5|^ge0\/0$|^ge0\/1$|^ge0\/1$|^ge0\/2$|^ge0\/3$)')
    csvDataList = []

    for i in range(2):
        while True:
            csvFile = input(f"Please enter the path to the CSV file {i + 1}: ")
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
                        csvDataList.append(filteredRowText)
                        break
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
                
    mergedData = [item for sublist in csvDataList for item in sublist]
    return mergedData

def chooseDocx(rowText):
    # cityRegex = re.compile(r'city', re.IGNORECASE)
    # stateRegex = re.compile(r'state', re.IGNORECASE)
    # mplsRegex = re.compile(r'mpls-speed', re.IGNORECASE)
    # siteCodeRegex = re.compile(r'site-code', re.IGNORECASE)

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
            replaceText = {
                'cedge1-serial-no':rowText[0],
                'cedge-device-ip':rowText[1],
                'cEdge1-host':rowText[2],
                'snmp-location':rowText[3],
                'sw-host - sw-cEdge1-port':rowText[4],
                'cEdge1-rtr-ip':rowText[5],
                'cEdge1-loop':rowText[6],
                'cEdge-asn':rowText[7],
                'cEdge1-loop':rowText[8],
                'cEdge1-loop':rowText[9],
                'cEdge1-sw-ip':rowText[10],
                'sw-host-sw-cEdge1-port':rowText[11],
                'switch-asn':rowText[12],
                'mpls-pe-ip':rowText[13],
                'cEdge2-tloc3-ext-ip':rowText[14],
                'cedge2-host - gi0/0/3 - TLOC3':rowText[15],
                'cEdge1-tloc3-ip':rowText[16],
                'sw-host - sw-cEdge1-mpls-port - LUM - mpls-circuitid':rowText[17],
                'mpls-ce1-ip':rowText[18],
                'cEdge1-host':rowText[19],
                'latitude':rowText[20],
                'longitude':rowText[21],
                'cEdge1-loop':rowText[22],
                'site-no':rowText[23],
                # Here starts the second CSV File
                'cedge2-serial-no': rowText[24],
                'cedge2-device-ip': rowText[25],
                'cEdge2-host': rowText[26],
                'snmp-location': rowText[27],
                'cEdge2-rtr-ip': rowText[28],
                '':''
            }

            stringRegexPatt = {
                'city': city,
                'state': state,
                'mpls-speed': mplsSpeed,
                'site-code': siteCode
            }

            manualReplacements = {re.compile(r'\b{}\b'.format(pattern), re.IGNORECASE): value for pattern, value in stringRegexPatt.items()}

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
