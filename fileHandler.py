from log import authLog
from docx import Document
from docx.shared import RGBColor

import re
import os
import csv
import traceback

def chooseCSV():
    ignoreStrings = re.compile(r'(FALSE|TRUE|100000|^100$|full|biz-internet|private5|TPX|core|^ge0\/0$|^ge0\/1$|^ge0\/1$|^ge0\/2$|^ge0\/3$)')
    csvDataList = []
    ignoredStrMatchList = []

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
                        ignoredStrMatch = [row for row in rowText if ignoreStrings.search(row)]
                        print("Found the following strings in the CSV file:")
                        for row in filteredRowText:
                            print(f"{row}")
                        authLog.info(f"Found the following strings in the CSV file:\n{filteredRowText}")
                        ignoredStrMatchList.append(ignoredStrMatch)
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
    mergedData1 = [item for sublist in ignoredStrMatchList for item in sublist]
    return mergedData, mergedData1

def chooseDocx(rowText, ignoredStrings=""):
    matchTPXRegex = re.compile(r'TPX')
    matchTPXRegexOut = [row for row in ignoredStrings if matchTPXRegex.search(row)]
    print(f"\nINFO: TPX Circuit Information: {matchTPXRegexOut[0]}")

    matchLUMRegex = re.compile(r'LUM')
    matchLUMRegexOut = [row for row in ignoredStrings if matchLUMRegex.search(row)]
    print(f"INFO: LUM Circuit Information for SDW-01: {matchLUMRegexOut[0]}")
    print(f"INFO: LUM Circuit Information for SDW-02: {matchLUMRegexOut[1]} \n")

    while True:
        wordFile = input("Please enter the path to the DOCX file: ")
        try:
            wordDOC = Document(wordFile)
            hostname = input("Please input the Hostname of the device: ")
            siteCode = input(f"Please input the Site Code: ")
            city = input("Please input the City: ")
            state = input("Please input the State: ")
            swMgmtIP = input("Please input the Switch Core Management IP: ")
            swHost = input("Please input the Switch Core Hostname: ")
            swcEdge1_mplsPort = input("Please enter the Switch port for cEdge1 connection to Lumen (sw-cEdge1-mpls-port): ")
            swcEdge2_mplsPort = input("Please enter the Switch port for cEdge1 connection to Lumen (sw-cEdge2-mpls-port): ")
            mplsCircuitID = input("Please input the MPLS Circuit ID:")
            mplsSpeed = input("Please input the MPLS speed: ")
            bb1Carrier = input("Please input the bb1-carrier: ")
            bb1Circuitid = input("Please input the bb1-circuitid: ")
            bb1UPspeed =  input("Please input the bb1 Upload speed: ")
            bb1DWspeed =  input("Please input the bb1 Download speed: ")
            cEdge2TLOC3_Port = input("Please input the cedge2-tloc3-port: ")

            authLog.info(f"User chose  the DOCX File path: {wordFile}")
            print(f"INFO: file successfully found: {wordFile}.")
            print(rowText)
            os.system("PAUSE")
            replaceText = {
                'cedge1-serial-no': f'{rowText[0]}',
                'cedge1-device-ip': f'{rowText[1]}',
                'cEdge1-host': f'{rowText[2]}',
                'snmp-location': f'{rowText[3]}',
                'cEdge1-rtr-ip': f'{rowText[4]}',
                'cEdge1-loop': f'{rowText[7]}', #Changed to rowText[7] since we only need the IP, no prefix-length
                'cEdge-asn': f'{rowText[6]}',
                'cEdge1-loop': f'{rowText[7]}', #Changed to rowText[7] since we only need the IP, no prefix-length
                'cEdge1-loop' : f'{rowText[7]}', #Changed to rowText[7] since we only need the IP, no prefix-length
                'cEdge1-sw-ip': f'{rowText[9]}',
                'switch-asn': f'{rowText[10]}',
                'mpls-pe-ip': f'{rowText[11]}',
                'cEdge2-tloc3-ext-ip': f'{rowText[12]}',
                'cedge2-host - gi0/0/3 - TLOC3': f'{rowText[13]}',
                'cEdge1-tloc3-ip': f'{rowText[14]}',
                'mpls-ce1-ip': f'{rowText[15]}',
                'cEdge1-host': f'{rowText[16]}',
                'latitude': f'{rowText[17]}',
                'longitude': f'{rowText[18]}',
                'cEdge1-loop': f'{rowText[7]}', #Changed to rowText[7] since we only need the IP, no prefix-length
                'site-no': f'{rowText[20]}',
                # Here starts the second CSV File
                'cedge2-serial-no': f'{rowText[21]}',
                'cedge2-device-ip': f'{rowText[22]}',
                'cEdge2-host': f'{rowText[23]}',
                'snmp-location': f'{rowText[24]}',
                'cEdge2-rtr-ip': f'{rowText[25]}',
                'cEdge2-loop': f'{rowText[28]}', # Changed to rowText[28] since we only need the IP, no prefix-length
                'cEdge-asn': f'{rowText[27]}',
                'cEdge2-loop': f'{rowText[28]}', # Changed to rowText[28] since we only need the IP, no prefix-length
                'cEdge2-loop': f'{rowText[28]}', # Changed to rowText[28] since we only need the IP, no prefix-length
                'cEdge2-sw-ip': f'{rowText[30]}',
                'switch-asn': f'{rowText[31]}',
                'cEdge2-tloc3-gate': f'{rowText[32]}',
                'mpls-pe-ip': f'{rowText[33]}',
                'cEdge1-host TLOC3 gi0/0/3': f'{rowText[34]}',
                'cEdge2-tloc3-ext-ip': f'{rowText[35]}',
                'cedge2-tloc3-ip/cedge2-tloc3-cidr': f'{rowText[36]}',
                'mpls-ce2-ip': f'{rowText[37]}',
                'cEdge2-host': f'{rowText[38]}',
                'latitude': f'{rowText[39]}',
                'longitude': f'{rowText[40]}',
                'cEdge2-loop': f'{rowText[28]}', # Changed to rowText[28] since we only need the IP, no prefix-length
                'site-no': f'{rowText[42]}'
            }

            stringRegexPatt = {
                'city': city,
                'state': state,
                'mpls-speed': mplsSpeed,
                'site-code': siteCode,
                'sw-mgmt-ip' : swMgmtIP,
                'sw-host' : swHost,
                'sw-cEdge1-mpls-port': swcEdge1_mplsPort,
                'sw-cEdge2-mpls-port': swcEdge2_mplsPort,
                'mpls-circuitid':  mplsCircuitID,
                'bb1-carrier': bb1Carrier,
                'bb1-circuitid': bb1Circuitid,
                'bb1-up-speed': bb1UPspeed,
                'bb1-down-speed': bb1DWspeed,
                'cedge2-tloc3-port': cEdge2TLOC3_Port
            }

            manualReplacements = {re.compile(r'\b{}\b'.format(pattern), re.IGNORECASE): value for pattern, value in stringRegexPatt.items()}

            for para in wordDOC.paragraphs:
                if any(run.font.color.rgb == RGBColor(255, 0, 0) for run in para.runs):
                    print(f"Found red text: {para.text}")
                    for wordString, csvString in zip(replaceText, rowText):
                        if re.search(r'\b{}\b'.format(re.escape(wordString)), para.text, re.IGNORECASE):
                            print(f"INFO: Replacing '{wordString}' with '{csvString}'")
                            authLog.info(f"Replacing '{wordString}' with '{csvString}'")
                            para.text = re.sub(r'\b{}\b'.format(re.escape(wordString)), csvString, para.text, flags=re.IGNORECASE)

                    for placeholder, replacement in manualReplacements.items():
                        if placeholder.search(para.text):
                            print(f"Replacing '{placeholder.pattern}' with '{replacement}'")
                            authLog.info(f"Replacing '{placeholder.pattern}' with '{replacement}'")
                            para.text = placeholder.sub(replacement, para.text)

            for table in wordDOC.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            if any(run.font.color.rgb == RGBColor(255, 0, 0) for run in paragraph.runs):
                                print(f"Found red text: {paragraph.text}")
                                for wordString, csvString in zip(replaceText, rowText):
                                    if re.search(r'\b{}\b'.format(re.escape(wordString)), paragraph.text, re.IGNORECASE):
                                        print(f"INFO: Replacing '{wordString}' with '{csvString}'")
                                        authLog.info(f"Replacing in Table: '{wordString}' with '{csvString}'")
                                        paragraph.text = re.sub(r'\b{}\b'.format(re.escape(wordString)), csvString, paragraph.text, flags=re.IGNORECASE)

                                for placeholder, replacement in manualReplacements.items():
                                    if placeholder.search(paragraph.text):
                                        print(f"Replacing '{placeholder.pattern}' with '{replacement}'")
                                        authLog.info(f"Replacing in Table: '{placeholder.pattern}' with '{replacement}'")
                                        paragraph.text = placeholder.sub(replacement, paragraph.text)

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
