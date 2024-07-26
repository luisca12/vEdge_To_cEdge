from log import authLog
from docx import Document
from docx.shared import RGBColor
from auth import Auth
from commandsCLI import shCoreInfo, shIntDesSDW

import re
import os
import csv
import json
import traceback
import ipaddress

removeCIDR_Patt = r'/\d{2}'

PID_SDW03 = 'C8300-1N1S-4T2X-'
PID_SDW04 = 'C8300-1N1S-4T2X-'

def chooseCSV():
    # ignoreStrings = re.compile(r'(FALSE|TRUE|100000|^100$|full|biz-internet|private5|TPX|core|^ge0\/0$|^ge0\/1$|^ge0\/1$|^ge0\/2$|^ge0\/3$)')
    csvDataList = []
    # ignoredStrMatchList = []

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
                        for row in rowText:
                            print(f"{row}")
                        csvDataList.append(rowText)                         
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
    for index, item in enumerate(mergedData):
        print(f"rowText[{index}] with string: {item}")
    return mergedData

def chooseDocx(rowText):
    print(f"This is rowText[12]: {rowText[12]}")
    swHostname, username, netDevice = Auth(rowText[12])
    shHostnameOut, netVlan1101, netVlan1103, shIntDesSDWOut, shIntDesCONOut1, shIntStatMPLSOut1, shVlanMgmtIP, shVlanMgmtCIDR, shLoop0Out = shCoreInfo(swHostname, username, netDevice)

    print(f"\n","="*76)
    print(f"INFO: Location: {rowText[3]}\n")

    print(f"INFO: TPX Circuit Information: {rowText[65]}\n")

    print(f"INFO: LUM Circuit Information {rowText[28]}")
    print(f"="*76, "\n")

    while True:
        try:
            wordFile = input("Please enter the path to the DOCX file: ")
            wordDOC = Document(wordFile)
            authLog.info(f"User chose  the DOCX File path: {wordFile}")
            print(f"INFO: file successfully found: {wordFile}.")
            serialNumSDW03 = input("Please input the serial number of SDW-03: ")
            serialNumSDW04 = input("Please input the serial number of SDW-04: ")
            cEdge1Loop = input("Please input the SDW03 Loopback IP Address: ")
            cEdge2Loop = input("Please input the SDW04 Loopback IP Address: ")
            siteNo = input(f"Please input the new Site ID (Old Site ID: {rowText[41]}):")
            city = input("Please input the City: ")
            state = input("Please input the State: ")
            siteCode = input(f"Please input the Site Code: ")
            mplsCircuitID = input("Please input the MPLS Circuit ID:")
            bb1Carrier = input("Please input the bb1-carrier: ")
            bb1Circuitid = input("Please input the bb1-circuitid: ")
            cEdge2TLOC3_Port = input(f"Please input the cedge2-tloc3-port (TenGigabitEthernet0/0/5 or GigabitEthernet0/0/1 for {bb1Carrier} port): ")
            print("=" * 61,"\n\tINFO: Now begins information of the Core Switch")
            print("=" * 61)
            print(f"{shHostnameOut}{shIntDesSDW}\n{shIntDesSDWOut}\n")
            swcEdge1_vlan = input("Please input the VLAN for SDW-03, 1101 if possible: ")
            swcEdge2_vlan = input("Please input the VLAN for SDW-04, 1103 if possible: ")
            swcEdge1_port = input("Please input the connection to SDW-03 gi0/0/0 in VPN 1 from the switch: ")
            swcEdge2_port = input("Please input the connection to SDW-04 gi0/0/0 in VPN 1 from the switch: ")
            swcEdge1_mplsPort = input("Please input the Switch port for SDW-03 gi0/0/2 connection to Lumen: ")
            swcEdge2_mplsPort = input("Please input the Switch port for SDW-04 gi0/0/2 connection to Lumen: ")

            print("\nrowText 2:", rowText[2], "rowText 17:", rowText[17])
            print("After changes:")
            rowText[2] = re.sub('01', '03', rowText[2])
            rowText[17] = re.sub('02', '04', rowText[17])
            rowText[17] = re.sub('ge0/3', 'ge0/0/3', rowText[17])
            print("rowText 2:", rowText[2], "rowText 17:", rowText[17])
            os.system("PAUSE")

            print("\nrowText 44:", rowText[44], "rowText 59:", rowText[59])
            print("After changes:")
            rowText[44] = re.sub('02', '04', rowText[44])
            rowText[59] = re.sub('01', '03', rowText[59])
            rowText[59] = re.sub('ge0/3', 'ge0/0/3', rowText[59])
            print("rowText 44:", rowText[44], "rowText 59:", rowText[59])
            os.system("PAUSE")

            print("rowText 6:", rowText[6], "rowText 18:", rowText[18], "rowText 29:", rowText[29], \
                  f"rowText 48:", rowText[48], "rowText 79:", rowText[79],"")
            print("After changes:")
            rowText[6] = re.sub(removeCIDR_Patt, '', rowText[6])
            rowText[18] = re.sub(removeCIDR_Patt, '', rowText[18])
            rowText[29] = re.sub(removeCIDR_Patt, '', rowText[29])
            rowText[48] = re.sub(removeCIDR_Patt, '', rowText[48])
            rowText[79] = re.sub(removeCIDR_Patt, '', rowText[79])
            print("rowText 6:", rowText[6], "rowText 18:", rowText[18], "rowText 29:", rowText[29], \
                  f"rowText 48:", rowText[48], "rowText 79:", rowText[79],"")
            os.system("PAUSE")

            cedge2TLOC3_List = rowText[66]
            cedge2TLOC3_STR = ''.join(cedge2TLOC3_List)
            cedge2TLOC3_IP_STR = cedge2TLOC3_STR.split('/')[0]
            cedge2TLOC3_CIDR_STR = cedge2TLOC3_STR.split('/')[1]
            cedge2TLOC3_MASK_STR = ipaddress.IPv4Network(cedge2TLOC3_STR, strict=False)
            cedge2TLOC3_MASK_STR = str(cedge2TLOC3_MASK_STR.netmask)

            serialNumSDW03 = PID_SDW03 + serialNumSDW03
            serialNumSDW04 = PID_SDW04 + serialNumSDW04

            replaceText = {
                'cedge1-host' : f'{rowText[2]}',
                'snmp-location' : f'{rowText[3]}',
                'cedge1-rtr-ip' : f'{rowText[6]}',
                'cEdge-asn' : f'{rowText[8]}',
                'cedge1-sw-ip' : f'{rowText[11]}',
                'switch-asn' : f'{rowText[13]}',
                'mpls-pe-ip' : f'{rowText[14]}',
                'cedge2-tloc3-ext-ip' : f'{rowText[15]}',
                'cedge2-host - gi0/0/3 - TLOC3' : f'{rowText[17]}',
                'cedge1-tloc3-ip'	: f'{rowText[18]}',
                'mpls-ce1-ip' : f'{rowText[29]}',
                'mpls-speed' : f'{rowText[35]}',
                'latitude' : f'{rowText[38]}',
                'longitude' : f'{rowText[39]}',
                # Here starts the second CSV file #
                'cedge2-host'	: f'{rowText[44]}',
                'bb1-down-speed' : f'{rowText[76]}',
                'cedge2-rtr-ip' : f'{rowText[48]}',
                'cedge2-sw-ip' : f'{rowText[53]}',	
                'cedge2-tloc3-gate' : f'{rowText[57]}',	
                'cedge1-host TLOC3 gi0/0/3' : f'{rowText[59]}',
                'cedge2-tloc3-ext-ip/30' : f'{rowText[60]}',
                'bb1-up-speed' : f'{rowText[75]}',	
                'mpls-ce2-ip'	: f'{rowText[79]}'
            }

            print(json.dumps(replaceText, indent=4))
            os.system("PAUSE")

            stringRegexPatt = {
                'cedge1-serial-no' : serialNumSDW03,
                'cedge2-serial-no' : serialNumSDW04,
                'cedge1-loop' : cEdge1Loop,
                'cedge2-loop' : cEdge2Loop,
                'site-no'	: siteNo,
                'city': city,
                'state': state,
                'site-code': siteCode,
                'sw-mgmt-ip' : shVlanMgmtIP,
                'sw-host' : f'{rowText[12]}',
                'sw-cEdge1-mpls-port': swcEdge1_mplsPort,
                'sw-cEdge2-mpls-port': swcEdge2_mplsPort,
                'mpls-circuitid':  mplsCircuitID,
                'bb1-carrier': bb1Carrier,
                'bb1-circuitid': bb1Circuitid,
                'cedge2-tloc3-port': cEdge2TLOC3_Port,
                'cedge2-tloc3-ip': cedge2TLOC3_IP_STR,
                'cedge2-tloc3-mask' : cedge2TLOC3_MASK_STR,
                'cedge2-tloc3-cidr': cedge2TLOC3_CIDR_STR,
                'cedge1-lan-net': netVlan1101,
                'cedge2-lan-net': netVlan1103,
                'sw-loop': shLoop0Out,
                'sw-mgmt-cidr': shVlanMgmtCIDR,
                'sw-cedge1-port': swcEdge1_port,
                'sw-cedge1-vlan': swcEdge1_vlan,
                'sw-cedge2-port': swcEdge2_port,
                'sw-cedge2-vlan': swcEdge2_vlan,
                'sw-mpls-port': shIntStatMPLSOut1[0],
                'sw-remote-con-net1': shIntDesCONOut1[0],
                'sw-remote-con-net2': shIntDesCONOut1[1],
                'sw-mgmt-vlan' : '1500'
            }

            manualReplacements = {re.compile(r'\b{}\b'.format(pattern), re.IGNORECASE): value for pattern, value in stringRegexPatt.items()}

            for para in wordDOC.paragraphs:
                if any(run.font.color.rgb == RGBColor(255, 0, 0) for run in para.runs):
                    print(f"Found red text: {para.text}")
                    for wordString, csvString in replaceText.items():
                        if re.search(r'\b{}\b'.format(re.escape(wordString)), para.text, re.IGNORECASE):
                            print(f"INFO: Replacing '{wordString}' with '{csvString}'")
                            authLog.info(f"Replacing '{wordString}' with '{csvString}'")
                            para.text = re.sub(r'\b{}\b'.format(re.escape(wordString)), csvString, para.text, flags=re.IGNORECASE)

                    for placeholder, replacement in manualReplacements.items():
                        replacement = str(replacement)
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
                                for wordString, csvString in replaceText.items():
                                    if re.search(r'\b{}\b'.format(re.escape(wordString)), paragraph.text, re.IGNORECASE):
                                        print(f"INFO: Replacing '{wordString}' with '{csvString}'")
                                        authLog.info(f"Replacing in Table: '{wordString}' with '{csvString}'")
                                        paragraph.text = re.sub(r'\b{}\b'.format(re.escape(wordString)), csvString, paragraph.text, flags=re.IGNORECASE)

                                for placeholder, replacement in manualReplacements.items():
                                    replacement = str(replacement)
                                    if placeholder.search(paragraph.text):
                                        print(f"Replacing '{placeholder.pattern}' with '{replacement}'")
                                        authLog.info(f"Replacing in Table: '{placeholder.pattern}' with '{replacement}'")
                                        paragraph.text = placeholder.sub(replacement, paragraph.text)

            newWordDoc = f"{siteCode}_ImplementationPlan.docx"
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
            authLog.error(f"Wasn't possible to choose the DOCX file, error message: {error}\n{traceback.format_exc()}")