from log import authLog
from docx import Document
from docx.shared import RGBColor
from auth import Auth
from commandsCLI import shCoreInfo, shIntDesSDW
from openpyxl import Workbook

import re
import os
import csv
import json
import traceback
import ipaddress
import openpyxl

removeCIDR_Patt = r'/\d{2}'

PID_SDW03 = 'C8300-1N1S-4T2X-'
PID_SDW04 = 'C8300-1N1S-4T2X-'

ndlmPath1 = "NDLM_Template.xlsx"
ndlmPath2 = "NDLM_Tier2_Template.xlsx"

sdw03Template = "sdw-03-template.csv"
sdw04Template = "sdw-04-template.csv"

returnList = []

def chooseCSV():
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
    os.system("PAUSE")
    return mergedData

def chooseDocx_ISR(rowText):
    print(f"This is rowText[12]: {rowText[12]}")
    swHostname, username, netDevice = Auth(rowText[12])
    shHostnameOut, netVlan1101, netVlan1103, shIntDesSDWOut, shIntDesCONOut1, shIntStatMPLSOut1, shVlanMgmtIP, shVlanMgmtCIDR, shLoop0Out = shCoreInfo(swHostname, username, netDevice)

    print(f"\n","="*76)
    print(f"INFO: Location: {rowText[3]}\n")

    print(f"INFO: BB1 Circuit Information: {rowText[65]}\n")

    print(f"INFO: MPLS Circuit Information {rowText[28]}")
    print(f"="*76, "\n")

    while True:
        try: 
            wordFile = "Caremore - Tier II - 8300 - vEdge to cEdge - gold.docx"
            wordDOC = Document(wordFile)
            authLog.info(f"User chose  the DOCX File path: {wordFile}")
            print(f"INFO: file successfully found: {wordFile}.")
            serialNumSDW01 = input("Please input the serial number of SDW-01: ")
            serialNumSDW02 = input("Please input the serial number of SDW-02: ")
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

            serialNumSDW03New = PID_SDW03 + serialNumSDW03
            serialNumSDW04New = PID_SDW04 + serialNumSDW04

            snmpLocation = f'{rowText[3]}'
            cedge1_host = f'{rowText[2]}'
            cedge2_host = f'{rowText[44]}'

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
                'cedge1-serial-no' : serialNumSDW03New,
                'cedge2-serial-no' : serialNumSDW04New,
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

            return {
                'site-code': siteCode,
                'serialNumSDW01': serialNumSDW01,
                'serialNumSDW02': serialNumSDW02,
                'serialNumSDW03': serialNumSDW03,
                'serialNumSDW04': serialNumSDW04,
                'cedge1-loop': cEdge1Loop,
                'cedge2-loop': cEdge2Loop,
                'snmp-location': snmpLocation,
                'city': city,
                'state': state,
                'site-no': siteNo,
                'cedge1-host': cedge1_host,
                'cedge2-host': cedge2_host,
                'sw-host' : f'{rowText[12]}',
                'sw-mpls-port' : shIntStatMPLSOut1[0],
                'cedge2-tloc3-port': cEdge2TLOC3_Port,
                'sw-cedge1-port' : swcEdge1_port,
                'sw-cedge2-port' : swcEdge2_port,
                'sw-cedge1-mpls-port' : swcEdge1_mplsPort,
                'sw-cedge2-mpls-port' : swcEdge2_mplsPort,
                'vedge1-loop' : f'{rowText[9]}',
                'vedge2-loop' : f'{rowText[51]}',

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
                'bb1-down-speed' : f'{rowText[76]}',
                'cedge2-rtr-ip' : f'{rowText[48]}',
                'cedge2-sw-ip' : f'{rowText[53]}',	
                'cedge2-tloc3-gate' : f'{rowText[57]}',	
                'cedge1-host TLOC3 gi0/0/3' : f'{rowText[59]}',
                'cedge2-tloc3-ext-ip/30' : f'{rowText[60]}',
                'bb1-up-speed' : f'{rowText[75]}',	
                'mpls-ce2-ip'	: f'{rowText[79]}',

                'cedge1-serial-no' : serialNumSDW03New,
                'cedge2-serial-no' : serialNumSDW04New,
                'sw-mgmt-ip' : shVlanMgmtIP,
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
                'sw-cedge1-vlan': swcEdge1_vlan,
                'sw-cedge2-vlan': swcEdge2_vlan,
                'sw-mpls-port': shIntStatMPLSOut1[0],
                'sw-remote-con-net1': shIntDesCONOut1[0],
                'sw-remote-con-net2': shIntDesCONOut1[1],
                'sw-mgmt-vlan' : '1500'
            }

        except FileNotFoundError:
            print("File not found. Please check the file path and try again.")
            authLog.error(f"File not found in path {wordFile}")
            authLog.error(traceback.format_exc())
            continue

        except Exception as error:
            print(f"ERROR: {error}\n", traceback.format_exc())
            authLog.error(f"Wasn't possible to choose the DOCX file, error message: {error}\n{traceback.format_exc()}")

def chooseDocx_vEdge(rowText):
    print(f"This is rowText[13], switch hostname: {rowText[13]}")
    swHostname, username, netDevice = Auth(rowText[13])
    shHostnameOut, netVlan1101, netVlan1103, shIntDesSDWOut, shIntDesCONOut1, shIntStatMPLSOut1, shVlanMgmtIP, shVlanMgmtCIDR, shLoop0Out = shCoreInfo(swHostname, username, netDevice)

    print(f"\n","="*76)
    print(f"INFO: Location: {rowText[3]}\n")

    print(f"INFO: BB1 Circuit Information: {rowText[71]}\n")

    print(f"INFO: MPLS Circuit Information {rowText[31]}")
    print(f"="*76, "\n")

    while True:
        try:
            wordFile = "Caremore - Tier II - 8300 - vEdge to cEdge - gold.docx"
            wordDOC = Document(wordFile)
            authLog.info(f"User chose  the DOCX File path: {wordFile}")
            print(f"INFO: file successfully found: {wordFile}.")
            serialNumSDW01 = input("Please input the serial number of SDW-01: ")
            serialNumSDW02 = input("Please input the serial number of SDW-02: ")
            serialNumSDW03 = input("Please input the serial number of SDW-03: ")
            serialNumSDW04 = input("Please input the serial number of SDW-04: ")
            cEdge1Loop = input("Please input the SDW03 Loopback IP Address: ")
            cEdge2Loop = input("Please input the SDW04 Loopback IP Address: ")
            siteNo = input(f"Please input the new Site ID (Old Site ID: {rowText[44]}):")
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

            print("\nrowText 2:", rowText[2], "rowText 20:", rowText[20])
            print("After changes:")
            rowText[2] = re.sub('01', '03', rowText[2])
            rowText[20] = re.sub('02', '04', rowText[20])
            rowText[20] = re.sub('ge0/3', 'ge0/0/3', rowText[20])
            print("rowText 2:", rowText[2], "rowText 20:", rowText[20])
            os.system("PAUSE")

            print("\nrowText 47:", rowText[47], "rowText 65:", rowText[65])
            print("After changes:")
            rowText[47] = re.sub('02', '04', rowText[47])
            rowText[65] = re.sub('01', '03', rowText[65])
            rowText[65] = re.sub('ge0/3', 'ge0/0/3', rowText[65])
            print("rowText 47:", rowText[47], "rowText 65:", rowText[65])
            os.system("PAUSE")

            print("rowText 6:", rowText[6], "rowText 21:", rowText[21], "rowText 32:", rowText[32], \
                  f"rowText 51:", rowText[51], "rowText 85:", rowText[85],"")
            print("After changes:")
            rowText[6] = re.sub(removeCIDR_Patt, '', rowText[6])
            rowText[21] = re.sub(removeCIDR_Patt, '', rowText[21])
            rowText[32] = re.sub(removeCIDR_Patt, '', rowText[32])
            rowText[51] = re.sub(removeCIDR_Patt, '', rowText[51])
            rowText[85] = re.sub(removeCIDR_Patt, '', rowText[85])
            print("rowText 6:", rowText[6], "rowText 21:", rowText[21], "rowText 32:", rowText[32], \
                  f"rowText 51:", rowText[51], "rowText 85:", rowText[85],"")
            os.system("PAUSE")

            cedge2TLOC3_List = rowText[72]
            cedge2TLOC3_STR = ''.join(cedge2TLOC3_List)
            cedge2TLOC3_IP_STR = cedge2TLOC3_STR.split('/')[0]
            cedge2TLOC3_CIDR_STR = cedge2TLOC3_STR.split('/')[1]
            cedge2TLOC3_MASK_STR = ipaddress.IPv4Network(cedge2TLOC3_STR, strict=False)
            cedge2TLOC3_MASK_STR = str(cedge2TLOC3_MASK_STR.netmask)

            serialNumSDW03New = PID_SDW03 + serialNumSDW03
            serialNumSDW04New = PID_SDW04 + serialNumSDW04

            snmpLocation = f'{rowText[3]}'
            cedge1_host = f'{rowText[2]}'
            cedge2_host = f'{rowText[47]}'

            replaceText = {
                'cedge1-host' : f'{rowText[2]}',
                'snmp-location' : f'{rowText[3]}',
                'cedge1-rtr-ip' : f'{rowText[6]}',
                'cEdge-asn' : f'{rowText[9]}',
                'cedge1-sw-ip' : f'{rowText[12]}',
                'switch-asn' : f'{rowText[14]}',
                'mpls-pe-ip' : f'{rowText[17]}',
                'cedge2-tloc3-ext-ip' : f'{rowText[18]}',
                'cedge2-host - gi0/0/3 - TLOC3' : f'{rowText[20]}',
                'cedge1-tloc3-ip'	: f'{rowText[21]}',
                'mpls-ce1-ip' : f'{rowText[32]}',
                'mpls-speed' : f'{rowText[38]}',
                'latitude' : f'{rowText[41]}',
                'longitude' : f'{rowText[42]}',
                # Here starts the second CSV file #
                'cedge2-host'	: f'{rowText[47]}',
                'bb1-down-speed' : f'{rowText[82]}',
                'cedge2-rtr-ip' : f'{rowText[51]}',
                'cedge2-sw-ip' : f'{rowText[57]}',	
                'cedge2-tloc3-gate' : f'{rowText[63]}',	
                'cedge1-host TLOC3 gi0/0/3' : f'{rowText[59]}',
                'cedge2-tloc3-ext-ip/30' : f'{rowText[60]}',
                'bb1-up-speed' : f'{rowText[82]}',	
                'mpls-ce2-ip'	: f'{rowText[85]}'
            }

            print(json.dumps(replaceText, indent=4))
            os.system("PAUSE")

            stringRegexPatt = {
                'cedge1-serial-no' : serialNumSDW03New,
                'cedge2-serial-no' : serialNumSDW04New,
                'cedge1-loop' : cEdge1Loop,
                'cedge2-loop' : cEdge2Loop,
                'site-no'	: siteNo,
                'city': city,
                'state': state,
                'site-code': siteCode,
                'sw-mgmt-ip' : shVlanMgmtIP,
                'sw-host' : f'{rowText[13]}',
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

            returnList.append(rowText)

            for index, item in enumerate(rowText):
                print(f"returnList[{index}] with string: {item}")

            print("\n")

            for index, item in enumerate(returnList):
                print(f"returnList[{index}] with string: {item}")
            
            os.system("PAUSE")

            return [
                returnList,
                serialNumSDW01,
                serialNumSDW02,
                serialNumSDW03,
                serialNumSDW04,
                cEdge1Loop,
                cEdge2Loop,
                siteNo,
                city,
                state,
                siteCode,
                mplsCircuitID,
                bb1Carrier,
                bb1Circuitid,
                cEdge2TLOC3_Port,
                swcEdge1_vlan,
                swcEdge2_vlan,
                swcEdge1_port,
                swcEdge2_port,
                swcEdge1_mplsPort,
                swcEdge2_mplsPort
            ]

            # return {
            #     'cedge1-host' : f'{rowText[2]}',
            #     'snmp-location' : f'{rowText[3]}',
            #     'cedge1-rtr-ip' : f'{rowText[6]}',
            #     'cEdge-asn' : f'{rowText[9]}',
            #     'cedge1-sw-ip' : f'{rowText[12]}',
            #     'switch-asn' : f'{rowText[14]}',
            #     'mpls-pe-ip' : f'{rowText[17]}',
            #     'cedge2-tloc3-ext-ip' : f'{rowText[18]}',
            #     'cedge2-host - gi0/0/3 - TLOC3' : f'{rowText[20]}',
            #     'cedge1-tloc3-ip'	: f'{rowText[21]}',
            #     'mpls-ce1-ip' : f'{rowText[32]}',
            #     'mpls-speed' : f'{rowText[38]}',
            #     'latitude' : f'{rowText[41]}',
            #     'longitude' : f'{rowText[42]}',
            #     # Here starts the second CSV file #
            #     'cedge2-host'	: f'{rowText[47]}',
            #     'bb1-down-speed' : f'{rowText[82]}',
            #     'cedge2-rtr-ip' : f'{rowText[51]}',
            #     'cedge2-sw-ip' : f'{rowText[57]}',	
            #     'cedge2-tloc3-gate' : f'{rowText[63]}',	
            #     'cedge1-host TLOC3 gi0/0/3' : f'{rowText[59]}',
            #     'cedge2-tloc3-ext-ip/30' : f'{rowText[60]}',
            #     'bb1-up-speed' : f'{rowText[82]}',	
            #     'mpls-ce2-ip'	: f'{rowText[85]}',

            #     'cedge1-serial-no' : serialNumSDW03New,
            #     'cedge2-serial-no' : serialNumSDW04New,
            #     'cedge1-loop' : cEdge1Loop,
            #     'cedge2-loop' : cEdge2Loop,
            #     'site-no'	: siteNo,
            #     'city': city,
            #     'state': state,
            #     'site-code': siteCode,
            #     'sw-mgmt-ip' : shVlanMgmtIP,
            #     'sw-host' : f'{rowText[13]}',
            #     'sw-cEdge1-mpls-port': swcEdge1_mplsPort,
            #     'sw-cEdge2-mpls-port': swcEdge2_mplsPort,
            #     'mpls-circuitid':  mplsCircuitID,
            #     'bb1-carrier': bb1Carrier,
            #     'bb1-circuitid': bb1Circuitid,
            #     'cedge2-tloc3-port': cEdge2TLOC3_Port,
            #     'cedge2-tloc3-ip': cedge2TLOC3_IP_STR,
            #     'cedge2-tloc3-mask' : cedge2TLOC3_MASK_STR,
            #     'cedge2-tloc3-cidr': cedge2TLOC3_CIDR_STR,
            #     'cedge1-lan-net': netVlan1101,
            #     'cedge2-lan-net': netVlan1103,
            #     'sw-loop': shLoop0Out,
            #     'sw-mgmt-cidr': shVlanMgmtCIDR,
            #     'sw-cedge1-port': swcEdge1_port,
            #     'sw-cedge1-vlan': swcEdge1_vlan,
            #     'sw-cedge2-port': swcEdge2_port,
            #     'sw-cedge2-vlan': swcEdge2_vlan,
            #     'sw-mpls-port': shIntStatMPLSOut1[0],
            #     'sw-remote-con-net1': shIntDesCONOut1[0],
            #     'sw-remote-con-net2': shIntDesCONOut1[1],
            #     'sw-mgmt-vlan' : '1500'


                # 'site-code': siteCode,
                # 'serialNumSDW01': serialNumSDW01,
                # 'serialNumSDW02': serialNumSDW02,
                # 'serialNumSDW03': serialNumSDW03,
                # 'serialNumSDW04': serialNumSDW04,
                # 'cedge1-loop': cEdge1Loop,
                # 'cedge2-loop': cEdge2Loop,
                # 'snmp-location': snmpLocation,
                # 'city': city,
                # 'state': state,
                # 'site-no': siteNo,
                # 'cedge1-host': cedge1_host,
                # 'cedge2-host': cedge2_host,
                # 'sw-host' : f'{rowText[12]}',
                # 'sw-mpls-port' : shIntStatMPLSOut1[0],
                # 'cedge2-tloc3-port': cEdge2TLOC3_Port,
                # 'sw-cedge1-port' : swcEdge1_port,
                # 'sw-cedge2-port' : swcEdge2_port,
                # 'sw-cedge1-mpls-port' : swcEdge1_mplsPort,
                # 'sw-cedge2-mpls-port' : swcEdge2_mplsPort,
                # 'vedge1-loop' : f'{rowText[10]}',
                # 'vedge2-loop' : f'{rowText[55]}',

                # 'cedge1-rtr-ip' : f'{rowText[6]}',
                # 'cEdge-asn' : f'{rowText[8]}',
                # 'cedge1-sw-ip' : f'{rowText[11]}',
                # 'switch-asn' : f'{rowText[13]}',
                # 'mpls-pe-ip' : f'{rowText[14]}',
                # 'cedge2-tloc3-ext-ip' : f'{rowText[15]}',
                # 'cedge2-host - gi0/0/3 - TLOC3' : f'{rowText[17]}',
                # 'cedge1-tloc3-ip'	: f'{rowText[18]}',
                # 'mpls-ce1-ip' : f'{rowText[29]}',
                # 'mpls-speed' : f'{rowText[35]}',
                # 'latitude' : f'{rowText[38]}',
                # 'longitude' : f'{rowText[39]}',
                # # Here starts the second CSV file #
                # 'bb1-down-speed' : f'{rowText[76]}',
                # 'cedge2-rtr-ip' : f'{rowText[51]}',
                # 'cedge2-sw-ip' : f'{rowText[57]}',	
                # 'cedge2-tloc3-gate' : f'{rowText[63]}',	
                # 'cedge1-host TLOC3 gi0/0/3' : f'{rowText[59]}',
                # 'cedge2-tloc3-ext-ip/30' : f'{rowText[60]}',
                # 'bb1-up-speed' : f'{rowText[82]}',	
                # 'mpls-ce2-ip'	: f'{rowText[85]}',

                # 'cedge1-serial-no' : serialNumSDW03New,
                # 'cedge2-serial-no' : serialNumSDW04New,
                # 'sw-mgmt-ip' : shVlanMgmtIP,
                # 'mpls-circuitid':  mplsCircuitID,
                # 'bb1-carrier': bb1Carrier,
                # 'bb1-circuitid': bb1Circuitid,
                # 'cedge2-tloc3-port': cEdge2TLOC3_Port,
                # 'cedge2-tloc3-ip': cedge2TLOC3_IP_STR,
                # 'cedge2-tloc3-mask' : cedge2TLOC3_MASK_STR,
                # 'cedge2-tloc3-cidr': cedge2TLOC3_CIDR_STR,
                # 'cedge1-lan-net': netVlan1101,
                # 'cedge2-lan-net': netVlan1103,
                # 'sw-loop': shLoop0Out,
                # 'sw-mgmt-cidr': shVlanMgmtCIDR,
                # 'sw-cedge1-vlan': swcEdge1_vlan,
                # 'sw-cedge2-vlan': swcEdge2_vlan,
                # 'sw-mpls-port': shIntStatMPLSOut1[0],
                # 'sw-remote-con-net1': shIntDesCONOut1[0],
                # 'sw-remote-con-net2': shIntDesCONOut1[1],
                # 'sw-mgmt-vlan' : '1500'
            # }
        
        except FileNotFoundError:
            print("File not found. Please check the file path and try again.")
            authLog.error(f"File not found in path {wordFile}")
            authLog.error(traceback.format_exc())
            continue

        except Exception as error:
            print(f"ERROR: {error}\n", traceback.format_exc())
            authLog.error(f"Wasn't possible to choose the DOCX file, error message: {error}\n{traceback.format_exc()}")

def modNDLM(siteCode, serialNumSDW01, serialNumSDW02, serialNumSDW03, serialNumSDW04, cEdge1Loop, cEdge2Loop, snmpLocation, vEdge1Loop, vEdge2Loop):
    try:
        replaceText = {
            'site-code' : f'{siteCode}',
            'vedge1-serial-no' : f'{serialNumSDW01}',
            'vedge2-serial-no' : f'{serialNumSDW02}',
            'cedge1-serial-no' : f'{serialNumSDW03}',
            'cedge2-serial-no' : f'{serialNumSDW04}',
            'cedge1-loop' : f'{cEdge1Loop}',
            'cedge2-loop' : f'{cEdge2Loop}',
            'snmp-location' : f'{snmpLocation}',
            'vedge1-loop': f'{vEdge1Loop}',
            'vedge2-loop': f'{vEdge2Loop}'
        }

        ndlmFile = openpyxl.load_workbook(ndlmPath1)
        ndlmFileSheet = ndlmFile.active

        for row in ndlmFileSheet.iter_rows():
            for cell in row:
                if cell.value:
                    cellValue = str(cell.value).strip()
                    for key, value in replaceText.items():
                        if key.lower() in cellValue.lower():
                            cellValue = cellValue.replace(key, value)
                    cell.value = cellValue

            newNDLMFile = f'{siteCode}-NDLM.xlsx'
            ndlmFile.save(newNDLMFile)

    except FileNotFoundError:
        print("File not found. Please check the file path and try again.")
        authLog.error(f"File not found in path {ndlmPath1}")
        authLog.error(traceback.format_exc())

    except Exception as error:
        print(f"ERROR: {error}\n", traceback.format_exc())
        authLog.error(f"Wasn't possible to choose the CSV file, error message: {error}\n", traceback.format_exc())

def modNDLM2(siteCode, cEdge1Loop, cEdge2Loop, snmpLocation, city, state, siteNo,
             cedge1_host, cedge2_host, sw_host, sw_mpls_port, cEdge2TLOC3_Port,
             swcEdge1_port, swcEdge2_port, swcEdge1_mplsPort, swcEdge2_mplsPort
            ):
    try:

        replaceText = {
            'site-code' : f'{siteCode}',
            'cedge1-loop' : f'{cEdge1Loop}',
            'cedge2-loop' : f'{cEdge2Loop}',
            'snmp-location' : f'{snmpLocation}',
            'city': city,
            'state': state,
            'site-no': siteNo,
            'cedge1-host': cedge1_host,
            'cedge2-host': cedge2_host,
            'sw-host' : sw_host,
            'sw-mpls-port' : sw_mpls_port,
            'cedge2-tloc3-port': cEdge2TLOC3_Port,
            'sw-cedge1-port' : swcEdge1_port,
            'sw-cedge2-port' : swcEdge2_port,
            'sw-cedge1-mpls-port' : swcEdge1_mplsPort,
            'sw-cedge2-mpls-port' : swcEdge2_mplsPort
        }

        ndlmFile = openpyxl.load_workbook(ndlmPath2)
        ndlmFileSheet = ndlmFile.active

        for row in ndlmFileSheet.iter_rows():
            for cell in row:
                if cell.value:
                    cellValue = str(cell.value).strip()
                    for key, value in replaceText.items():
                        if key.lower() in cellValue.lower():
                            cellValue = cellValue.replace(key, value)
                    cell.value = cellValue

            newNDLMFile = f'{siteCode}-NDLM-Tier2.xlsx'
            ndlmFile.save(newNDLMFile)

    except FileNotFoundError:
        print("File not found. Please check the file path and try again.")
        authLog.error(f"File not found in path {ndlmPath1}")
        authLog.error(traceback.format_exc())

    except Exception as error:
        print(f"ERROR: {error}\n", traceback.format_exc())
        authLog.error(f"Wasn't possible to choose the CSV file, error message: {error}\n", traceback.format_exc())

def cEdgeTemplate(rowText):
    site_code = "test"
    newSDW03Template = f'{site_code}-SDW-03-Template.csv'
    newSDW04Template = f'{site_code}-SDW-04-Template.csv'

    sdw03Replacements = {
        'cedge1-host' : f'{rowText[2]}',
        'snmp-location' : f'{rowText[3]}',
        'cedge1-rtr-ip' : f'{rowText[6]}',
        'cEdge-asn' : f'{rowText[9]}',
        'cedge1-sw-ip' : f'{rowText[12]}',
        'switch-asn' : f'{rowText[14]}',
        'mpls-pe-ip' : f'{rowText[17]}',
        'cedge2-tloc3-ext-ip' : f'{rowText[18]}',
        'cedge2-host - gi0/0/3 - TLOC3' : f'{rowText[20]}',
        'cedge1-tloc3-ip'	: f'{rowText[21]}',
        'mpls-ce1-ip' : f'{rowText[32]}',
        'mpls-speed' : f'{rowText[38]}',
        'latitude' : f'{rowText[41]}',
        'longitude' : f'{rowText[42]}',
        # Here starts the second CSV file #
        'cedge2-host'	: f'{rowText[47]}',
        'bb1-down-speed' : f'{rowText[82]}',
        'cedge2-rtr-ip' : f'{rowText[51]}',
        'cedge2-sw-ip' : f'{rowText[57]}',	
        'cedge2-tloc3-gate' : f'{rowText[63]}',	
        'cedge1-host TLOC3 gi0/0/3' : f'{rowText[59]}',
        'cedge2-tloc3-ext-ip/30' : f'{rowText[60]}',
        'bb1-up-speed' : f'{rowText[82]}',	
        'mpls-ce2-ip'	: f'{rowText[85]}',
        'sw-host' : f'{rowText[13]}'
    }

    sdw04Replacements = {


    }


    # sdw03Replacements = {
    #     'cedge1-serial-no': cedge1_serial_no,
    #     'cedge1-loop': cedge1_loop,
    #     'cedge1-host': cedge1_host,
    #     'snmp-location': snmp_location,
    #     'sw-host': sw_host,
    #     'sw-cedge1-port': sw_cedge1_port,
    #     'cedge1-rtr-ip': cedge1_rtr_ip,
    #     'cEdge-asn': cEdge_asn,
    #     'cedge1-sw-ip': cedge1_sw_ip,
    #     'switch-asn': switch_asn,
    #     'sw-mgmt-ip': sw_mgmt_ip,
    #     'mpls-pe-ip': mpls_pe_ip,
    #     'cedge2-tloc3-ext-ip': cedge2_tloc3_ext_ip,
    #     'cedge2-host': cedge2_host,
    #     'cedge1-tloc3-ip': cedge1_tloc3_ip,
    #     'sw-cedge1-mpls-port': sw_cedge1_mpls_port,
    #     'mpls-circuitid': mpls_circuitid,
    #     'mpls-ce1-ip': mpls_ce1_ip,
    #     'mpls-speed': mpls_speed,
    #     'latitude': latitude,
    #     'longitude': longitude,
    #     'site-no': site_no
    #     }
    
    # sdw04Replacements = {
    #     'cedge2-serial-no': cedge2_serial_no,
    #     'cedge2-loop': cedge2_loop, 
    #     'cedge2-host': cedge2_host,
    #     'cedge2-tloc3-port': cedge2_tloc3_port,
    #     'bb1-down-speed': bb1_down_speed,
    #     'snmp-location': snmp_location,
    #     'sw-host': sw_host,
    #     'sw-cedge2-port': sw_cedge2_port,
    #     'cedge2-rtr-ip': cedge2_rtr_ip,
    #     'cEdge-asn': cEdge_asn,
    #     'cedge2-sw-ip': cedge2_sw_ip,
    #     'switch-asn': switch_asn,
    #     'sw-mgmt-ip': sw_mgmt_ip,
    #     'cedge2-tloc3-gate': cedge2_tloc3_gate,
    #     'mpls-pe-ip': mpls_pe_ip,
    #     'cedge1-host': cedge1_host,
    #     'cedge2-tloc3-ext-ip': cedge2_tloc3_ext_ip,
    #     'bb1-carrier': bb1_carrier,
    #     'bb1-circuitid': bb1_circuitid,
    #     'cedge2-tloc3-ip': cedge2_tloc3_ip,
    #     'cedge2-tloc3-cidr': cedge2_tloc3_cidr,
    #     'bb1-up-speed': bb1_up_speed,
    #     'sw-cedge2-mpls-port': sw_cedge2_mpls_port, 
    #     'mpls-circuitid': mpls_circuitid,
    #     'mpls-ce2-ip': mpls_ce2_ip,
    #     'mpls-speed': mpls_speed,
    #     'latitude': latitude,
    #     'longitude': longitude,
    #     'site-no': site_no
    # }

    try:
        with open(sdw03Template, "r") as inputCSV, \
            open(newSDW03Template, 'w') as outputCSV:
            authLog.info(f"Generating {site_code}-SDW-03-Template")
            print(f"INFO: Generating {site_code}-SDW-03-Template.")
            csvReader = csv.reader(inputCSV)
            csvWriter = csv.writer(outputCSV)   

            for rows in csvReader:
                rowData = []
                for cell in rows:
                    cellValue = str(cell).strip()
                    for key, value in sdw03Replacements.items():
                        if key.lower() in cellValue.lower():
                            cellValue = cellValue.replace(key, value)
                    rowData.append(cellValue)
                csvWriter.writerow(rowData)
        
        with open(sdw04Template, "r") as inputCSV1, \
            open(newSDW04Template, 'w') as outputCSV1:
            authLog.info(f"Generating {site_code}-SDW-04-Template")
            print(f"INFO: Generating {site_code}-SDW-04-Template.")
            csvReader1 = csv.reader(inputCSV1)
            csvWriter1 = csv.writer(outputCSV1)   

            for rows in csvReader1:
                rowData1 = []
                for cell in rows:
                    cellValue1 = str(cell).strip()
                    for key, value in sdw04Replacements.items():
                        if key.lower() in cellValue1.lower():
                            cellValue1 = cellValue1.replace(key, value)
                    rowData1.append(cellValue1)
                csvWriter1.writerow(rowData1)

    except Exception as error:
        print(f"ERROR: {error}\n", traceback.format_exc())
        authLog.error(f"Error message: {error}\n", traceback.format_exc())
