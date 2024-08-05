import os
from strings import greetingString, menuString, inputErrorString
from utils import mkdir


def main():  
    mkdir()  
    from log import authLog
    from fileHandler import chooseCSV, chooseDocx_ISR, chooseDocx_vEdge, modNDLM, modNDLM2
    from functions import checkIsDigit
    while True:
        os.system("CLS")
        greetingString()
        menuString()
        selection = input("Please choose the option that yyou want: ")
        if checkIsDigit(selection):
            if selection == "1":
                rowText = chooseCSV()
                docxValues = chooseDocx_vEdge(rowText)
                modNDLM(docxValues['site-code'], docxValues['serialNumSDW01'], docxValues['serialNumSDW02'], 
                        docxValues['serialNumSDW03'], docxValues['serialNumSDW04'], docxValues['cedge1-loop'], 
                        docxValues['cedge2-loop'], docxValues['snmp-location'], docxValues['vedge1-loop'], docxValues['vedge2-loop'])
                modNDLM2(docxValues['site-code'], docxValues['cedge1-loop'], docxValues['cedge2-loop'], 
                        docxValues['snmp-location'], docxValues['city'], docxValues['state'], docxValues['site-no'], 
                        docxValues['cedge1-host'], docxValues['cedge2-host'], docxValues['sw-host'], 
                        docxValues['sw-mpls-port'], docxValues['cedge2-tloc3-port'], docxValues['sw-cedge1-port'], 
                        docxValues['sw-cedge2-port'], docxValues['sw-cedge1-mpls-port'], docxValues['sw-cedge2-mpls-port'])
            if selection == "2":
                rowText = chooseCSV()
                docxValues = chooseDocx_ISR(rowText)
                modNDLM(docxValues['site-code'], docxValues['serialNumSDW01'], docxValues['serialNumSDW02'], 
                        docxValues['serialNumSDW03'], docxValues['serialNumSDW04'], docxValues['cedge1-loop'], 
                        docxValues['cedge2-loop'], docxValues['snmp-location'], docxValues['vedge1-loop'], docxValues['vedge2-loop'])
                modNDLM2(docxValues['site-code'], docxValues['cedge1-loop'], docxValues['cedge2-loop'], 
                        docxValues['snmp-location'], docxValues['city'], docxValues['state'], docxValues['site-no'], 
                        docxValues['cedge1-host'], docxValues['cedge2-host'], docxValues['sw-host'], 
                        docxValues['sw-mpls-port'], docxValues['cedge2-tloc3-port'], docxValues['sw-cedge1-port'], 
                        docxValues['sw-cedge2-port'], docxValues['sw-cedge1-mpls-port'], docxValues['sw-cedge2-mpls-port'])
        else:
            authLog.error(f"Wrong option chosen {selection}")
            inputErrorString()
            os.system("PAUSE")


if __name__ == "__main__":
    main()