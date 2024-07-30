import os
from strings import greetingString
from utils import mkdir


def main():  
    mkdir()  
    os.system("CLS")
    greetingString()
    from log import authLog
    from fileHandler import chooseCSV, chooseDocx, modNDLM, modNDLM2
    while True:
        rowText = chooseCSV()
        docxValues = chooseDocx(rowText)
        modNDLM(docxValues['site-code'], docxValues['serialNumSDW01'], docxValues['serialNumSDW02'], 
                docxValues['serialNumSDW03'], docxValues['serialNumSDW04'], docxValues['cedge1-loop'], 
                docxValues['cedge2-loop'], docxValues['snmp-location'])
        modNDLM2(docxValues['site-code'], docxValues['cedge1-loop'], docxValues['cedge2-loop'], 
                docxValues['snmp-location'], docxValues['city'], docxValues['state'], docxValues['site-no'], 
                docxValues['cedge1-host'], docxValues['cedge2-host'], docxValues['sw-host'], 
                docxValues['sw-mpls-port'], docxValues['cedge2-tloc3-port'], docxValues['sw-cedge1-port'], 
                docxValues['sw-cedge2-port'], docxValues['sw-cedge1-mpls-port'], docxValues['sw-cedge2-mpls-port'])
        break

if __name__ == "__main__":
    main()