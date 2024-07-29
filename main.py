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
        siteCode, serialNumSDW01, serialNumSDW02, serialNumSDW03, serialNumSDW04, cEdge1Loop, cEdge2Loop, snmpLocation = chooseDocx(rowText)
        modNDLM(siteCode, serialNumSDW01, serialNumSDW02, serialNumSDW03, serialNumSDW04, cEdge1Loop, cEdge2Loop, snmpLocation)
        modNDLM2(siteCode)
        break

if __name__ == "__main__":
    main()