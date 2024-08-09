import os
from strings import greetingString, menuString, inputErrorString, menuStringEnd
from utils import mkdir


def main():  
    mkdir()  
    from log import authLog
    from fileHandler import chooseCSV, chooseDocx_ISR, chooseDocx_vEdge, modNDLMvEdge, modNDLM2vEdge, cEdgeTemplate
    from functions import checkIsDigit
    while True:
        os.system("CLS")
        greetingString()
        menuString()
        selection = input("Please choose the option that you want: ")
        if checkIsDigit(selection):
            if selection == "1":
                csvValues = chooseCSV()
                docxValues = chooseDocx_vEdge(csvValues)
                rowText = docxValues['rowText']
                rowText1 = docxValues['rowText1']
                for index, item in enumerate(rowText):
                        print(f"rowText[{index}] with string: {item}")
                os.system("PAUSE")
                for index, item in enumerate(rowText1):
                        print(f"rowText1[{index}] with string: {item}")
                os.system("PAUSE")
                modNDLMvEdge(rowText, rowText1)
                modNDLM2vEdge(rowText, rowText1)
                cEdgeTemplate(rowText, rowText1)

            if selection == "2":
                csvValues = chooseCSV()
                docxValues = chooseDocx_ISR(csvValues)
                rowText = docxValues['rowText']
                rowText1 = docxValues['rowText1']
                modNDLMvEdge(rowText, rowText1)
                modNDLM2vEdge(rowText, rowText1)
                cEdgeTemplate(rowText, rowText1)

        else:
            authLog.error(f"Wrong option chosen {selection}")
            inputErrorString()
            os.system("PAUSE")

        menuStringEnd()
        selection = input("Please choose the option that you want: ")
        if checkIsDigit(selection):
            if selection == "1":
                csvValues = chooseCSV()
                docxValues = chooseDocx_vEdge(csvValues)
                rowText = docxValues['rowText']
                rowText1 = docxValues['rowText1']
                for index, item in enumerate(rowText):
                        print(f"rowText[{index}] with string: {item}")
                os.system("PAUSE")
                for index, item in enumerate(rowText1):
                        print(f"rowText1[{index}] with string: {item}")
                os.system("PAUSE")
                modNDLMvEdge(rowText, rowText1)
                modNDLM2vEdge(rowText, rowText1)
                cEdgeTemplate(rowText, rowText1)

            if selection == "2":
                csvValues = chooseCSV()
                docxValues = chooseDocx_ISR(csvValues)
                rowText = docxValues['rowText']
                rowText1 = docxValues['rowText1']
                modNDLMvEdge(rowText, rowText1)
                modNDLM2vEdge(rowText, rowText1)
                cEdgeTemplate(rowText, rowText1)

            if selection == "3":
                 break
        
        else:
            authLog.error(f"Wrong option chosen {selection}")
            inputErrorString()
            os.system("PAUSE")

if __name__ == "__main__":
    main()