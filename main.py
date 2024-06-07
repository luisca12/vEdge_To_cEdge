import os
from strings import greetingString
from utils import mkdir


def main():  
    mkdir()  
    os.system("CLS")
    greetingString()
    from log import authLog
    from fileHandler import chooseCSV, chooseDocx
    while True:
        rowText, ignoredStrings = chooseCSV()
        chooseDocx(rowText, ignoredStrings)
        break

if __name__ == "__main__":
    main()