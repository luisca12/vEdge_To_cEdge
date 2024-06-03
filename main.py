from log import authLog
import os
import logging.config
from strings import greetingString
from utils import mkdir
from fileHandler import chooseCSV, chooseDocx

def main():    
    mkdir()
    os.system("CLS")
    greetingString()

    while True:
        chooseCSV()
        break

if __name__ == "__main__":
    main()