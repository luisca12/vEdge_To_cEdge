
from netmiko.exceptions import NetMikoAuthenticationException, NetMikoTimeoutException
from functions import checkYNInput,validateIP,requestLogin,checkReachPort22
from strings import greetingString
from log import *
from log import invalidIPLog
import socket
import traceback
import csv
import os
import logging

username = ""
execPrivPassword = ""
netDevice = {}
swHostname = ""

def Auth(swHostname1):
    global username, execPrivPassword, netDevice, swHostname

    swHostname = swHostname1

    os.system("CLS")
    greetingString()
    while True:
        swHostname
        validateIP(swHostname) # True or False
        swHostname = checkReachPort22(swHostname)
        authLog.error(f"User {username} input the following invalid IP: {swHostname}")
        authLog.debug(traceback.format_exc())
        break
        
    swHostname, username, netDevice = requestLogin(swHostname)

    return swHostname,username,netDevice