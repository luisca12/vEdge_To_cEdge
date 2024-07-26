
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
validIPs = []

def Auth():
    global username, execPrivPassword, netDevice, validIPs

    os.system("CLS")
    greetingString()
    while True:
        deviceIPs = input("\nPlease enter the devices IPs/hostnames separated by commas: ")
        deviceIPsList = deviceIPs.split(',')

        for ip in deviceIPsList:
            ip = ip.strip()
            if validateIP(ip):
                IPreachChecked = checkReachPort22(ip)
                validIPs.append(IPreachChecked)
            else:
                print(f"Invalid IP address format: {ip}, will be skipped.")
                authLog.error(f"User {username} input the following invalid IP: {ip}")
                authLog.debug(traceback.format_exc())
        if validIPs:
            break
    validIPs, username, netDevice = requestLogin(validIPs)

    return validIPs,username,netDevice