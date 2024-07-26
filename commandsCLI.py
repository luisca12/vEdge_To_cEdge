from netmiko import ConnectHandler
from log import authLog

import traceback
import ipaddress
import re
import os

shInt1101 = "show run interface vlan 1101 | inc ip address"
shInt1103 = "show run interface vlan 1103 | inc ip address"
shVlanMgmt = "show run interface vlan 1500 | inc ip address"
shLoop0 = "show run interface lo0 | inc ip address"

shIntDesSDW = "show interface description | inc SDW|sdw"
shIntDesCON = "show interface description | inc CON"

shIntStatMPLS = "show interface status | inc LUM"
shHostname = "show run | i hostname"

intPatt = r'[a-zA-Z]+\d+\/(?:\d+\/)*\d+'

def shCoreInfo(swHostname, username, netDevice):

    try:
        currentNetDevice = {
            'device_type': 'cisco_xe',
            'ip': swHostname,
            'username': username,
            'password': netDevice['password'],
            'secret': netDevice['secret'],
            'global_delay_factor': 2.0,
            'timeout': 120,
            'session_log': 'netmikoLog.txt',
            'verbose': True,
            'session_log_file_mode': 'append'
        }

        print(f"Connecting to device {swHostname}...")
        with ConnectHandler(**currentNetDevice) as sshAccess:
            try:
                sshAccess.enable()
                shHostnameOut = sshAccess.send_command(shHostname)
                authLog.info(f"User {username} successfully found the hostname {shHostnameOut}")
                shHostnameOut = shHostnameOut.split(' ')[1]
                shHostnameOut = shHostnameOut + "#"
                print(f"This is the hostname: {shHostnameOut}")

                shInt1101Out = sshAccess.send_command(shInt1101)
                ipVlan1101 = shInt1101Out.split(' ')[3]
                print(f"This is ipVlan1101: {ipVlan1101}")
                maskVlan1101 = shInt1101Out.split(' ')[4]
                print(f"This is maskVlan1101: {maskVlan1101}\n")
                netVlan1101 = ipaddress.IPv4Network(f"{ipVlan1101}/{maskVlan1101}", strict=False).network_address

                shInt1103Out = sshAccess.send_command(shInt1103)
                ipVlan1103 = shInt1103Out.split(' ')[3]
                print(f"This is ipVlan1103: {ipVlan1103}")
                maskVlan1103 = shInt1103Out.split(' ')[4]
                print(f"This is maskVlan1103: {maskVlan1103}\n")
                netVlan1103 = ipaddress.IPv4Network(f"{ipVlan1103}/{maskVlan1103}", strict=False).network_address

                shIntDesSDWOut = sshAccess.send_command(shIntDesSDW)
                print(f"This is {shIntDesSDW}:\n{shIntDesSDWOut}\n")

                shIntDesCONOut = sshAccess.send_command(shIntDesCON)
                shIntDesCONOut1 = re.findall(intPatt, shIntDesCONOut)
                print(f"Show int Des | inc Con:\n{shIntDesCONOut}\nInterfaces:{shIntDesCONOut1}\n")

                shIntStatMPLSOut = sshAccess.send_command(shIntStatMPLS)
                shIntStatMPLSOut1 = re.findall(intPatt, shIntStatMPLSOut)
                print(f"Show int Des | inc Con:\n{shIntStatMPLSOut}\nInterfaces:{shIntStatMPLSOut1}\n")

                shVlanMgmtOut = sshAccess.send_command(shVlanMgmt)
                shVlanMgmtIP = shVlanMgmtOut.split(' ')[3]
                print(f"This is shVlanMgmtIP: {shVlanMgmtIP}")
                shVlanMgmtMask = shVlanMgmtOut.split(' ')[4]
                print(f"This is shVlanMgmtMask: {shVlanMgmtMask}")
                shVlanMgmtCIDR = ipaddress.IPv4Network(f'{shVlanMgmtIP}/{shVlanMgmtMask}', strict=False).prefixlen
                print(f"This is shVlanMgmtCIDR: {shVlanMgmtCIDR}\n")

                shLoop0Out = sshAccess.send_command(shLoop0)
                shLoop0Out = shLoop0Out.split(' ')[3]
                print(f"This is switch loopback 0 IP: {shLoop0Out}\n")
                os.system("PAUSE")

                return shHostnameOut, netVlan1101, netVlan1103, shIntDesSDWOut, shIntDesCONOut1, shIntStatMPLSOut1, shVlanMgmtIP, shVlanMgmtCIDR, shLoop0Out

            except Exception as error:
                print(f"ERROR: An error occurred: {error}\n", traceback.format_exc())
                authLog.error(f"User {username} connected to {swHostname} got an error: {error}")
                authLog.debug(traceback.format_exc(),"\n")
    
    except Exception as error:
        print(f"ERROR: An error occurred: {error}\n", traceback.format_exc())
        authLog.error(f"User {username} connected to {swHostname} got an error: {error}")
        authLog.debug(traceback.format_exc(),"\n")
        with open(f"failedDevices.txt","a") as failedDevices:
            failedDevices.write(f"User {username} connected to {swHostname} got an error.\n")