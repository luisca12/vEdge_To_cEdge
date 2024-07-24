from netmiko import ConnectHandler
from log import authLog

import traceback
import ipaddress
import re
import os

shInt1101 = "show run interface vlan 1101 | inc ip address"
shInt1103 = "show run interface vlan 1103 | inc ip address"
shIntDesSDW = "show interface description | inc SDW|sdw"
shIntDesCON = "show interface description | inc CON|con"
shIntStatMPLS = "show interface status | inc LUM"

shVlanMgmt = f"show interface vlan 1500"
shHostname = "show run | i hostname"

def shCoreInfo(validIPs, username, netDevice):
    # This function is to take a show run

    for validDeviceIP in validIPs:
        try:
            validDeviceIP = validDeviceIP.strip()
            currentNetDevice = {
                'device_type': 'cisco_xe',
                'ip': validDeviceIP,
                'username': username,
                'password': netDevice['password'],
                'secret': netDevice['secret'],
                'global_delay_factor': 2.0,
                'timeout': 120,
                'session_log': 'netmikoLog.txt',
                'verbose': True,
                'session_log_file_mode': 'append'
            }

            print(f"Connecting to device {validDeviceIP}...")
            with ConnectHandler(**currentNetDevice) as sshAccess:
                try:
                    sshAccess.enable()
                    shHostnameOut = sshAccess.send_command(shHostname)
                    authLog.info(f"User {username} successfully found the hostname {shHostnameOut}")
                    shHostnameOut = shHostnameOut.replace('hostname', '')
                    shHostnameOut = shHostnameOut.strip()
                    shHostnameOut = shHostnameOut + "#"

                    shInt1101Out = sshAccess.send_command(shInt1101)
                    ipVlan1101 = shInt1101Out.split(' ')[2]
                    maskVlan1101 = shInt1101Out.split(' ')[3]
                    netVlan1101 = ipaddress.IPv4Network(f"{ipVlan1101}/{maskVlan1101}", strict=False)
                    netVlan1101 = netVlan1101.network_address

                    shInt1103Out = sshAccess.send_command(shInt1103)
                    ipVlan1103 = shInt1103Out.split(' ')[2]
                    maskVlan1103 = shInt1103Out.split(' ')[3]
                    netVlan1103 = ipaddress.IPv4Network(f"{ipVlan1103}/{maskVlan1103}", strict=False)
                    netVlan1103 = netVlan1103.network_address

                    shIntDesSDWOut = sshAccess.send_command(shIntDesSDW)

                    shIntDesCONOut = sshAccess.send_command(shIntDesCON)

                    shVlanMgmtOut = sshAccess.send_command(shVlanMgmt)

                    return netVlan1101, netVlan1103

                except Exception as error:
                    print(f"ERROR: An error occurred: {error}\n", traceback.format_exc())
                    authLog.error(f"User {username} connected to {validDeviceIP} got an error: {error}")
                    authLog.debug(traceback.format_exc(),"\n")
       
        except Exception as error:
            print(f"ERROR: An error occurred: {error}\n", traceback.format_exc())
            authLog.error(f"User {username} connected to {validDeviceIP} got an error: {error}")
            authLog.debug(traceback.format_exc(),"\n")
            with open(f"failedDevices.txt","a") as failedDevices:
                failedDevices.write(f"User {username} connected to {validDeviceIP} got an error.\n")
        
        finally:
            print(f"Outputs and files successfully created for device {validDeviceIP}.\n")
            print("For any erros or logs please check Logs -> authLog.txt\n")