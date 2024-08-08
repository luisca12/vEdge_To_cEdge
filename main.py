import os
from strings import greetingString, menuString, inputErrorString
from utils import mkdir


def main():  
    mkdir()  
    from log import authLog
    from fileHandler import chooseCSV, chooseDocx_ISR, chooseDocx_vEdge, modNDLM, modNDLM2, cEdgeTemplate
    from functions import checkIsDigit
    while True:
        os.system("CLS")
        greetingString()
        menuString()
        selection = input("Please choose the option that yyou want: ")
        if checkIsDigit(selection):
            if selection == "1":
                rowText = chooseCSV()
                docxValues, serialNumSDW01, serialNumSDW02, serialNumSDW03, serialNumSDW04, cEdge1Loop, cEdge2Loop, siteNo, city, state, siteCode, mplsCircuitID, bb1Carrier, bb1Circuitid, cEdge2TLOC3_Port, swcEdge1_vlan, swcEdge2_vlan, swcEdge1_port, swcEdge2_port, swcEdge1_mplsPort, swcEdge2_mplsPort = chooseDocx_vEdge(rowText)
                # modNDLM(docxValues['site-code'], docxValues['serialNumSDW01'], docxValues['serialNumSDW02'], 
                #         docxValues['serialNumSDW03'], docxValues['serialNumSDW04'], docxValues['cedge1-loop'], 
                #         docxValues['cedge2-loop'], docxValues['snmp-location'], docxValues['vedge1-loop'], docxValues['vedge2-loop'])
                # modNDLM2(docxValues['site-code'], docxValues['cedge1-loop'], docxValues['cedge2-loop'], 
                #         docxValues['snmp-location'], docxValues['city'], docxValues['state'], docxValues['site-no'], 
                #         docxValues['cedge1-host'], docxValues['cedge2-host'], docxValues['sw-host'], 
                #         docxValues['sw-mpls-port'], docxValues['cedge2-tloc3-port'], docxValues['sw-cedge1-port'], 
                #         docxValues['sw-cedge2-port'], docxValues['sw-cedge1-mpls-port'], docxValues['sw-cedge2-mpls-port'])
                cEdgeTemplate(docxValues)

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
                cEdgeTemplate(docxValues['cedge1-serial-no'], docxValues['cedge1-loop'], docxValues['cedge1-host'], docxValues['snmp-location'], 
                              docxValues['sw-host'], docxValues['sw-cedge1-port'], docxValues['cedge1-rtr-ip'],
                              docxValues['cEdge-asn'], docxValues['cedge1-sw-ip'], docxValues['switch-asn'], docxValues['sw-mgmt-ip'],
                              docxValues['mpls-pe-ip'], docxValues['cedge2-tloc3-ext-ip'], docxValues['cedge2-host'], docxValues['cedge1-tloc3-ip'],
                              docxValues['sw-cedge1-mpls-port'], docxValues['mpls-circuitid'], docxValues['mpls-ce1-ip'], docxValues['mpls-speed'],
                              docxValues['latitude'], docxValues['longitude'], docxValues['site-no'], docxValues['cedge2-serial-no'],
                              docxValues['cedge2-loop'], docxValues['cedge2-tloc3-port'], docxValues['bb1-down-speed'],
                              docxValues['sw-cedge2-port'], docxValues['cedge2-rtr-ip'], docxValues['cedge2-sw-ip'], docxValues['cedge2-tloc3-gate'],
                              docxValues['bb1-carrier'], docxValues['bb1-circuitid'], docxValues['cedge2-tloc3-ip'], docxValues['cedge2-tloc3-cidr'],
                              docxValues['bb1-up-speed'], docxValues['sw-cedge2-mpls-port'], docxValues['mpls-ce2-ip'], docxValues['site-code'])
        else:
            authLog.error(f"Wrong option chosen {selection}")
            inputErrorString()
            os.system("PAUSE")


if __name__ == "__main__":
    main()