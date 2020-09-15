## CHANGELOG ##
# Pakeiciau rssi value parsinimo logika
# Panaikinau try ir except

import serial
import statistics

serialPort = serial.Serial(
    port = 'COM11', 
    baudrate=115200, 
    bytesize=8, 
    timeout=2, 
    stopbits=serial.STOPBITS_ONE)

rssiList = []

def main():
    while (True):
        
        if (serialPort.in_waiting > 0):
            serialString = serialPort.readline().strip().decode("utf-8")
            
            if "6E:0B:8D:71:A5:11" in serialString:
                print(serialString)
                stringElementsList = serialString.split()

                for value in stringElementsList:  
                    
                    if value.startswith('-') and value != ('->'):
                        rssiList.append(float(value)) 
                        print("RSSI COUNT:", len(rssiList))
                        print("MIN:", min(rssiList))
                        print("MAX:", max(rssiList))
                        print("AVG:", round(statistics.mean(rssiList)),"\n")

main()