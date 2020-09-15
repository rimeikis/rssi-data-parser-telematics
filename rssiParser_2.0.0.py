import serial
import statistics

serialPort = serial.Serial(
    port = "COM11", 
    baudrate=115200, 
    bytesize=8, 
    timeout=2, 
    stopbits=serial.STOPBITS_ONE)

rssiList = []

def pasholnx():
    while (1):
        if (serialPort.in_waiting > 0):
            serialString = serialPort.readline().strip().decode("utf-8")
            if "5B:F0:E9:55:7B:57" in serialString:
                try:
                    print(serialString)
                    rssiValue = serialString.split()
                    print(rssiValue)
                    rssiList.append(float(rssiValue[-6]))
                    print(rssiList)  
                    print("RSSI COUNT:", len(rssiList))
                    print("MIN:", min(rssiList))
                    print("MAX:", max(rssiList))
                    print("AVG:", round(statistics.mean(rssiList)),"\n")
                except:
                    print("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
                    print("Missing RSSI data!")
                    print("However, "+rssiValue[13]+"found and added!!!!!")
                    print("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
                    rssiList.append(float(rssiValue[13]))
                    rssiCount += 1

pasholnx()