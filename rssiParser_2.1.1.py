## CHANGELOG ##
# Pridejau scatter piesima su matplotlib
# Lyginant su rssiParser_2.1.0 is optimizavimo puses veikia gerai, bet funkcionaluma nenaudingas siuo atveju

import numpy as np
import matplotlib.pyplot as plt
import matplotlib.animation as animation
import serial
import statistics

rssiList = []

plt.axis("auto")

serialPort = serial.Serial(
    port = "COM11", 
    baudrate=115200, 
    bytesize=8, 
    timeout=2, 
    stopbits=serial.STOPBITS_ONE)

while (True):
    if (serialPort.in_waiting > 0):
        serialString = serialPort.readline().strip().decode("utf-8")
        if "43:9C:27:33:F0:F3" in serialString:
            print(serialString)
            stringElementsList = serialString.split()
            for value in stringElementsList:  
                if value.startswith('-') and value != ('->'):
                    rssiList.append(float(value))
                    plt.scatter(len(rssiList), value)
                    plt.pause(0.05)

plt.show()

############################
