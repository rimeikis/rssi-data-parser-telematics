## CHANGELOG ##
# Pridejau graph piesima su matplotlib
# Problemos su optimizavimu - serial while true loop ir matplotlib animate pjaunasi tarpusavyje (letas graph piesimas)

import numpy as np
import matplotlib.pyplot as plt
import matplotlib.animation as animation
import serial
import statistics

serialPort = serial.Serial(
    port = "COM11", 
    baudrate=115200, 
    bytesize=8, 
    timeout=2, 
    stopbits=serial.STOPBITS_ONE)

fig = plt.figure()
ax1 = fig.add_subplot(1,1,1)

rssiList = []
xs = []
ys = []

def animate(i):
    if (serialPort.in_waiting > 0):
        serialString = serialPort.readline().strip().decode("utf-8")
        if "48:AD:1B:A0:11:F4" in serialString:
            print(serialString)
            stringElementsList = serialString.split()
            for value in stringElementsList:  
                if value.startswith('-') and value != ('->'):
                    rssiList.append(float(value))
                    xs.append(len(rssiList))
                    ys.append(float(value))

    ax1.clear()
    ax1.plot(xs, ys)

    plt.xlabel('Data count')
    plt.ylabel('RSSI level')
    plt.title('Live RSSI level data')	
	
ani = animation.FuncAnimation(fig, animate, interval=1000) 
plt.show()