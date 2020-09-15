## CHANGELOG ##
# sutvarkiau laika, kuris bloga duration ikelia i exceli.
# pridet available MAC adresu sarasa po scan'o. Islistint unikalius IMEI ir ju MAC salia vienas kito. Su 1,2,3... inputais pasirinkt kurio rssi nori detektint

from datetime import datetime
import statistics
from serial.tools.list_ports import comports
import serial
import sys
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font
import re
VERSION = "rssiParser_2.0.6"


def ask_for_port():
    sys.stderr.write('\n--- Available ports:\n')
    ports = []
    for n, (port, desc, hwid) in enumerate(sorted(comports()), 1):
        sys.stderr.write('--- {:2}: {:20} {!r}\n'.format(n, port, desc))
        ports.append(port)
    while True:
        port = input('>>> Enter port index or full name: ')
        try:
            index = int(port) - 1
            if not 0 <= index < len(ports):
                sys.stderr.write('--- Invalid index!\n')
                continue
        except ValueError:
            pass
        else:
            port = ports[index]
        return port


serialPort = serial.Serial(
    port=ask_for_port(),
    baudrate=115200,
    bytesize=8,
    timeout=2,
    stopbits=serial.STOPBITS_ONE)


def device_scan():
    deviceList = []
    while (True):
        if (serialPort.in_waiting > 0):
            serialString = serialPort.read_until(b'\r')
            if b', Power: ' in serialString:

                stringElementsList = serialString.split()
                for value in stringElementsList:
                    deviceList.append()
                    print('---------------------')
                    print(value[7])


device_scan()
