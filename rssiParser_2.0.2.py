## CHANGELOG ##
# Pridejau COM portu pasirinkima programoje
# Suformatavau duomenu rodyma cmd

import sys
import serial
from serial.tools.list_ports import comports
import statistics
import datetime

def ask_for_port():
	sys.stderr.write('\n--- Available ports:\n')
	ports = []
	for n, (port, desc, hwid) in enumerate(sorted(comports()), 1):
		sys.stderr.write('--- {:2}: {:20} {!r}\n'.format(n, port, desc))
		ports.append(port)
	while True:
		port = input('--- Enter port index or full name: ')
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
	port = ask_for_port(), 
	baudrate=115200, 
	bytesize=8, 
	timeout=2, 
	stopbits=serial.STOPBITS_ONE)

mac = str(input("--- Enter target MAC address: "))

rssiList = []

def main():
	startTime = str(datetime.datetime.now())
	while (True):
		if (serialPort.in_waiting > 0):
			serialString = serialPort.readline().strip().decode("utf-8")
			if mac in serialString:
				stringElementsList = serialString.split()
				for value in stringElementsList:  
					if value.startswith('-') and value != ('->'):
						rssiList.append(float(value)) 
						sys.stderr.write('---------------------------------------------------------------------------------------------------\n')
						sys.stderr.write('|~~~~~~~~~~~~~~~~~~~RSSI DATA~~~~~~~~~~~~~~~~~~~~|~~~~~~~~~~~~~~~~~~~PARSE DETAILS~~~~~~~~~~~~~~~~|\n')
						sys.stderr.write('---------------------------------------------------------------------------------------------------\n')
						sys.stderr.write('|    {:^40}    |    {:^40}    |\n'.format("Data count: " + str(len(rssiList)), "Start: " + startTime))
						sys.stderr.write('---------------------------------------------------------------------------------------------------\n')
						sys.stderr.write('|    {:^40}    |    {:^40}    |\n'.format("Average: " + str(round(statistics.mean(rssiList))), "Target MAC: " + mac))
						sys.stderr.write('---------------------------------------------------------------------------------------------------\n')
						sys.stderr.write('|    {:^40}    |    {:^40}    |\n'.format("Max: " + str(max(rssiList)), "Hit 'CTRL + C' to interrupt"))
						sys.stderr.write('---------------------------------------------------------------------------------------------------\n')
						sys.stderr.write('|    {:^40}    |    {:^40}     \n'.format("Min: " + str(min(rssiList)), ""))
						sys.stderr.write('--------------------------------------------------\n')

main()