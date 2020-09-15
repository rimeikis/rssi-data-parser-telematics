## CHANGELOG ##
# Pridejau parse duomenu saugojima i faila
# Papildziau/pakeiciau cmd menu

from datetime import datetime
import statistics
from serial.tools.list_ports import comports
import serial
import sys
VERSION = "rssiParser_2.0.3"


def ask_for_port():
    sys.stderr.write('\n--- Available ports:\n')
    ports = []
    for n, (port, desc, hwid) in enumerate(sorted(comports()), 1):
        sys.stderr.write('--- {:2}: {:20} {!r}\n'.format(n, port, desc))
        ports.append(port)
    while True:
        port = input('> Enter port index or full name: ')
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


def formattedTime():
    now = datetime.now()
    dateTimeFormatted = now.strftime("%Y-%m-%d %H_%M_%S")
    return dateTimeFormatted


serialPort = serial.Serial(
    port=ask_for_port(),
    baudrate=115200,
    bytesize=8,
    timeout=2,
    stopbits=serial.STOPBITS_ONE)

mac = str(input("> Enter target MAC address: "))

rssiList = []


def main():
    startTime = datetime.now()

    try:
        while (True):
            if (serialPort.in_waiting > 0):
                serialString = serialPort.readline().strip().decode("utf-8")
                if mac in serialString:
                    stringElementsList = serialString.split()
                    for value in stringElementsList:
                        if value.startswith('-') and value != ('->'):
                            rssiList.append(int(value))
                            sys.stderr.write(
                                '\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n')
                            sys.stderr.write(
                                '|------------------------------------------------|------------------------------------------------|\n')
                            sys.stderr.write(
                                '|~~~~~~~~~~~~~~~~~~ RSSI DATA ~~~~~~~~~~~~~~~~~~~|~~~~~~~~~~~~~~~~~~ PARSE DETAILS ~~~~~~~~~~~~~~~|\n')
                            sys.stderr.write(
                                '|------------------------------------------------|------------------------------------------------|\n')
                            sys.stderr.write('|    {:^40}    |    {:^40}    |\n'.format(
                                "", "Start: " + str(startTime)))
                            sys.stderr.write(
                                '|    {:^40}    |------------------------------------------------|\n'.format("Current: " + str(value) + 'dBm'))
                            sys.stderr.write('|    {:^40}    |    {:^40}    |\n'.format(
                                "", "Target MAC: " + mac))
                            sys.stderr.write(
                                '|------------------------------------------------|------------------------------------------------|\n')
                            sys.stderr.write('|  {:^20}  | {:^20}  |  {:^40}  |\n'.format("Count: " + str(len(rssiList)), "Avg: " + str(
                                round(statistics.mean(rssiList))), "Hit 'CTRL+C' to stop parsing. A file will be"))
                            sys.stderr.write(
                                '|------------------------------------------------|{:>49}\n'.format('|'))
                            sys.stderr.write('|  {:^20}  | {:^20}  |  {:^40}  |\n'.format(
                                "Max: " + str(max(rssiList)), "Min: " + str(min(rssiList)), "generated containing summary of parsed data."))
                            sys.stderr.write(
                                '|------------------------------------------------|------------------------------------------------|')
                            sys.stderr.write('\n\n\n\n\n\n\n\n\n\n')

    except KeyboardInterrupt:
        parseResults = open('parser_results_' + formattedTime() + '.txt', "w")
        sys.stderr.write('\n--- Parsing terminated!\n')
        sys.stderr.write('--- Process duration: ' +
                         str(datetime.now() - startTime) + '\n')

        parseResults.write("RSSI PARSE RESULTS\n\n")

        parseResults.write("Start\n")
        parseResults.write(str(startTime)+"\n\n")

        parseResults.write("Duration\n")
        parseResults.write(str(datetime.now() - startTime)+"\n\n")

        parseResults.write("Target MAC\n")
        parseResults.write(str(mac)+"\n\n")

        parseResults.write("Count\n")
        parseResults.write(str(len(rssiList))+"\n\n")

        parseResults.write("Avg\n")
        parseResults.write(str(round(statistics.mean(rssiList)))+"\n\n")

        parseResults.write("Max\n")
        parseResults.write(str(max(rssiList))+"\n\n")

        parseResults.write("Min\n")
        parseResults.write(str(min(rssiList))+"\n\n")

        parseResults.write("Data\n")
        parseResults.write(str(rssiList))

        parseResults.close()

        sys.stderr.write('--- File "parser_results_' + formattedTime() +
                         '.txt" containing parsed RSSI data\n--- is next to the ' + VERSION + '.exe" launcher.')

        # Prevent CMD close
        input()


main()
