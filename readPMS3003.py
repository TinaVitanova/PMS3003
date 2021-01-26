import time
import math
import serial
from datetime import datetime
from openpyxl import Workbook

bytes_in_one_reading = 24


wb = Workbook()


def main():
    ser = serial.Serial(
        port='COM5',
        baudrate=9600,
        timeout=1,
        parity=serial.PARITY_ODD,
        stopbits=serial.STOPBITS_ONE,
        bytesize=serial.SEVENBITS
    )
    ser.isOpen()
    # Reading the data from the serial port. This will be running in an infinite loop.
    ws = wb.active

    while 1:
        bytes_to_read = ser.inWaiting()
        data = ser.read(bytes_to_read)
        time.sleep(1)
        now = datetime.now()
        number_of_loops = math.floor(bytes_to_read / bytes_in_one_reading)
        for i in range(number_of_loops):
            count = i * bytes_in_one_reading
            if data[count:count + bytes_in_one_reading].find(0x42) != -1 \
                    and (data[count:count + bytes_in_one_reading].find(0x42) + 1
                         == data[count:count + bytes_in_one_reading].find(0x4d)):
                if check_value(data[count:count + bytes_in_one_reading]):
                    print(now)
                    pm_1_value_standard = get_pm_1_value_standard(data[count:count + bytes_in_one_reading])
                    pm_2_5_value_standard = get_pm_2_5_value_standard(data[count:count + bytes_in_one_reading])
                    pm_10_value_standard = get_pm_10_value_standard(data[count:count + bytes_in_one_reading])
                    print('standard')
                    print(pm_1_value_standard)
                    print(pm_2_5_value_standard)
                    print(pm_10_value_standard)
                    pm_1_value_atmospheric = get_pm_1_value_atmospheric(data[count:count + bytes_in_one_reading])
                    pm_2_5_value_atmospheric = get_pm_2_5_value_atmospheric(data[count:count + bytes_in_one_reading])
                    mass_concentration = get_mass_concentration(data[count:count + bytes_in_one_reading])
                    print('atmospheric')
                    print(pm_1_value_atmospheric)
                    print(pm_2_5_value_atmospheric)
                    print(mass_concentration)
                    ws.append([now, pm_1_value_standard, pm_2_5_value_standard, pm_10_value_standard,
                               pm_1_value_atmospheric, pm_2_5_value_atmospheric, mass_concentration])
        wb.save('daj_bre_raboti.xslx')


def check_value(data):
    receive_flag = 0
    receive_sum = 0
    for i in range(bytes_in_one_reading - 2):
        receive_sum = receive_sum + int(data[i:i + 1].hex(), 16)

    if (receive_sum % 256) == (int(data[bytes_in_one_reading - 1:bytes_in_one_reading].hex(), 16)):
        receive_flag = 1
    return receive_flag


def get_pm_1_value_standard(data):
    return int(data[4:5].hex(), 16) * 256 + int(data[5:6].hex(), 16)


def get_pm_2_5_value_standard(data):
    return int(data[6:7].hex(), 16) * 256 + int(data[7:8].hex(), 16)


def get_pm_10_value_standard(data):
    return int(data[8:9].hex(), 16) * 256 + int(data[9:10].hex(), 16)


def get_pm_1_value_atmospheric(data):
    return int(data[10:11].hex(), 16) * 256 + int(data[11:12].hex(), 16)


def get_pm_2_5_value_atmospheric(data):
    return int(data[12:13].hex(), 16) * 256 + int(data[13:14].hex(), 16)


def get_mass_concentration(data):
    return int(data[14:15].hex(), 16) * 256 + int(data[15:16].hex(), 16)


if __name__ == '__main__':
    main()
