'''
import serial
import time
x = 0
while x!=1:
    message = input("message here:")
    ser = serial.Serial(port='COM4', baudrate=9600)
    time.sleep(2)
    ser.write(message.encode())
    print(message)
    ser.close()
'''
x = "Hello there"
print(x.split("?"))