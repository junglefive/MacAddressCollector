import serial
import serial.tools.list_ports
import sys,time,datetime
# 打印输出当前可用串口
port_list = serial.tools.list_ports.comports()
for port in port_list:
    print(port)
print('请输入串口（例：COM28）')
name = input()
try:
    ser = serial.Serial()
    ser.port = name
    serial.baudrate = 115200
    ser.open()
    send_number = 0
    address_cmd = [0x53,0x4D,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x10,0x06,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x1E]
    while True:
        timeStamp = hex(int(time.time()))
        bytes_time = bytes().fromhex(timeStamp[2:10])
        # 初始化时间戳
        address_cmd[11] = ((send_number&&0xffffff)>>16)&0xff
        address_cmd[12] = ((send_number&&0xffffff)>>8)&0xff
        address_cmd[13] = ((send_number&&0xffffff))&0xff
        address_cmd[14] = 0x1e
        address_cmd[15] = 0xb2
        address_cmd[16] = 0xc8
        ser.write(address_cmd)
        send_number+=1
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')+'---->发送循环次数:', send_number)
        time.sleep(1)

except Exception as e:
    print(str(e))
finally:
    ser.close()
    pass




