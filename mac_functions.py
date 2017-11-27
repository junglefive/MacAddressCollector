import sys; import time; import traceback
import  datetime
import sqlite3 as lite
import sys
import os
import xlwt
import serial
import serial.tools.list_ports
import pyttsx3
import winsound


def check_creat_dir(string):
    flag = os.path.exists(string)
    if flag == False:
        print("没有当前目录文，将自动创建")
        os.mkdir(string)

def save_to_erro_log(string):

    time_stamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    short_time_stamp = datetime.datetime.now().strftime('%Y-%m-%d')
    file = open("./data/%s.log"%(short_time_stamp),'a+')
    file.write(time_stamp+":"+string+"\n")
    file.close()

def get_defualt_num():
    total_macs = 0
    try:
        con = lite.connect('./data/chipsea.db')
        cur = con.cursor();
        cmd = "SELECT * FROM mac_address_table"
        cur.execute(cmd)
        macs = cur.fetchall()
        total_macs = len(macs)
        print("获得数据库数量:"+str(total_macs))
        con.commit()
        con.close()
    except Exception as e:
        print("get_defualt_num:"+str(e))
        save_to_erro_log("get_defualt_num异常："+str(e))
    finally:
        return  total_macs




def auto_detected_cc2640():
        had_detected = False
        com_name = None
        while True:
            try:
                print('正在识别串口....:')
                found_port = False
                print("搜索到以下串口")
                com_name_list = serial.tools.list_ports.comports()
                com_name_list.reverse()
                for port in com_name_list:
                    print(port)
                slave_state = [0x4D,0x53,0x00,0x00,0x00,0x00,0x00,0x01,0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x1E]
                for port in com_name_list:
                    print("正在测试串口:", port[0])
                    try:
                        ser = serial.Serial(port[0], baudrate=115200, timeout=1)
                        print("发送:"+bytes(slave_state).hex())
                        ser.write(slave_state)
                        recieve = ser.read(32)
                        print("收到:" + str(recieve))
                        if recieve:

                            if recieve[0] == 0x53 and recieve[1] == 0x4D:
                                com_name = port[0]
                                found_port = True
                                # 成功获取串口名
                                print("成功检测到串口-CC2640")
                                save_to_erro_log("成功检测到串口-CC2640")
                                had_detected = True
                                break
                        ser.flushOutput()
                        ser.flushInput()
                        ser.close()
                    except Exception as e:
                        print(str(e))
                        save_to_erro_log("异常"+str(e))
                    finally:
                        if ser:
                            ser.close()
                ser = None
                if found_port==True:
                    return True, com_name
            except Exception as e:
                 print(str(e))
                 save_to_erro_log("异常："+str(e))
                 return False,None

def save_to_database(mac_address):
    address = mac_address
    time_stamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    try:
        con = lite.connect('./data/chipsea.db')
        cur = con.cursor();
        cur.execute("CREATE TABLE IF NOT EXISTS mac_address_table(cur_time TEXT, mac_address TEXT)")
        cur.execute("CREATE TABLE IF NOT EXISTS mac_address_table_repeat(cur_time TEXT, mac_address TEXT)")
        cmd_select = "SELECT * FROM mac_address_table WHERE mac_address='%s'" % (address)
        cur.execute(cmd_select)
        macs = None
        macs = cur.fetchall()
        # 去重复
        print("发现重复："+str(len(macs)))
        if len(macs) == 1:
            cmd_repeat_pro = "DELETE FROM mac_address_table WHERE mac_address='%s'" % (address)
            cur.execute(cmd_repeat_pro)
            for row in macs:
                print("当前mac地址已存在!")
                save_to_erro_log("当前mac地址已存在:" + str(row))
            cmd = "insert into mac_address_table(cur_time, mac_address)values('%s', '%s');" % (
            time_stamp, address)
            cmd_rep = "insert into mac_address_table_repeat(cur_time, mac_address)values('%s', '%s');" % (
                time_stamp, address)
            cur.execute(cmd)
            cur.execute(cmd_rep)
            con.commit()
            con.close()
            return False
        else:
            cmd = "INSERT INTO mac_address_table(cur_time, mac_address)VALUES('%s', '%s');" % (
            time_stamp, address)
            cmd_rep = "INSERT INTO mac_address_table_repeat(cur_time, mac_address)VALUES('%s', '%s');" % (
            time_stamp, address)
            cur.execute(cmd)
            cur.execute(cmd_rep)
            con.commit()
            con.close()

            print("发射数据量：")
            print("当前总数量："+ ",    " + time_stamp, ":", address)
            return True

    except Exception as e:
        print("save_to_database:",str(e))


def get_green_html(i_num):
    html = """
            <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0//EN" "http://www.w3.org/TR/REC-html40/strict.dtd">
            <html><head><meta name="qrichtext" content="1" /><style type="text/css">
            p, li { white-space: pre-wrap; }
            </style></head><body style=" font-family:'Microsoft YaHei'; font-size:9pt; font-weight:400; font-style:normal;">
            <p align="center" style="-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;"><br />当前个数</p>
            <p align="center" style="-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;"><br /></p>
            <p align="center" style="-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;"><br /></p>
            <p align="center" style=" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;"><span style=" font-size:72pt; color:#00aa00;">%s</span></p></body></html>
            """%(i_num)
    return  html

def get_red_html(i_num):
    html = """
            <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0//EN" "http://www.w3.org/TR/REC-html40/strict.dtd">
            <html><head><meta name="qrichtext" content="1" /><style type="text/css">
            p, li { white-space: pre-wrap; }
            </style></head><body style=" font-family:'Microsoft YaHei'; font-size:9pt; font-weight:400; font-style:normal;">
            <p align="center" style="-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;"><br />当前个数</p>
            <p align="center" style="-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;"><br /></p>
            <p align="center" style="-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;"><br /></p>
            <p align="center" style=" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;"><span style=" font-size:72pt; color:#FF0000;">%s</span></p></body></html>
            """%(i_num)
    return  html


def sys_beep():
    winsound.Beep(2000,300)
    pass

def sys_speak_pass(num):
    engine = pyttsx3.init()
    engine.say(u"入库成功,%s个"%(num))
    engine.runAndWait()

def sys_speak_had_in():

    engine = pyttsx3.init()
    engine.say(u"已录入")
    engine.runAndWait()


if __name__ == '__main__':
    sys_speak_had_in()

