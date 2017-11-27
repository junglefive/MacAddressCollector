import  queue
import  _thread
import sys; import time; import traceback; import  datetime
import serial
import serial.tools.list_ports
from   serial.threaded import *
import sqlite3 as lite
import os
import datetime
from PIL import Image
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog, QPrinterInfo
from PyQt5.QtWidgets import *
from serial.tools.list_ports import *
import shutil
from main_window_ui import *
import sqlite3 as lite
from mac_functions import  *
import xlwt
import shutil
from picture_rc import  *


class MyApp(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(MyApp,self).__init__()
        QtWidgets.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        # logo
        self.setWindowIcon(QIcon(":picture/img/110.png"))
        # object
        self.total_macs = get_defualt_num()
        # function
        self.resize_display()
        self.update_current_num(False)
        self.uart_thread = UartReaderQThread()
        self.uart_detect_thread = UartAutoDetect()
        self.gen_xls_thread = GenXls()
        self.append_dis_info("自动识别串口中,请连接测试架")
        self.on_clicked_btn_get_name()
        # sin
        self.uart_detect_thread.sin_dis_port_name.connect(self.run_decode_thread)
        self.uart_thread.sin_dis_str.connect(self.append_dis_info)
        self.uart_thread.sin_btn_dis_str.connect(self.set_btn_dis_text)
        self.uart_thread.sin_current_num_bool.connect(self.update_current_num)
        self.uart_thread.sin_uart_lost_bool.connect(self.uart_lost)
        self.gen_xls_thread.sin_update_progress_bar_int.connect(self.update_progress_bar)
        self.gen_xls_thread.sin_gen_xls_int.connect(self.gen_xle_erro_callback)
        self.btn_export_excle.clicked.connect(self.on_clicked_btn_export_excle)
        self.btn_save_log.clicked.connect(self.on_clicked_btn_save_log)
        self.btn_get_name.clicked.connect(self.on_clicked_btn_get_name)
        self.btn_close_voice.clicked.connect(self.btn_close_voice_on_click)
        self.btn_change_voice.clicked.connect(self.btn_change_voice_on_click)


    def btn_close_voice_on_click(self):
        print("btn_close_voice_on_click")
        if self.uart_thread.bool_close_voice:
            self.uart_thread.bool_close_voice = False
            self.btn_close_voice.setText("点击开启声音")
            QMessageBox.information(self,"提示","提示声音已关闭")
        else:
            self.uart_thread.bool_close_voice = True
            QMessageBox.information(self, "提示", "提示声音已打开")

    def btn_change_voice_on_click(self):
        print("btn_change_voice_on_click")
        if self.uart_thread.bool_voice_model :
            self.uart_thread.bool_voice_model = False
            QMessageBox.information(self, "提示", "已切换到语音模式")
        else:
            self.uart_thread.bool_voice_model = True
            QMessageBox.information(self, "提示", "已切换到蜂鸣器模式")

    def set_btn_dis_text(self,dis_str):

        self.btn_get_name.setText(dis_str)
        self.set_green_text(self.btn_get_name)
        time.sleep(0.5)
        self.set_black_text(self.btn_get_name)


    def set_green_text(self, comp):

        palette = QPalette()
        palette.setColor(QPalette.ButtonText, Qt.darkRed)
        comp.setPalette(palette)
        self.refresh_app()

    def set_black_text(self, comp):
        palette = QPalette()
        palette.setColor(QPalette.ButtonText, Qt.black)
        comp.setPalette(palette)
        self.refresh_app()
        pass
    def gen_xle_erro_callback(self,erro_code):

        if erro_code == self.gen_xls_thread.GEN_XLS_NO_DB:
            QMessageBox.information(self,"错误","数据库中没有数据！")
        elif erro_code == self.gen_xls_thread.GEN_XLS_NO_RAM:
            QMessageBox.information(self,"错误","导出excle表格失败,请检查内存！")
        else:
            pass

    def update_progress_bar(self, num):

        self.progressBar_xls.setValue(num)
        if num == 100:
            QMessageBox.information(self,"提示","导出成功，请到当前目录excle文件夹下边查看")

    def resize_display(self):
        self.__desktop = QApplication.desktop()
        qRect = self.__desktop.screenGeometry()  # 设备屏幕尺寸
        self.resize(qRect.width()*2 / 3, qRect.height() *3/ 4)
        self.move(qRect.width() / 3, qRect.height() / 30)
    def uart_lost(self):
        self.btn_get_name.setText("点击自动识别串口")
        self.uart_thread.quit()
        QMessageBox.information(self,"提示","串口丢失，请检查连接")

    def run_decode_thread(self,name):
        self.uart_thread.name = name
        if self.uart_detect_thread.isRunning():
            self.uart_detect_thread.quit()
        self.btn_get_name.setText("请放入蓝牙设备，进行测试")
        if not self.uart_thread.alive:
            self.uart_thread.start()
        self.append_dis_info("成功连接到串口:"+name)
        QMessageBox.information(self,"提示","成功连接到串口:"+name)
        self.set_green_text(self.btn_get_name)

    def on_clicked_btn_get_name(self):
        print("正在识别中....")

        if self.uart_thread.isRunning():
           QMessageBox.information(self,"提示","串口正在运行中...")
        else:
            if self.uart_detect_thread.isRunning():
                print("关闭重启扫描线程")
                self.uart_detect_thread.quit()
            self.btn_get_name.setText("识别中，请等待")
            self.append_dis_info("正在识别中，请等待...")
            self.uart_detect_thread.start()

    def on_clicked_btn_save_log(self):
        print("save")
        reply = QMessageBox.question(self,"提示","此操作将删除之前的数据库，请确认已经导出Excle，再执行此操作！",QMessageBox.Yes|QMessageBox.No)
        if reply == QMessageBox.Yes:
            time_stamp = datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
            try:
                self.total_macs = 0
                self.append_dis_info("已删除记录")
                dir_db = "./data/chipsea.db"
                flag = os.path.isfile(dir_db)
                if flag == True:
                    shutil.copy(dir_db,"./log/%s.db"%(time_stamp))
                    os.remove(dir_db)
                else:
                    print("目录不存在")
            except Exception as e:
                print("on_clicked_btn_save_log:"+str(e))
            finally:
                self.update_current_num(False)

    def on_clicked_btn_export_excle(self):
        con = None
        try:

            if not self.gen_xls_thread.isRunning():
                self.append_dis_info("正在导出，请勿关闭软件")
                self.gen_xls_thread.start()
            else:
                self.append_dis_info("正在导出中....")
        except Exception as e:
            save_to_erro_log("发生异常:" + str(e))
            print("on_clicked_btn_export_excle:"+str(e))

    def refresh_app(self):
        qApp.processEvents()

    def append_dis_info(self,info):
        time_stamp = datetime.datetime.now().strftime('%H:%M:%S: ')
        self.plainText_display.appendPlainText(time_stamp+info+"->当前数量:"+str(self.total_macs))

        self.refresh_app()
    def update_current_num(self,inc_flag):
        if inc_flag:
            self.total_macs += 1
            self.textEdit_currentNum.setHtml(get_green_html(self.total_macs))
            print("更新数量")
        else:
            self.textEdit_currentNum.setHtml(get_red_html(self.total_macs))
        self.refresh_app()

class GenXls(QThread):
    """docstring for GenXls"""
    GEN_XLS_NO_DB = 1
    GEN_XLS_NO_RAM = 2
    sin_update_progress_bar_int = pyqtSignal(int)
    sin_gen_xls_int = pyqtSignal(int)
    def __init__(self):

        super(GenXls, self).__init__()
    def run(self):
         self.gen_chipsea_xls()
    def gen_chipsea_xls(self):
        file = None
        con = None
        cur = None
        time_stamp = datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
        dir_db = "./data/chipsea.db"
        flag = os.path.isfile(dir_db)
        if flag == False:
            self.sin_gen_xls_int.emit(self.GEN_XLS_NO_DB)
        else:
            con = lite.connect('./data/chipsea.db')
            cur = con.cursor()
            cur.execute("CREATE TABLE IF NOT EXISTS mac_address_table(cur_time TEXT, mac_address TEXT)")
            cur.execute("CREATE TABLE IF NOT EXISTS mac_address_table_repeat(cur_time TEXT, mac_address TEXT)")
            cur.execute("SELECT * FROM mac_address_table")
            # 生成数据
            try:
                address = cur.fetchall()
                total = len(address)
                if total == 0:
                    self.sin_gen_xls_int.emit(self.GEN_XLS_NO_DB)
                file = open("./excle/[%s]芯海MAC地址-共%s个.csv"%(time_stamp,str(total)),"w+")
                i = 0;
                for row in address:
                    i = i+1;
                    file.write(row[0]+","+row[1]+"\n")
                    self.sin_update_progress_bar_int.emit(int(i*100/total))
            except Exception as e:
                print("gen_chipsea_xls:"+str(e))
                self.sin_gen_xls_int.emit(self.GEN_XLS_NO_RAM)
            finally:
                if cur:
                    cur.close()
                if con:
                    con.commit()
                    con.close()
                if file:
                    file.close()




class UartAutoDetect(QThread):
    """docstring for UartAutoDetect"""
    sin_dis_port_name = pyqtSignal(str)
    def __init__(self):
        super(UartAutoDetect, self).__init__()
    def run(self):
        self.flag, self.name = auto_detected_cc2640()
        if self.flag:
            self.sin_dis_port_name.emit(self.name)

class UartReaderQThread(QThread):
    """docstring for UartReaderQThread"""
    sin_dis_str = pyqtSignal(str)
    sin_btn_dis_str = pyqtSignal(str)
    sin_current_num_bool = pyqtSignal(bool)
    sin_uart_lost_bool = pyqtSignal(bool)
    last_mac = " "
    bytes_queue = None
    name = False
    port_name = None
    alive = False
    bool_close_voice = False
    bool_voice_model = False
    timeout = 0
    def __init__(self):
        super(UartReaderQThread, self).__init__()
        self.alive = False
        self.bytes_queue = queue.Queue()
    #thread-decoding
    def run(self):
        try:
            try:
                self.serial = serial.Serial(self.name,baudrate=115200,timeout=1)
                self.alive = True
                print("alive")
            except Exception as e:
                self.alive = False
                print("dead")
            while True:
                print("running...")
                if self.alive and self.serial.is_open:
                    # time.sleep(0.1)
                    self.timeout = self.timeout +1
                    if self.timeout > 4:
                        self.sin_btn_dis_str.emit("请放入蓝牙设备！")
                        self.last_mac = " "
                    print("timeout:"+str(self.timeout))
                    data = self.serial.read(self.serial.in_waiting or 1)
                    for byte in data:
                        self.bytes_queue.put(byte)
                    if self.bytes_queue.qsize() >= 32:

                        if self.bytes_queue.get() == 0x53:
                            if self.bytes_queue.get()== 0x4D:
                                # 丢掉3bytes的ID标识、4bytes的命令字节
                                for i in range(7):
                                    self.bytes_queue.get();
                                command = self.bytes_queue.get();
                                cmd_length = self.bytes_queue.get();
                                if command == 0x10 and cmd_length == 0x06:
                                    mac_address = [0, 0, 0, 0, 0, 0]
                                    for i in range(6):
                                        mac_address[i] = self.bytes_queue.get()
                                    mac_address.reverse()
                                    address = ' '.join('{:02x}'.format(x) for x in mac_address)
                                    for i in range(15):
                                        self.bytes_queue.get()
                                    print("解析到Mac地址:"+str(address))
                                    if mac_address[0] == 0xC8 and mac_address[1] == 0xB2 and mac_address[2] == 0x1E:
                                        result = save_to_database(address)

                                        if result:
                                            self.timeout = 0
                                            self.sin_current_num_bool.emit(True)
                                            self.sin_dis_str.emit("入库成功: "+address)
                                            self.sin_btn_dis_str.emit("入库成功，请拿下模块")
                                            if not self.bool_close_voice:
                                                # if self.bool_voice_model:
                                                    sys_beep()
                                                    try:
                                                        sys_speak_pass(get_defualt_num())
                                                    except Exception as e:
                                                        print("语音异常",str(e))
                                            time.sleep(2)

                                        else:
                                            if self.last_mac == address:
                                                print("与上一次测试一样")
                                                self.timeout = 0
                                                self.sin_btn_dis_str.emit("此蓝牙设备已入库, 请拿下设备！")
                                            else:
                                                # if not self.bool_close_voice:
                                                    # if self.bool_voice_model:
                                                sys_beep()
                                                sys_beep()
                                                sys_beep()
                                                # else:
                                                try:
                                                    sys_speak_had_in()
                                                except Exception as e:
                                                    print("语音异常", str(e))
                                                self.timeout = 0
                                                self.sin_current_num_bool.emit(False)
                                                self.sin_dis_str.emit("此蓝牙设备已入库，检查是否混料！")
                                                self.sin_btn_dis_str.emit("此蓝牙设备已入库，检查是否混料！")
                                                save_to_erro_log("此蓝牙设备已测试过，检查是否混料！" + str(address))
                                                time.sleep(3)
                                        self.last_mac = address

                                    else:
                                       print("接收到非芯海的Mac地址")
                else:
                    print("串口丢失...")
        except Exception as e:
            self.alive = False
            # raise e
            self.sin_uart_lost_bool.emit(True)
            print("UartReaderQThread-run异常",str(e))
        finally:
            print("end....")



def mainloop_app():
        try:
            pass
            app = QtWidgets.QApplication(sys.argv)
            window = MyApp()
            window.show()
            pass
        except Exception as e:
            print(str(e))
        finally:
            sys.exit(app.exec_())

if __name__ == "__main__":
    """主函数"""
    check_creat_dir("./data")
    check_creat_dir("./excle")
    check_creat_dir("./log")
    save_to_erro_log("打开软件MacSaver")
    mainloop_app()