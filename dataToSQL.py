from PyQt5.QtWidgets import QSystemTrayIcon, QWidget, QApplication, QMenu, QAction, QTableWidgetItem, QMessageBox
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QDateTime, QDate, QTime
import os, signal
import sys
import autoupdate
import logging
import pypyodbc
import pymssql
from configparser import ConfigParser
import threading
import sqldataselect
import time     #获取当前时间
import xlwt
import pathlib
# import xlrd


class Mywindow(QWidget, autoupdate.Ui_Form):
    def __init__(self):
        super(Mywindow, self).__init__()
        self.setupUi(self)
        self.setWindowTitle('数据自动上传程序V.1.0.5')
        self.setWindowIcon(QIcon('load.ico'))
        self.trayico = MytrayIcon(self)
        self.x = 0
        self.j = 0
        self.y = 0
        self.messagess = '准备上传数据...'
        # 创建一个logger
        self.logger = logging.getLogger('mylogger')
        self.logger.setLevel(logging.INFO)


        # 创建一个handler，用于生成日志文件
        fh = logging.FileHandler('mylog.log')
        fh.setLevel(logging.INFO)
        # 定义日志输出格式
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(message)s')
        fh.setFormatter(formatter)
        # 给logger添加handler
        self.logger.addHandler(fh)

        try:
            myconfig = ConfigParser()
            myconfig.read(r'.\configuer.ini')
            self.mdb_path = myconfig.get('DB', 'mdbpath')
            self.dbuser = myconfig.get('DB', 'dbuser')
            self.dbpassword = myconfig.get('DB', 'dbpassword')
            self.dbname = myconfig.get('DB', 'dbname')
            self.dbip = myconfig.get('DB', 'dbip')
            self.table = myconfig.get('DB', 'table')
            self.pid = myconfig.get('DB', 'pid')
        except:
            self.mdb_path = ".\SmarX1231_OM5256_P1_local_suite 2018-10-08 19-41-17.mdb"
            self.dbuser = 'sa'
            self.dbpassword = '!ysod2018Wuhan'
            self.dbname = 'OltDb'
            self.dbip = '172.16.19.99'
            self.table = 'OM5256_P1'
            self.pid = ''

        if self.pid == 'FT1-MP1' and 'OM5256' in self.table:
            timer = threading.Timer(3, self.insert_p1_data)
            timer.start()
            self.label.setText(self.messagess)
        elif self.pid == 'FT1-MP2' and 'OM5256' in self.table:
            timer = threading.Timer(3, self.insert_p2_data)
            timer.start()
            self.label.setText(self.messagess)
        elif self.pid == 'FT1-MP3' and 'OM5256' in self.table:
            timer = threading.Timer(3, self.insert_p3_data)
            timer.start()
            self.label.setText(self.messagess)
        elif self.pid == 'FT1-MP4' and 'OM5256' in self.table:
            timer = threading.Timer(3, self.insert_p4_data)
            timer.start()
            self.label.setText(self.messagess)
        #以下部分为IP50G数据上传
        elif self.pid in ('FT1-MP1', 'FT1-MP2', 'FT2-MP3') and 'IP50G' in self.table:      #IP50G烧录和带电老化和温循没有性能数据上传
            timer = threading.Timer(3, self.insert_ip50gp0p1_data)
            timer.start()
            self.label.setText(self.messagess)
        elif self.pid == 'FT2-MP1' and 'IP50G' in self.table:       #标定
            timer = threading.Timer(3, self.insert_ip50g_mp1_data)
            timer.start()
            self.label.setText(self.messagess)
        elif self.pid == 'FT2-MP2' and 'IP50G' in self.table:       #三温测试
            timer = threading.Timer(3, self.insert_ip50g_mp2_data)
            timer.start()
            self.label.setText(self.messagess)
        elif self.pid == 'FT2-MP4' and 'IP50G' in self.table:       #回损测试
            timer = threading.Timer(3, self.insert_ip50g_mp4_data)
            timer.start()
            self.label.setText(self.messagess)
        else:
            self.label.setText('工序配置错误！')

    def closeEvent(self, event):
        event.ignore()      #忽略退出事件
        self.hide()

    def insert_p1_data(self):
        try:
            #本地mdb路劲可以在configuer.ini文件中配置
            con = pypyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb)};DBQ=%s' % self.mdb_path)
            cur = con.cursor()
            sql = "select top 1 BarCode, ToTime, IsOk, R1_GUID from Result1 where Reserved5 = '-' ORDER BY Result1ID"
            cur.execute(sql)
            result = cur.fetchall()
            self.logger.info('表1查询数据记录：%s' % result)
            if len(result) == 1:
                sql2 = "select SubItemName,ResultDesc from Result3 where R1_GUID = '%s'" % result[0][-1]
                cur.execute(sql2)
                result2 = cur.fetchall()

                if len(result2) > 0:
                    #给待插入的值赋初值
                    mATEName = ''
                    mEleVal = 0
                    mVolVal = 0
                    mPowVal = 0
                    mModuleTemp = 0
                    mEleF = 0
                    mVolF = 0
                    mPowF = 0
                    m10GBias = 0
                    m1GBias = 0
                    #赋值待插入的值
                    for i in range(len(result2)):
                        if result2[i][0] == 'ATEName':
                            mATEName = result2[i][1]
                        elif result2[i][0] == 'EleVal (P1)':
                            mEleVal = result2[i][1]
                        elif result2[i][0] == 'VolVal (P1)':
                            mVolVal = result2[i][1]
                        elif result2[i][0] == 'PowVal (P1)':
                            mPowVal = result2[i][1]
                        elif result2[i][0] == 'ModuleTemp Alarm Test':
                            mModuleTemp = result2[i][1]
                        elif result2[i][0] == 'EleVal Final(P1)':
                            mEleF = result2[i][1]
                        elif result2[i][0] == 'VolVal Final(P1)':
                            mVolF = result2[i][1]
                        elif result2[i][0] == 'PowVal Final(P1)':
                            mPowF = result2[i][1]
                        elif result2[i][0] == '10G Bias(P1)':
                            m10GBias = result2[i][1]
                        elif result2[i][0] == '1G Bias(P1)':
                            m1GBias = result2[i][1]
                    self.logger.info('表3查询数据记录：%s, %s, %s, %s, %s, %s, %s, %s, %s, %s' % (mATEName, mEleVal, mVolVal, mPowVal, mModuleTemp, \
                    mEleF, mVolF, mPowF, m10GBias, m1GBias))

                    con2 = pymssql.connect(self.dbip, self.dbuser, self.dbpassword, self.dbname)
                    sql3 = "INSERT INTO %s (SN, ToTime, IsOk, ATEName, EleVal, VolVal, PowVal,ModuleTemp, EleVal_F, VolVal_F, PowVal_F, \
                    [10GBias], [1GBias]) VALUES ('%s', '%s', '%s', '%s', '%.3f', '%.3f', '%.3f', '%.3f', '%.3f', '%.3f', '%.3f', '%.3f', '%.3f')" % (self.table, result[0][0], result[0][1], \
                    result[0][2], mATEName, float(mEleVal), float(mVolVal), float(mPowVal), float(mModuleTemp), float(mEleF), float(mVolF), float(mPowF), float(m10GBias), float(m1GBias))
                    cur2 = con2.cursor()
                    self.logger.info(sql3)
                    cur2.execute(sql3)
                    con2.commit()

                    #插入数据之后更新状态为1
                    sql4 = "update Result1 set Reserved5 = '1' where BarCode = '%s'" % result[0][0]
                    cur.execute(sql4)
                    con.commit()

                    cur.close()
                    con.close()
                    cur2.close()
                    con2.close()
                else:
                    pass
                self.x += 1
                self.j = 0
                self.messagess = '上传%s数据成功第%d次' % (self.pid, self.x)
                self.logger.info('上传数据成功第%d次' % self.x)
                self.label.setText(self.messagess)
            else:
                self.label.setText('当前无需要上传数据')
        except:
            self.messagess = '断线重连接第%d次' % self.j
            self.logger.error('断线重连接第%d次' % self.j)
            self.j += 1
            self.x = 0
        global timer
        timer = threading.Timer(2, self.insert_p1_data)
        timer.start()

    def insert_p2_data(self):
        try:
            #本地mdb路劲可以在configuer.ini文件中配置
            con = pypyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb)};DBQ=%s' % self.mdb_path)
            cur = con.cursor()
            sql = "select top 1 BarCode, ToTime, IsOk, R1_GUID from Result1 where Reserved5 = '-' ORDER BY Result1ID"
            cur.execute(sql)
            result = cur.fetchall()
            self.logger.info('表1查询数据记录：%s' % result)
            if len(result) == 1:
                sql2 = "select SubItemName,ResultDesc from Result3 where R1_GUID = '%s'" % result[0][-1]
                cur.execute(sql2)
                result2 = cur.fetchall()

                if len(result2) > 0:
                    #给待插入的值赋初值
                    mATEName = ''
                    m10GTXExtinDA = 0
                    m10GTXPWR_dBm = 0
                    m10GTXExtin_dB = 0
                    m10GTXCross = 0
                    m10GSens = 0
                    m10GSensAfter = 0
                    m1GTXExtinDA = 0
                    m1GTXPWR_dBm = 0
                    m1GTXExtin_dB = 0
                    m1GTXCross = 0
                    m10GEyeCross = 0
                    m10GEyeExtin_dB = 0
                    m10GDSAPwr_dBm = 0
                    m1GEyeCross = 0
                    m1GEyeExtin_dB = 0
                    m1GDSAPwr_dBm = 0
                    #赋值待插入的值
                    for i in range(len(result2)):
                        if result2[i][0] == 'ATEName':
                            mATEName = result2[i][1]
                        elif result2[i][0] == '10G TX Extin DA':
                            m10GTXExtinDA = result2[i][1]
                        elif result2[i][0] == '10G TX PWR(dBm)':
                            m10GTXPWR_dBm = result2[i][1]
                        elif result2[i][0] == '10G TX Extin(dB)':
                            m10GTXExtin_dB = result2[i][1]
                        elif result2[i][0] == '10G TX Cross(%)':
                            m10GTXCross = result2[i][1]
                        elif result2[i][0] == '10G Sens':
                            m10GSens = result2[i][1]
                        elif result2[i][0] == '10G Sens After':
                            m10GSensAfter = result2[i][1]
                        elif result2[i][0] == '1G TX Extin DA':
                            m1GTXExtinDA = result2[i][1]
                        elif result2[i][0] == '1G TX PWR(dBm)':
                            m1GTXPWR_dBm = result2[i][1]
                        elif result2[i][0] == '1G TX Extin(dB)':
                            m1GTXExtin_dB = result2[i][1]
                        elif result2[i][0] == '1G TX Cross(%)':
                            m1GTXCross = result2[i][1]
                        elif result2[i][0] == '10G Eye Cross(%)':
                            m10GEyeCross = result2[i][1]
                        elif result2[i][0] == '10G Eye Extin(dB)':
                            m10GEyeExtin_dB = result2[i][1]
                        elif result2[i][0] == '10G DSA Pwr(dBm)':
                            m10GDSAPwr_dBm = result2[i][1]
                        elif result2[i][0] == '1G Eye Cross(%)':
                            m1GEyeCross = result2[i][1]
                        elif result2[i][0] == '1G Eye Extin(dB)':
                            m1GEyeExtin_dB = result2[i][1]
                        elif result2[i][0] == '1G DSA Pwr(dBm)':
                            m1GDSAPwr_dBm = result2[i][1]
                    self.logger.info('表3查询数据记录：%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s' % (mATEName, m10GTXExtinDA, m10GTXPWR_dBm, \
                    m10GTXExtin_dB, m10GTXCross, m10GSens, m10GSensAfter, m1GTXExtinDA, m1GTXPWR_dBm, m1GTXExtin_dB,m1GTXCross, m10GEyeCross, m10GEyeExtin_dB, \
                    m10GDSAPwr_dBm, m1GEyeCross, m1GEyeExtin_dB, m1GDSAPwr_dBm))

                    con2 = pymssql.connect(self.dbip, self.dbuser, self.dbpassword, self.dbname)
                    sql3 = "INSERT INTO %s (SN, ToTime, IsOk, ATEName, m10GTXExtinDA, m10GTXPWR_dBm, m10GTXExtin_dB, m10GTXCross, m10GSens, m10GSensAfter, m1GTXExtinDA, m1GTXPWR_dBm, \
                    m1GTXExtin_dB, m1GTXCross, m10GEyeCross, m10GEyeExtin_dB, m10GDSAPwr_dBm, m1GEyeCross, m1GEyeExtin_dB, m1GDSAPwr_dBm) VALUES \
                    ('%s', '%s', '%s', '%s', %.2f, %.2f, %.2f, %.2f, %.2f, %.2f, %.2f, %.2f, %.2f, %.2f, %.2f, %.2f, %.2f, %.2f, %.2f, %.2f)" % (self.table, result[0][0], result[0][1], \
                    result[0][2], mATEName, float(m10GTXExtinDA), float(m10GTXPWR_dBm), float(m10GTXExtin_dB),float(m10GTXCross), float(m10GSens), float(m10GSensAfter), \
                    float(m1GTXExtinDA), float(m1GTXPWR_dBm), float(m1GTXExtin_dB), float(m1GTXCross),float(m10GEyeCross), float(m10GEyeExtin_dB), float(m10GDSAPwr_dBm), \
                    float(m1GEyeCross), float(m1GEyeExtin_dB), float(m1GDSAPwr_dBm))
                    cur2 = con2.cursor()
                    self.logger.info(sql3)
                    cur2.execute(sql3)
                    con2.commit()

                    # 插入数据之后更新状态为1
                    sql4 = "update Result1 set Reserved5 = '1' where BarCode = '%s'" % result[0][0]
                    cur.execute(sql4)
                    con.commit()

                    cur.close()
                    con.close()
                    cur2.close()
                    con2.close()
                else:
                    pass
                self.x += 1
                self.j = 0
                self.messagess = '上传%s数据成功第%d次' % (self.pid, self.x)
                self.logger.info('上传数据成功第%d次' % self.x)
                self.label.setText(self.messagess)
            else:
                self.label.setText('当前无需要上传数据')
        except:
            self.messagess = '断线重连接第%d次' % self.j
            self.logger.error('断线重连接第%d次' % self.j)
            self.j += 1
            self.x = 0
        global timer
        timer = threading.Timer(2, self.insert_p2_data)
        timer.start()

    def insert_p3_data(self):
        try:
            #本地mdb路劲可以在configuer.ini文件中配置
            con = pypyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb)};DBQ=%s' % self.mdb_path)
            cur = con.cursor()
            sql = "select top 1 BarCode, ToTime, IsOk, R1_GUID from Result1 where Reserved5 = '-' ORDER BY Result1ID"
            cur.execute(sql)
            result = cur.fetchall()
            self.logger.info('表1查询数据记录：%s' % result)
            if len(result) == 1:
                sql2 = "select SubItemName,ResultDesc from Result3 where R1_GUID = '%s'" % result[0][-1]
                cur.execute(sql2)
                result2 = cur.fetchall()

                if len(result2) > 0:
                    #给待插入的值赋初值
                    mATEName = ''
                    mAPDVol = 0
                    m10GTxNoPwrAD = 0
                    mAPDVolTest = 0
                    m1GTxPwrTest_P3 = 0
                    m1GTxPwrPerfTest_P3 = 0
                    m10GTxPwrTest_P3 = 0
                    m10GTxPwrPerfTest_P3 = 0
                    m10GTxPwrTest2_P3 = 0
                    m10GTxPwrPerfTest2_P3 = 0
                    m10GBias_P3 = 0
                    m1GBias_P3 = 0
                    #赋值待插入的值
                    for i in range(len(result2)):
                        if result2[i][0] == 'ATEName':
                            mATEName = result2[i][1]
                        elif result2[i][0] == 'APD Vol':
                            mAPDVol = result2[i][1]
                        elif result2[i][0] == '10G Tx No_Pwr AD':
                            m10GTxNoPwrAD = result2[i][1]
                        elif result2[i][0] == 'APD Vol Test':
                            mAPDVolTest = result2[i][1]
                        elif result2[i][0] == '1G TxPwrTest(P3)':
                            m1GTxPwrTest_P3 = result2[i][1]
                        elif result2[i][0] == '1G TxPwrPerfTest(P3)':
                            m1GTxPwrPerfTest_P3 = result2[i][1]
                        elif result2[i][0] == '10G TxPwrTest(P3)':
                            m10GTxPwrTest_P3 = result2[i][1]
                        elif result2[i][0] == '10G TxPwrPerfTest(P3)':
                            m10GTxPwrPerfTest_P3 = result2[i][1]
                        elif result2[i][0] == '10G TxPwrTest2(P3)':
                            m10GTxPwrTest2_P3 = result2[i][1]
                        elif result2[i][0] == '10G TxPwrPerfTest2(P3)':
                            m10GTxPwrPerfTest2_P3 = result2[i][1]
                        elif result2[i][0] == '10G Bias(P3)':
                            m10GBias_P3 = result2[i][1]
                        elif result2[i][0] == '1G Bias(P3)':
                            m1GBias_P3 = result2[i][1]
                    # result3表取12个值
                    self.logger.info('表3查询数据记录：%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s' % \
                    (mATEName, mAPDVol, m10GTxNoPwrAD, mAPDVolTest, m1GTxPwrTest_P3,m1GTxPwrPerfTest_P3, m10GTxPwrTest_P3, \
                    m10GTxPwrPerfTest_P3, m10GTxPwrTest2_P3, m10GTxPwrPerfTest2_P3, m10GBias_P3,m1GBias_P3))
                    con2 = pymssql.connect(self.dbip, self.dbuser, self.dbpassword, self.dbname)
                    sql3 = "INSERT INTO %s (SN, ToTime, IsOk, mATEName, mAPDVol, m10GTxNoPwrAD, mAPDVolTest, m1GTxPwrTest_P3, m1GTxPwrPerfTest_P3, m10GTxPwrTest_P3, \
                     m10GTxPwrPerfTest_P3, m10GTxPwrTest2_P3, m10GTxPwrPerfTest2_P3, m10GBias_P3, m1GBias_P3) VALUES \
                    ('%s', '%s', '%s', '%s', %.2f, %.2f, %.2f, %.2f, %.2f, %.2f, %.2f, %.2f, %.2f, %.2f, %.2f)" % (self.table, result[0][0], result[0][1], \
                    result[0][2], mATEName, float(mAPDVol), float(m10GTxNoPwrAD), float(mAPDVolTest),float(m1GTxPwrTest_P3), float(m1GTxPwrPerfTest_P3), float(m10GTxPwrTest_P3), \
                    float(m10GTxPwrPerfTest_P3), float(m10GTxPwrTest2_P3), float(m10GTxPwrPerfTest2_P3),float(m10GBias_P3), float(m1GBias_P3))
                    cur2 = con2.cursor()
                    self.logger.info(sql3)
                    cur2.execute(sql3)
                    con2.commit()

                    # 插入数据之后更新状态为1
                    sql4 = "update Result1 set Reserved5 = '1' where BarCode = '%s'" % result[0][0]
                    cur.execute(sql4)
                    con.commit()

                    cur.close()
                    con.close()
                    cur2.close()
                    con2.close()

                else:
                    pass
                self.x += 1
                self.j = 0
                self.messagess = '上传%s数据成功第%d次' % (self.pid, self.x)
                self.logger.info('上传数据成功第%d次' % self.x)
                self.label.setText(self.messagess)
            else:
                self.label.setText('当前无需要上传数据')
        except:
            self.messagess = '断线重连接第%d次' % self.j
            self.logger.error('断线重连接第%d次' % self.j)
            self.j += 1
            self.x = 0
        global timer
        timer = threading.Timer(2, self.insert_p3_data)
        timer.start()

    def insert_p4_data(self):
        try:
            #本地mdb路劲可以在configuer.ini文件中配置
            con = pypyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb)};DBQ=%s' % self.mdb_path)
            cur = con.cursor()
            sql = "select top 1 BarCode, ToTime, IsOk, R1_GUID from Result1 where Reserved5 = '-' ORDER BY Result1ID"
            cur.execute(sql)
            result = cur.fetchall()
            self.logger.info('表1查询数据记录：%s' % result)
            if len(result) == 1:
                sql2 = "select SubItemName,ResultDesc from Result3 where R1_GUID = '%s'" % result[0][-1]
                cur.execute(sql2)
                result2 = cur.fetchall()

                if len(result2) > 0:
                    #给待插入的值赋初值
                    mATEName = ''
                    m1GOverLoadPerf = 0
                    m10GOverLoadPerf = 0
                    m1GSENSPerf = 0
                    m10GSENSPerf = 0
                    m1GTxPwrTest_P4 = 0
                    m1GTxPwrPerfTest_P4 = 0
                    m10GTxPwrTest_P4 = 0
                    m10GTxPwrPerfTest_P4 = 0
                    #赋值待插入的值
                    for i in range(len(result2)):
                        if result2[i][0] == 'ATEName':
                            mATEName = result2[i][1]
                        elif result2[i][0] == '1G OverLoad Perf':
                            m1GOverLoadPerf = result2[i][1]
                        elif result2[i][0] == '10G OverLoad Perf':
                            m10GOverLoadPerf = result2[i][1]
                        elif result2[i][0] == '1G SENS Perf':
                            m1GSENSPerf = result2[i][1]
                        elif result2[i][0] == '10G SENS Perf':
                            m10GSENSPerf = result2[i][1]
                        elif result2[i][0] == '1G TxPwrTest(P4)':
                            m1GTxPwrTest_P4 = result2[i][1]
                        elif result2[i][0] == '1G TxPwrPerfTest(P4)':
                            m1GTxPwrPerfTest_P4 = result2[i][1]
                        elif result2[i][0] == '10G TxPwrTest(P4)':
                            m10GTxPwrTest_P4 = result2[i][1]
                        elif result2[i][0] == '10G TxPwrPerfTest(P4)':
                            m10GTxPwrPerfTest_P4 = result2[i][1]
                    # result3表取9个值
                    self.logger.info('表3查询数据记录：%s, %s, %s, %s, %s, %s, %s, %s, %s' % (mATEName, \
                    m1GOverLoadPerf, m10GOverLoadPerf,m1GSENSPerf, m10GSENSPerf,m1GTxPwrTest_P4,m1GTxPwrPerfTest_P4, \
                    m10GTxPwrTest_P4,m10GTxPwrPerfTest_P4))
                    con2 = pymssql.connect(self.dbip, self.dbuser, self.dbpassword, self.dbname)
                    sql3 = "INSERT INTO %s (SN, ToTime, IsOk, mATEName, m1GOverLoadPerf, m10GOverLoadPerf, m1GSENSPerf, \
                    m10GSENSPerf, m1GTxPwrTest_P4, m1GTxPwrPerfTest_P4, m10GTxPwrTest_P4, m10GTxPwrPerfTest_P4) VALUES \
                    ('%s', '%s', '%s', '%s', %.2f, %.2f, %.2f, %.2f, %.2f, %.2f, %.2f, %.2f)" % (self.table, result[0][0], \
                    result[0][1], result[0][2], mATEName, float(m1GOverLoadPerf), float(m10GOverLoadPerf),float(m1GSENSPerf), \
                    float(m10GSENSPerf), float(m1GTxPwrTest_P4), float(m1GTxPwrPerfTest_P4), float(m10GTxPwrTest_P4),float(m10GTxPwrPerfTest_P4))
                    cur2 = con2.cursor()
                    self.logger.info(sql3)
                    cur2.execute(sql3)
                    con2.commit()

                    # 插入数据之后更新状态为1
                    sql4 = "update Result1 set Reserved5 = '1' where BarCode = '%s'" % result[0][0]
                    cur.execute(sql4)
                    con.commit()

                    cur.close()
                    con.close()
                    cur2.close()
                    con2.close()

                else:
                    pass
                self.x += 1
                self.j = 0
                self.messagess = '上传%s数据成功第%d次' % (self.pid, self.x)
                self.logger.info('上传数据成功第%d次' % self.x)
                self.label.setText(self.messagess)
            else:
                self.label.setText('当前无需要上传数据')
        except:
            self.messagess = '断线重连接第%d次' % self.j
            self.logger.error('断线重连接第%d次' % self.j)
            self.j += 1
            self.x = 0
        global timer
        timer = threading.Timer(2, self.insert_p4_data)
        timer.start()

    def insert_ip50gp0p1_data(self):
        try:
            #本地mdb路劲可以在configuer.ini文件中配置
            con = pypyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb)};DBQ=%s' % self.mdb_path)
            cur = con.cursor()
            sql = "select top 1 BarCode, ToTime, IsOk, R1_GUID, FromTime, OperatorID, ATE_Name from Result1 where Reserved5 = '' order by R1_GUID"
            cur.execute(sql)
            result = cur.fetchall()
            self.logger.info('表1查询数据记录：%s' % result)

            if len(result) == 1:
                con2 = pymssql.connect(self.dbip, self.dbuser, self.dbpassword, self.dbname)
                sql3 = "INSERT into IP50G_General (SN, ATEName, FromTime, ToTime, IsOk, OperatorID) values ('{0}', '{1}', '{2}', \
                '{3}', '{4}', '{5}')".format(result[0][0], result[0][6], result[0][4], result[0][1], result[0][2], result[0][5])
                cur2 = con2.cursor()
                self.logger.info(sql3)
                cur2.execute(sql3)
                con2.commit()

                #插入数据之后更新状态为1
                sql4 = "update Result1 set Reserved5 = '1' where BarCode = '%s'" % result[0][0]
                cur.execute(sql4)
                con.commit()

                cur.close()
                con.close()
                cur2.close()
                con2.close()

                self.x += 1
                self.j = 0
                self.messagess = '上传%s数据成功第%d次' % (self.pid, self.x)
                self.logger.info('上传数据成功第%d次' % self.x)
                self.label.setText(self.messagess)
            else:
                self.label.setText('当前无需要上传数据')
        except:
            self.messagess = '断线重连接第%d次' % self.j
            self.logger.error('断线重连接第%d次' % self.j)
            self.j += 1
            self.x = 0
        global timer
        timer = threading.Timer(2, self.insert_ip50gp0p1_data)
        timer.start()

    def insert_ip50g_mp1_data(self):        #标定
        try:
            #本地mdb路劲可以在configuer.ini文件中配置
            con = pypyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb)};DBQ=%s' % self.mdb_path)
            cur = con.cursor()
            sql = "select top 1 BarCode, ToTime, IsOk, R1_GUID, FromTime, OperatorID, ATE_Name from Result1 where Reserved5 = '' ORDER BY Result1ID"
            cur.execute(sql)
            result = cur.fetchall()
            self.logger.info('表1查询数据记录：%s' % result)
            if len(result) == 1:
                value_list = [result[0][0], result[0][6], str(result[0][4]), str(result[0][1]), result[0][2], result[0][5]]
                # 同一个sn（Result1ID）有多条记录的，取后面一条  order by Result3ID DESC
                sql2 = "select SubItemName,ResultDesc from Result3 where R1_GUID = '%s' order by Result3ID DESC " % result[0][3]
                cur.execute(sql2)
                result2 = cur.fetchall()

                if len(result2) > 0:
                    #给待插入的值赋初值
                    mPowerSupply = 0    #模块供电电压(V)
                    msupplycurrent = 0  #模块供电电流(A)
                    mIDD = 0            #IDD电流
                    mPwr1dBm = 0        #Pwr1(dBm)
                    mPwr1AD = 0         #Pwr1AD
                    mPwr2dBm = 0        #Pwr2(dBm)
                    mPwr2AD = 0         #Pwr2AD
                    mTxReportfk = 0     #发端上报fk
                    mTxReportfb = 0     #发端上报fb
                    mTxReportdwk = 0        #发端上报dwk
                    mTxReportdwb = 0        #发端上报dwb
                    mTosaFactory = 'YSOD'   #TOSA厂家
                    mRosaFactory = 'YSOD'   #ROSA厂家
                    mRxNoPwrAd = 0      #RxNoPwrAd
                    mRxPwrCt0Ad = 0     #RxPwrCt0Ad
                    mRxPwrCt1Ad = 0     #RxPwrCt1Ad
                    mRxPwrCt2Ad = 0     #RxPwrCt2Ad
                    mRxPwrCt3Ad = 0     #RxPwrCt3Ad
                    mRxPwrCt4Ad = 0     #RxPwrCt4Ad
                    mRxPwrLo0Ad = 0     #RxPwrLo0Ad
                    mRxPwrLo1Ad = 0     #RxPwrLo1Ad
                    mMidlightfa = 0     #中光曲线系数fa
                    mMidlightfb = 0     #中光曲线系数fb
                    mMidlightfc = 0     #中光曲线系数fc
                    mMidlightfd = 0     #中光曲线系数fd
                    mMidlightfe = 0     #中光曲线系数fe
                    mMinlightfk = 0     #小光曲线系数fk
                    mMinlightfb = 0     #小光直线系数fb
                    mFTTxPower = 0          #FT发端实际功率
                    mFTTxReport = 0         #FT发端AD上报
                    mFTTxReportPower = 0    #FT发端上报功率
                    mFTEX = 0               #FT消光比
                    mAOPSEN0 = 0        #FT 1.000000E-004 AOP灵敏度点0
                    mOMASEN0 = 0        #FT 1.000000E-004 OMA灵敏度点0
                    mAOPSEN1 = 0        #FT 1.000000E-004 AOP灵敏度点1
                    mOMASEN1 = 0        #FT 1.000000E-004 OMA灵敏度点1
                    mRxReport = 0       #FT收端无光上报
                    mRxAD = 0           #FT收端无光AD
                    mTxNoLightPower = 0     #FT发端无光功率
                    mFTBias = 0                 #FT偏流
                    value_list2 = [mPowerSupply, msupplycurrent, mIDD, mPwr1dBm, mPwr1AD, mPwr2dBm, mPwr2AD, mTxReportfk, mTxReportfb, \
                                   mTxReportdwk, mTxReportdwb, mTosaFactory, mRosaFactory, mRxNoPwrAd, mRxPwrCt0Ad, mRxPwrCt1Ad, \
                                   mRxPwrCt2Ad, mRxPwrCt3Ad, mRxPwrCt4Ad, mRxPwrLo0Ad, mRxPwrLo1Ad, mMidlightfa, mMidlightfb, \
                                   mMidlightfc, mMidlightfd, mMidlightfe, mMinlightfk, mMinlightfb, mFTTxPower, mFTTxReport, \
                                   mFTTxReportPower, mFTEX, mAOPSEN0, mOMASEN0, mAOPSEN1, mOMASEN1, mRxReport, mRxAD, mTxNoLightPower, \
                                   mFTBias]
                    value_list = value_list + value_list2

                    #赋值待插入的值
                    for i in range(len(result2)):
                        if result2[i][0] == '模块供电电压(V)':
                            mPowerSupply = result2[i][1]
                            value_list[6] = mPowerSupply
                        elif result2[i][0] == '模块供电电流(A)':
                            msupplycurrent = result2[i][1]
                            value_list[7] = msupplycurrent
                        elif result2[i][0] == 'IDD电流':
                            mIDD = result2[i][1]
                            value_list[8] = mIDD
                        elif result2[i][0] == 'Pwr1(dBm)':
                            mPwr1dBm = result2[i][1]
                            value_list[9] = mPwr1dBm
                        elif result2[i][0] == 'Pwr1AD':
                            mPwr1AD = result2[i][1]
                            value_list[10] = mPwr1AD
                        elif result2[i][0] == 'Pwr2(dBm)':
                            mPwr2dBm = result2[i][1]
                            value_list[11] = mPwr2dBm
                        elif result2[i][0] == 'Pwr2AD':
                            mPwr2AD = result2[i][1]
                            value_list[12] = mPwr2AD
                        elif result2[i][0] == '发端上报fk':
                            mTxReportfk = result2[i][1]
                            value_list[13] = mTxReportfk
                        elif result2[i][0] == '发端上报fb':
                            mTxReportfb = result2[i][1]
                            value_list[14] = mTxReportfb
                        elif result2[i][0] == '发端上报dwk':
                            mTxReportdwk = result2[i][1]
                            value_list[15] = mTxReportdwk
                        elif result2[i][0] == '发端上报dwb':
                            mTxReportdwb = result2[i][1]
                            value_list[16] = mTxReportdwb
                        elif result2[i][0] == 'TOSA厂家':
                            mTosaFactory = result2[i][1]
                            value_list[17] = mTosaFactory
                        elif result2[i][0] == 'ROSA厂家':
                            mRosaFactory = result2[i][1]
                            value_list[18] = mRosaFactory
                        elif result2[i][0] == 'RxNoPwrAd':
                            mRxNoPwrAd = result2[i][1]
                            value_list[19] = mRxNoPwrAd
                        elif result2[i][0] == 'RxPwrCt0Ad':
                            mRxPwrCt0Ad = result2[i][1]
                            value_list[20] = mRxPwrCt0Ad

                        elif result2[i][0] == 'RxPwrCt1Ad':
                            mRxPwrCt1Ad = result2[i][1]
                            value_list[6] = mPowerSupply
                        elif result2[i][0] == 'RxPwrCt2Ad':
                            mRxPwrCt2Ad = result2[i][1]
                            value_list[21] = mRxPwrCt2Ad
                        elif result2[i][0] == 'RxPwrCt3Ad':
                            mRxPwrCt3Ad = result2[i][1]
                            value_list[22] = mRxPwrCt3Ad
                        elif result2[i][0] == 'RxPwrCt4Ad':
                            mRxPwrCt4Ad = result2[i][1]
                            value_list[23] = mRxPwrCt4Ad
                        elif result2[i][0] == 'RxPwrLo0Ad':
                            mRxPwrLo0Ad = result2[i][1]
                            value_list[24] = mRxPwrLo0Ad
                        elif result2[i][0] == 'RxPwrLo1Ad':
                            mRxPwrLo1Ad = result2[i][1]
                            value_list[25] = mRxPwrLo1Ad
                        elif result2[i][0] == '中光曲线系数fa':
                            mMidlightfa = result2[i][1]
                            value_list[26] = mMidlightfa
                        elif result2[i][0] == '中光曲线系数fb':
                            mMidlightfb = result2[i][1]
                            value_list[27] = mMidlightfb
                        elif result2[i][0] == '中光曲线系数fc':
                            mMidlightfc = result2[i][1]
                            value_list[28] = mMidlightfc
                        elif result2[i][0] == '中光曲线系数fd':
                            mMidlightfd = result2[i][1]
                            value_list[29] = mMidlightfd
                        elif result2[i][0] == '中光曲线系数fe':
                            mMidlightfe = result2[i][1]
                            value_list[30] = mMidlightfe
                        elif result2[i][0] == '小光曲线系数fk':
                            mMinlightfk = result2[i][1]
                            value_list[31] = mMinlightfk
                        elif result2[i][0] == '小光直线系数fb':
                            mMinlightfb = result2[i][1]
                            value_list[32] = mMinlightfb
                        elif result2[i][0] == 'FT发端实际功率':
                            mFTTxPower = result2[i][1]
                            value_list[33] = mFTTxPower
                        elif result2[i][0] == 'FT发端AD上报':
                            mFTTxReport = result2[i][1]
                            value_list[34] = mFTTxReport
                        elif result2[i][0] == 'FT发端上报功率':
                            mFTTxReportPower = result2[i][1]
                            value_list[35] = mFTTxReportPower
                        elif result2[i][0] == 'FT消光比':
                            mFTEX = result2[i][1]
                            value_list[36] = mFTEX
                        elif result2[i][0] == 'FT 1.000000E-004 AOP灵敏度点0':
                            mAOPSEN0 = result2[i][1]
                            value_list[37] = mAOPSEN0
                        elif result2[i][0] == 'FT 1.000000E-004 OMA灵敏度点0':
                            mOMASEN0 = result2[i][1]
                            value_list[38] = mOMASEN0
                        elif result2[i][0] == 'FT 1.000000E-004 AOP灵敏度点1':
                            mAOPSEN1 = result2[i][1]
                            value_list[39] = mAOPSEN1
                        elif result2[i][0] == 'FT 1.000000E-004 OMA灵敏度点1':
                            mOMASEN1 = result2[i][1]
                            value_list[40] = mOMASEN1
                        elif result2[i][0] == 'FT收端无光上报':
                            mRxReport = result2[i][1]
                            value_list[41] = mRxReport
                        elif result2[i][0] == 'FT收端无光AD':
                            mRxAD = result2[i][1]
                            value_list[42] = mRxAD
                        elif result2[i][0] == 'FT发端无光功率':
                            mTxNoLightPower = result2[i][1]
                            value_list[43] = mTxNoLightPower
                        elif result2[i][0] == 'FT偏流':
                            mFTBias = result2[i][1]
                            value_list[44] = mFTBias

                    self.logger.info(value_list)
                    value_list_str = ','.join(map(lambda x: "'" + str(x) + "'", value_list))
                    self.logger.warning(value_list_str)
                    con2 = pymssql.connect(self.dbip, self.dbuser, self.dbpassword, self.dbname)
                    sql3 = "INSERT into IP50G_P1 VALUES ({})".format(value_list_str)
                    cur2 = con2.cursor()
                    self.logger.info(sql3)
                    cur2.execute(sql3)
                    con2.commit()

                    #插入数据之后更新状态为1
                    sql4 = "update Result1 set Reserved5 = '1' where BarCode = '%s'" % result[0][0]
                    cur.execute(sql4)
                    con.commit()

                    cur.close()
                    con.close()
                    cur2.close()
                    con2.close()
                else:
                    pass
                self.x += 1
                self.j = 0
                self.messagess = '上传%s数据成功第%d次' % (self.pid, self.x)
                self.logger.info('上传数据成功第%d次' % self.x)
                self.label.setText(self.messagess)
            else:
                self.label.setText('当前无需要上传数据')
        except:
            self.messagess = '断线重连接第%d次' % self.j
            self.logger.error('断线重连接第%d次' % self.j)
            self.j += 1
            self.x = 0
        global timer
        timer = threading.Timer(2, self.insert_ip50g_mp1_data)
        timer.start()

    def insert_ip50g_mp2_data(self):        #三温测试
        try:
            #本地mdb路劲可以在configuer.ini文件中配置
            con = pypyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb)};DBQ=%s' % self.mdb_path)
            cur = con.cursor()
            sql = "select top 1 BarCode, ToTime, IsOk, R1_GUID, FromTime, OperatorID, ATE_Name from Result1 where Reserved5 = '' ORDER BY Result1ID"
            cur.execute(sql)
            result = cur.fetchall()
            self.logger.info('表1查询数据记录：%s' % result)
            if len(result) == 1:
                value_list = [result[0][0], result[0][6], str(result[0][4]), str(result[0][1]), result[0][2], result[0][5]]
                # 同一个sn（Result1ID）有多条记录的，取后面一条  order by Result3ID DESC
                sql2 = "select SubItemName,ResultDesc from Result3 where R1_GUID = '%s' order by Result3ID DESC " % result[0][3]
                cur.execute(sql2)
                result2 = cur.fetchall()

                if len(result2) > 0:
                    #给待插入的值赋初值
                    mPowerSupply_LH = 0         #Lt@FT1-MP2@模块供电电压(V)
                    msupplycurrent_LH = 0       #Lt@FT1-MP2@模块供电电流(A)
                    mTxPower_LH = 0             #Lt@FT1-MP2@低温发端实际功率
                    mTxADReport_LH = 0             #Lt@FT1-MP2@低温发端AD上报
                    mTxReportPower_LH = 0          #Lt@FT1-MP2@低温发端上报功率
                    mEX_LH = 0                  #Lt@FT1-MP2@低温消光比
                    mAOPSEN0_LH = 0                #Lt@FT1-MP2@低温 1.000000E-004 AOP灵敏度点0
                    mOMASEN0_LH = 0                #Lt@FT1-MP2@低温 1.000000E-004 OMA灵敏度点0
                    mLOS_LH = 0                 #Lt@FT1-MP2@低温LOS建立点
                    mRxNoLightReport_LH = 0     #Lt@FT1-MP2@低温收端无光上报
                    mRxNoLightAD_LH = 0            #Lt@FT1-MP2@低温收端无光AD
                    mRxReportP0_LH = 0             #Lt@FT1-MP2@低温收端上报功率点0
                    mRxReportP1_LH = 0             #Lt@FT1-MP2@低温收端上报功率点1
                    mRxReportP2_LH = 0             #Lt@FT1-MP2@低温收端上报功率点2
                    mRxReportP3_LH = 0             #Lt@FT1-MP2@低温收端上报功率点3
                    mRxReportP4_LH = 0             #Lt@FT1-MP2@低温收端上报功率点4
                    mTxNoLightPower_LH = 0         #Lt@FT1-MP2@低温发端无光功率
                    mFTBias_LH = 0              #Lt@FT1-MP2@低温偏流
                    mPowerSupply_H = 0          #Ht@FT1-MP2@模块供电电压(V)
                    msupplycurrent_H = 0        #Ht@FT1-MP2@模块供电电流(A)
                    mTxPower_H = 0              #Ht@FT1-MP2@高温发端实际功率
                    mTxADReport_H = 0           #Ht@FT1-MP2@高温发端AD上报
                    mTxReportPower_H = 0        #Ht@FT1-MP2@高温发端上报功率
                    mEX_H = 0                   #Ht@FT1-MP2@高温消光比
                    mAOPSEN0_H = 0              #Ht@FT1-MP2@高温 1.000000E-004 AOP灵敏度点0
                    mOMASEN0_H = 0              #Ht@FT1-MP2@高温 1.000000E-004 OMA灵敏度点0
                    mAOPSEN1_H = 0              #Ht@FT1-MP2@高温 1.000000E-004 AOP灵敏度点1
                    mOMASEN1_H = 0              #Ht@FT1-MP2@高温 1.000000E-004 OMA灵敏度点1
                    mLOS_H = 0                  #Ht@FT1-MP2@高温LOS建立点
                    mRxNoLightReport_H = 0      #Ht@FT1-MP2@高温收端无光上报
                    mRxNoLightAD_H = 0            #Ht@FT1-MP2@高温收端无光AD
                    mRxReportP0_H = 0           #Ht@FT1-MP2@高温收端上报功率点0
                    mRxReportP1_H = 0           #Ht@FT1-MP2@高温收端上报功率点1
                    mRxReportP2_H = 0           #Ht@FT1-MP2@高温收端上报功率点2
                    mRxReportP3_H = 0           #Ht@FT1-MP2@高温收端上报功率点3
                    mRxReportP4_H = 0           #Ht@FT1-MP2@高温收端上报功率点4
                    mTxNoLightPower_H = 0       #Ht@FT1-MP2@高温发端无光功率
                    mFTBias_H = 0               #Ht@FT1-MP2@高温偏流

                    value_list2 = [mPowerSupply_LH, msupplycurrent_LH, mTxPower_LH, mTxADReport_LH, mTxReportPower_LH, mEX_LH, \
                                   mAOPSEN0_LH, mOMASEN0_LH, mLOS_LH, mRxNoLightReport_LH, mRxNoLightAD_LH, mRxReportP0_LH, \
                                   mRxReportP1_LH, mRxReportP2_LH, mRxReportP3_LH, mRxReportP4_LH, mTxNoLightPower_LH, mFTBias_LH, \
                                   mPowerSupply_H, msupplycurrent_H, mTxPower_H, mTxADReport_H, mTxReportPower_H, mEX_H, mAOPSEN0_H, \
                                   mOMASEN0_H, mAOPSEN1_H, mOMASEN1_H, mLOS_H, mRxNoLightReport_H, mRxNoLightAD_H, mRxReportP0_H, \
                                   mRxReportP1_H, mRxReportP2_H, mRxReportP3_H, mRxReportP4_H, mTxNoLightPower_H, mFTBias_H]
                    value_list = value_list + value_list2

                    for i in range(len(result2)):
                        if result2[i][0] == 'Lt@FT1-MP2@模块供电电压(V)':
                            value_list[6] = result2[i][1]
                        elif result2[i][0] == 'Lt@FT1-MP2@模块供电电流(A)':
                            value_list[7] = result2[i][1]
                        elif result2[i][0] == 'Lt@FT1-MP2@低温发端实际功率':
                            value_list[8] = result2[i][1]
                        elif result2[i][0] == 'Lt@FT1-MP2@低温发端AD上报':
                            value_list[9] = result2[i][1]
                        elif result2[i][0] == 'Lt@FT1-MP2@低温发端上报功率':
                            value_list[10] = result2[i][1]
                        elif result2[i][0] == 'Lt@FT1-MP2@低温消光比':
                            value_list[11] = result2[i][1]
                        elif result2[i][0] == 'Lt@FT1-MP2@低温 1.000000E-004 AOP灵敏度点0':
                            value_list[12] = result2[i][1]
                        elif result2[i][0] == 'Lt@FT1-MP2@低温 1.000000E-004 OMA灵敏度点0':
                            value_list[13] = result2[i][1]
                        elif result2[i][0] == 'Lt@FT1-MP2@低温LOS建立点':
                            value_list[14] = result2[i][1]
                        elif result2[i][0] == 'Lt@FT1-MP2@低温收端无光上报':
                            value_list[15] = result2[i][1]
                        elif result2[i][0] == 'Lt@FT1-MP2@低温收端无光AD':
                            value_list[16] = result2[i][1]
                        elif result2[i][0] == 'Lt@FT1-MP2@低温收端上报功率点0':
                            value_list[17] = result2[i][1]
                        elif result2[i][0] == 'Lt@FT1-MP2@低温收端上报功率点1':
                            value_list[18] = result2[i][1]
                        elif result2[i][0] == 'Lt@FT1-MP2@低温收端上报功率点2':
                            value_list[19] = result2[i][1]
                        elif result2[i][0] == 'Lt@FT1-MP2@低温收端上报功率点3':
                            value_list[20] = result2[i][1]
                        elif result2[i][0] == 'Lt@FT1-MP2@低温收端上报功率点4':
                            value_list[21] = result2[i][1]
                        elif result2[i][0] == 'Lt@FT1-MP2@低温发端无光功率':
                            value_list[22] = result2[i][1]
                        elif result2[i][0] == 'Lt@FT1-MP2@低温偏流':
                            value_list[23] = result2[i][1]
                        elif result2[i][0] == 'Ht@FT1-MP2@模块供电电压(V)':
                            value_list[24] = result2[i][1]
                        elif result2[i][0] == 'Ht@FT1-MP2@模块供电电流(A)':
                            value_list[25] = result2[i][1]
                        elif result2[i][0] == 'Ht@FT1-MP2@高温发端实际功率':
                            value_list[26] = result2[i][1]
                        elif result2[i][0] == 'Ht@FT1-MP2@高温发端AD上报':
                            value_list[27] = result2[i][1]
                        elif result2[i][0] == 'Ht@FT1-MP2@高温发端上报功率':
                            value_list[28] = result2[i][1]
                        elif result2[i][0] == 'Ht@FT1-MP2@高温消光比':
                            value_list[29] = result2[i][1]
                        elif result2[i][0] == 'Ht@FT1-MP2@高温 1.000000E-004 AOP灵敏度点0':
                            value_list[30] = result2[i][1]
                        elif result2[i][0] == 'Ht@FT1-MP2@高温 1.000000E-004 OMA灵敏度点0':
                            value_list[31] = result2[i][1]
                        elif result2[i][0] == 'Ht@FT1-MP2@高温 1.000000E-004 AOP灵敏度点1':
                            value_list[32] = result2[i][1]
                        elif result2[i][0] == 'Ht@FT1-MP2@高温 1.000000E-004 OMA灵敏度点1':
                            value_list[33] = result2[i][1]
                        elif result2[i][0] == 'Ht@FT1-MP2@高温LOS建立点':
                            value_list[34] = result2[i][1]
                        elif result2[i][0] == 'Ht@FT1-MP2@高温收端无光上报':
                            value_list[35] = result2[i][1]
                        elif result2[i][0] == 'Ht@FT1-MP2@高温收端无光AD':
                            value_list[36] = result2[i][1]
                        elif result2[i][0] == 'Ht@FT1-MP2@高温收端上报功率点0':
                            value_list[37] = result2[i][1]
                        elif result2[i][0] == 'Ht@FT1-MP2@高温收端上报功率点1':
                            value_list[38] = result2[i][1]
                        elif result2[i][0] == 'Ht@FT1-MP2@高温收端上报功率点2':
                            value_list[39] = result2[i][1]
                        elif result2[i][0] == 'Ht@FT1-MP2@高温收端上报功率点3':
                            value_list[40] = result2[i][1]
                        elif result2[i][0] == 'Ht@FT1-MP2@高温收端上报功率点4':
                            value_list[41] = result2[i][1]
                        elif result2[i][0] == 'Ht@FT1-MP2@高温发端无光功率':
                            value_list[42] = result2[i][1]
                        elif result2[i][0] == 'Ht@FT1-MP2@高温偏流':
                            value_list[43] = result2[i][1]

                    self.logger.warning(value_list)

                    value_list_str = ','.join(map(lambda x: "'" + str(x) + "'", value_list))
                    self.logger.warning(len(value_list))

                    con2 = pymssql.connect(self.dbip, self.dbuser, self.dbpassword, self.dbname)
                    sql3 = "INSERT into IP50G_P2 VALUES ({})".format(value_list_str)
                    cur2 = con2.cursor()
                    self.logger.info(sql3)
                    cur2.execute(sql3)
                    con2.commit()

                    #插入数据之后更新状态为1
                    sql4 = "update Result1 set Reserved5 = '1' where BarCode = '%s'" % result[0][0]
                    cur.execute(sql4)
                    con.commit()

                    cur.close()
                    con.close()
                    cur2.close()
                    con2.close()
                else:
                    pass
                self.x += 1
                self.j = 0
                self.messagess = '上传%s数据成功第%d次' % (self.pid, self.x)
                self.logger.info('上传数据成功第%d次' % self.x)
                self.label.setText(self.messagess)
            else:
                self.label.setText('当前无需要上传数据')
        except:
            self.messagess = '断线重连接第%d次' % self.j
            self.logger.error('断线重连接第%d次' % self.j)
            self.j += 1
            self.x = 0
        global timer
        timer = threading.Timer(2, self.insert_ip50g_mp2_data)
        timer.start()

    def insert_ip50g_mp4_data(self):        #回损
        try:
            #本地mdb路劲可以在configuer.ini文件中配置
            con = pypyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb)};DBQ=%s' % self.mdb_path)
            cur = con.cursor()
            sql = "select top 1 BarCode, ToTime, IsOk, R1_GUID, FromTime, OperatorID, ATE_Name from Result1 where Reserved5 = '' ORDER BY Result1ID"
            cur.execute(sql)
            result = cur.fetchall()
            self.logger.info('表1查询数据记录：%s' % result)
            if len(result) == 1:
                value_list = [result[0][0], result[0][6], str(result[0][4]), str(result[0][1]), result[0][2], result[0][5]]
                # 同一个sn（Result1ID）有多条记录的，取后面一条  order by Result3ID DESC
                sql2 = "select SubItemName,ResultDesc from Result3 where R1_GUID = '%s' order by Result3ID DESC " % result[0][3]
                cur.execute(sql2)
                result2 = cur.fetchall()

                if len(result2) > 0:
                    #给待插入的值赋初值
                    mPowerSupply = 0                #回损工位模块电压
                    msupplycurrent = 0              #回损工位模块电流
                    mdissPower = 0                  #回损工位模块功耗
                    mTxReturnLoss = 0               #发端回损
                    mRxReturnLoss = 0               #收端回损

                    value_list2 = [mPowerSupply, msupplycurrent, mdissPower, mTxReturnLoss, mRxReturnLoss]
                    value_list = value_list + value_list2

                    for i in range(len(result2)):
                        if result2[i][0] == '回损工位模块电压':
                            value_list[6] = result2[i][1]
                        elif result2[i][0] == '回损工位模块电流':
                            value_list[7] = result2[i][1]
                        elif result2[i][0] == '回损工位模块功耗':
                            value_list[8] = result2[i][1]
                        elif result2[i][0] == '发端回损':
                            value_list[9] = result2[i][1]
                        elif result2[i][0] == '收端回损':
                            value_list[10] = result2[i][1]

                    value_list_str = ','.join(map(lambda x: "'" + str(x) + "'", value_list))

                    con2 = pymssql.connect(self.dbip, self.dbuser, self.dbpassword, self.dbname)
                    sql3 = "INSERT into IP50G_P4 VALUES ({})".format(value_list_str)
                    cur2 = con2.cursor()
                    self.logger.info(sql3)
                    cur2.execute(sql3)
                    con2.commit()

                    #插入数据之后更新状态为1
                    sql4 = "update Result1 set Reserved5 = '1' where BarCode = '%s'" % result[0][0]
                    cur.execute(sql4)
                    con.commit()

                    cur.close()
                    con.close()
                    cur2.close()
                    con2.close()

                    self.x += 1
                    self.j = 0
                    self.messagess = '上传%s数据成功第%d次' % (self.pid, self.x)
                    self.logger.info('上传数据成功第%d次' % self.x)
                    self.label.setText(self.messagess)
                else:
                    self.y += 1
                    self.messagess = '当前SN在表3中无测试记录, %s' % self.y
                    self.logger.warning('上传数据成功第%d次' % self.y)
                    self.label.setText(self.messagess)

            else:
                self.label.setText('当前无需要上传数据')
        except:
            self.messagess = '断线重连接第%d次' % self.j
            self.logger.error('断线重连接第%d次' % self.j)
            self.j += 1
            self.x = 0
        global timer
        timer = threading.Timer(2, self.insert_ip50g_mp4_data)
        timer.start()

class MytrayIcon(QSystemTrayIcon):
    def __init__(self, parent):
        super(MytrayIcon, self).__init__(parent)
        self.setIcon(QIcon('load.ico'))
        self.bar()

    def bar(self):
        self.menu = QMenu()
        self.action1 = QAction('显示', self)
        self.action2 = QAction('退出', self)
        self.action3 = QAction('查询', self)
        self.action1.triggered.connect(self.uishow)
        self.action2.triggered.connect(self.appQuit)
        self.action3.triggered.connect(self.dataselect)
        self.menu.addAction(self.action1)
        self.menu.addAction(self.action2)
        self.menu.addAction(self.action3)
        self.setContextMenu(self.menu)

    def uishow(self):
        self.parent().show()

    def appQuit(self):          #如何优雅的退出？？
        self.hide()
        # self.data_ui.close()
        # QCoreApplication.instance().quit()
        # print(os.getpid())
        os.kill(os.getpid(), signal.SIGILL)

    def dataselect(self):
        self.data_ui = SQLdataselect()
        self.data_ui.show()

class SQLdataselect(QWidget, sqldataselect.Ui_Form):
    def __init__(self):
        super(SQLdataselect, self).__init__()
        self.setupUi(self)
        self.setWindowTitle('测试性能数据查询V1.0.5')
        self.setWindowIcon(QIcon('load.ico'))
        self.data_num = 0
        now_date = time.strftime('%Y-%m-%d', time.localtime())
        now_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
        self.dateTimeEdit_start.setDateTime(QDateTime.fromString(now_date, 'yyyy-MM-dd'))
        self.dateTimeEdit_end.setDateTime(QDateTime.fromString(now_time, 'yyyy-MM-dd hh:mm:ss'))

        #连接数据库
        try:
           database_config  = ConfigParser()
           database_config.read(r'.\configuer.ini')
           self.username = database_config.get('DB', 'dbuser')
           self.password = database_config.get('DB', 'dbpassword')
           self.database = database_config.get('DB', 'dbip')
           self.conn = pymssql.connect('%s' % self.database, '%s' % self.username, '%s' % self.password, 'oltDB')
           self.cur = self.conn.cursor()
        except:
            print('1111111')

        self.pushButton.clicked.connect(self.data_select)
        self.pushButton_load.clicked.connect(self.data_load)

    def closeEvent(self, event):
        event.ignore()
        self.hide()
        self.cur.close()
        self.conn.close()

    def data_select(self):          #按时间段查询
        self.sn_old = self.lineEdit_sn.text()
        if len(self.sn_old) < 20 and len(self.sn_old) != 16:
            self.sn = ''
        elif len(self.sn_old) == 16:
            self.sn = self.sn_old
        else:
            self.sn = self.sn_old[-20:-4]
        self.lineEdit_sn.setText(self.sn)

        self.time_start = self.dateTimeEdit_start.text()
        self.time_end = self.dateTimeEdit_end.text()
        if self.comboBox.currentIndex() == 0:
            self.table_name = 'OM5256_P1'
            self.columns_list = ['SN',
            'ToTime',
            'IsOk',
            'ATEName',
            'EleVal',
            'VolVal',
            'PowVal',
            'ModuleTemp',
            'EleVal_F',
            'VolVal_F',
            'PowVal_F',
            '10GBias',
            '1GBias']
        elif self.comboBox.currentIndex() == 1:
            self.table_name = 'OM5256_P2'
            self.columns_list = ['SN',
            'ToTime',
            'IsOk',
            'ATEName',
            'm10GTXExtinDA',
            'm10GTXPWR_dBm',
            'm10GTXExtin_dB',
            'm10GTXCross',
            'm10GSens',
            'm10GSensAfter',
            'm1GTXExtinDA',
            'm1GTXPWR_dBm',
            'm1GTXExtin_dB',
            'm1GTXCross',
            'm10GEyeCross',
            'm10GEyeExtin_dB',
            'm10GDSAPwr_dBm',
            'm1GEyeCross',
            'm1GEyeExtin_dB',
            'm1GDSAPwr_dBm']
        elif self.comboBox.currentIndex() == 2:
            self.table_name = 'OM5256_P3'
            self.columns_list = ['SN',
            'ToTime',
            'IsOk',
            'mATEName',
            'mAPDVol',
            'm10GTxNoPwrAD',
            'mAPDVolTest',
            'm1GTxPwrTest_P3',
            'm1GTxPwrPerfTest_P3',
            'm10GTxPwrTest_P3',
            'm10GTxPwrPerfTest_P3',
            'm10GTxPwrTest2_P3',
            'm10GTxPwrPerfTest2_P3',
            'm10GBias_P3',
            'm1GBias_P3']
        elif self.comboBox.currentIndex() == 3:
            self.table_name = 'OM5256_P4'
            self.columns_list = ['SN',
            'ToTime',
            'IsOk',
            'mATEName',
            'm1GOverLoadPerf',
            'm10GOverLoadPerf',
            'm1GSENSPerf',
            'm10GSENSPerf',
            'm1GTxPwrTest_P4',
            'm1GTxPwrPerfTest_P4',
            'm10GTxPwrTest_P4',
            'm10GTxPwrPerfTest_P4']
        else:
            QMessageBox.information(self, '提醒', '因为每个工序指标不一样，如全部查询字段太多，请提供需要查看的字段！')
            return 0
        if len(self.sn) > 0:
            #按单支查询
            sql1 = "select * from %s where SN = '%s'" % (self.table_name, self.sn)
            self.cur.execute(sql1)
        else:

            # 按时间段查询
            sql2 = "select * from %s where ToTime > '%s' and ToTime < '%s'" % (self.table_name, self.time_start, self.time_end)
            self.cur.execute(sql2)
        self.data_result = self.cur.fetchall()
        self.data_num = self.cur.rowcount
        self.tableWidget.setColumnCount(len(self.columns_list))
        self.tableWidget.setRowCount(self.data_num)
        self.tableWidget.setHorizontalHeaderLabels(self.columns_list)
        for i in range(len(self.data_result)):
            for j in range(len(self.columns_list)):
                if j == 1:
                    self.tableWidget.setItem(i, j, QTableWidgetItem(str(self.data_result[i][j][:-8])))
                else:
                    self.tableWidget.setItem(i, j ,QTableWidgetItem(str(self.data_result[i][j])))

    def data_load(self):
        if self.data_num > 0:
            wbk = xlwt.Workbook()
            sheet = wbk.add_sheet('sheet1', cell_overwrite_ok=True)
            for m in range(len(self.columns_list)):
                sheet.write(0,m, self.columns_list[m])

            for x in range(self.data_num):
                for y in range(len(self.columns_list)):
                    if y == 1:
                        sheet.write(x+1, y, str(self.data_result[x][y][:-8]))
                    else:
                        sheet.write(x+1, y, str(self.data_result[x][y]))

            cur_time = time.strftime("%Y%m%d %H%M%S", time.localtime())

            path = pathlib.Path('.\\导出数据')
            if path.exists() == False:
                os.makedirs('.\\导出数据')
            wbk.save('.\\导出数据\\' + '导出数据-%s.xls' % cur_time)
            QMessageBox.information(self, '提示', '导出数据成功！')

    def data_IP50G_select(self):
        pass

    def data_IP50G_load(self):
        pass


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ui = Mywindow()
    ui.show()
    ui.trayico.show()
    sys.exit(app.exec_())
