# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'sqldataselect.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(752, 390)
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(Form)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.gridLayout = QtWidgets.QGridLayout()
        self.gridLayout.setObjectName("gridLayout")
        self.label = QtWidgets.QLabel(Form)
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 0, 0, 1, 1)
        self.lineEdit_sn = QtWidgets.QLineEdit(Form)
        self.lineEdit_sn.setObjectName("lineEdit_sn")
        self.gridLayout.addWidget(self.lineEdit_sn, 0, 1, 1, 1)
        self.label_3 = QtWidgets.QLabel(Form)
        self.label_3.setObjectName("label_3")
        self.gridLayout.addWidget(self.label_3, 1, 2, 1, 1)
        self.dateTimeEdit_end = QtWidgets.QDateTimeEdit(Form)
        self.dateTimeEdit_end.setDateTime(QtCore.QDateTime(QtCore.QDate(2018, 1, 1), QtCore.QTime(23, 59, 59)))
        self.dateTimeEdit_end.setObjectName("dateTimeEdit_end")
        self.gridLayout.addWidget(self.dateTimeEdit_end, 1, 3, 1, 1)
        self.label_2 = QtWidgets.QLabel(Form)
        self.label_2.setObjectName("label_2")
        self.gridLayout.addWidget(self.label_2, 1, 0, 1, 1)
        self.dateTimeEdit_start = QtWidgets.QDateTimeEdit(Form)
        self.dateTimeEdit_start.setDateTime(QtCore.QDateTime(QtCore.QDate(2018, 1, 1), QtCore.QTime(0, 0, 0)))
        self.dateTimeEdit_start.setObjectName("dateTimeEdit_start")
        self.gridLayout.addWidget(self.dateTimeEdit_start, 1, 1, 1, 1)
        self.label_4 = QtWidgets.QLabel(Form)
        self.label_4.setObjectName("label_4")
        self.gridLayout.addWidget(self.label_4, 0, 2, 1, 1)
        self.comboBox = QtWidgets.QComboBox(Form)
        self.comboBox.setObjectName("comboBox")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.gridLayout.addWidget(self.comboBox, 0, 3, 1, 1)
        self.pushButton = QtWidgets.QPushButton(Form)
        self.pushButton.setObjectName("pushButton")
        self.gridLayout.addWidget(self.pushButton, 0, 4, 1, 1)
        self.pushButton_load = QtWidgets.QPushButton(Form)
        self.pushButton_load.setObjectName("pushButton_load")
        self.gridLayout.addWidget(self.pushButton_load, 1, 4, 1, 1)
        self.gridLayout.setColumnMinimumWidth(0, 1)
        self.gridLayout.setColumnMinimumWidth(1, 3)
        self.gridLayout.setColumnMinimumWidth(2, 1)
        self.gridLayout.setColumnMinimumWidth(3, 3)
        self.gridLayout.setColumnMinimumWidth(4, 3)
        self.gridLayout.setColumnStretch(0, 1)
        self.gridLayout.setColumnStretch(1, 3)
        self.gridLayout.setColumnStretch(2, 1)
        self.gridLayout.setColumnStretch(3, 3)
        self.gridLayout.setColumnStretch(4, 3)
        self.verticalLayout.addLayout(self.gridLayout)
        spacerItem = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem)
        self.tableWidget = QtWidgets.QTableWidget(Form)
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setRowCount(0)
        self.verticalLayout.addWidget(self.tableWidget)
        self.verticalLayout.setStretch(0, 2)
        self.verticalLayout.setStretch(1, 1)
        self.verticalLayout.setStretch(2, 10)
        self.verticalLayout_2.addLayout(self.verticalLayout)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.label.setText(_translate("Form", "模块SN"))
        self.label_3.setText(_translate("Form", "结束时间"))
        self.label_2.setText(_translate("Form", "起始时间"))
        self.label_4.setText(_translate("Form", "工序名称"))
        self.comboBox.setItemText(0, _translate("Form", "测试工序一"))
        self.comboBox.setItemText(1, _translate("Form", "测试工序二"))
        self.comboBox.setItemText(2, _translate("Form", "测试工序三"))
        self.comboBox.setItemText(3, _translate("Form", "测试工序四"))
        self.comboBox.setItemText(4, _translate("Form", "所有测试工序"))
        self.pushButton.setText(_translate("Form", "查询"))
        self.pushButton_load.setText(_translate("Form", "导出"))

