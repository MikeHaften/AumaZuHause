# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'e:\Programmieren\Python\AumaExe-main\Statistik.ui'
#
# Created by: PyQt5 UI code generator 5.15.7
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_StatDial(object):
    def setupUi(self, StatDial):
        StatDial.setObjectName("StatDial")
        StatDial.resize(1156, 509)
        self.gridLayout = QtWidgets.QGridLayout(StatDial)
        self.gridLayout.setObjectName("gridLayout")

        self.retranslateUi(StatDial)
        QtCore.QMetaObject.connectSlotsByName(StatDial)

    def retranslateUi(self, StatDial):
        _translate = QtCore.QCoreApplication.translate
        StatDial.setWindowTitle(_translate("StatDial", "Statistik"))
