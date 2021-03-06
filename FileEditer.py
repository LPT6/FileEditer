# -*- coding: utf-8 -*-

"""
Module implementing MainWindow.
"""

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QMainWindow, QFileDialog
from PyQt5 import QtCore, QtWidgets
from Ui_FileEditer import Ui_MainWindow


class MainWindow(QMainWindow, Ui_MainWindow):
    """
    Class documentation goes here.
    """
    def __init__(self, parent=None):
        """
        Constructor
        
        @param parent reference to the parent widget
        @type QWidget
        """
        super(MainWindow, self).__init__(parent)
        self.setupUi(self)
    
    @pyqtSlot()
    def on_action_triggered(self):
        """
        Slot documentation goes here.
        """
        #_translate = QtCore.QCoreApplication.translate
        print('新建文件')
        my_file, file_type =  QFileDialog.getSaveFileName(self, u'新建文件', './', '*.txt')
        #self.textBrowser.setText(_translate("MainWindow", ""))
        print(my_file, file_type)
        if my_file != '':
            f=open(my_file,'w')
            f.close()
        
    
    @pyqtSlot()
    def on_action_2_triggered(self):
        """
        Slot documentation goes here.
        """
        print('打开文件')
        my_file, file_type =  QFileDialog.getOpenFileName(self, u'打开文件', './')
        print(my_file, file_type)
        if my_file[-4:] == '.doc' or my_file[-5:] == '.docx':
            from win32com import client as wc #d导入client,并给client取个别名wc--word client
            word=wc.Dispatch('Word.Application')#引用完应用程序
            word.Visible=0 #设置后台打开，不可见打开
            #打开文件
            my_worddoc=word.Documents.Open(my_file)
            my_count = my_worddoc.Paragraphs.Count
            for i in range(my_count):
                my_pr=my_worddoc.Paragraphs[i].Range
                print(my_pr)
            my_worddoc.Close()
        elif my_file[-4:] == '.txt':
            f=open(my_file)
            my_data=f.read()
            self.textBrowser.setText(my_data)
    
    @pyqtSlot()
    def on_action_3_triggered(self):
        """
        Slot documentation goes here.
        """
        print('保存文件')
        my_file, file_type =  QFileDialog.getSaveFileName(self, u'保存文件', './', '*.txt')
        my_data = self.textBrowser.toPlainText()
        print(my_file, file_type)
        if my_file != '':
            f=open(my_file,'w')
            f.write(my_data)
            f.close()
    
    @pyqtSlot()
    def on_action_4_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    
    @pyqtSlot()
    def on_action_5_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    
    @pyqtSlot()
    def on_action_6_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
        
if  __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    ui = MainWindow()
    ui.show()
    sys.exit(app.exec_())
