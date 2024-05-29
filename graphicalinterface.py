



# -*- coding: utf-8 -*-
"""
Created on Sun Sep 18 11:57:50 2022

@author: hamza IŞIK
"""

import os.path
import sys
import time
import pandas as pd
from PyQt6.QtCore import QRunnable,QThreadPool,QThread,pyqtSignal,QObject
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton,QLabel,QVBoxLayout,QFileDialog,QProgressBar,QDialog,QMessageBox
from PyQt6.QtGui import QColor,QPalette,QIcon

import programBackside
import threading






class Workerone:


    def __init__(self):
        self.finished = []
        self.started = []
        self.record = []

        super().__init__()

        self.sayac = 0

    def run(self):
        self.started.append(True)
        programBackside.fiyatliste.dataListRbc()
        self.finished.append(True)
        if len(self.finished) == 6:
            window.buttonexcel.setEnabled(True)





    def runseven(self):
        self.started.append(True)
        programBackside.fiyatliste.dataListDirenc()

        self.finished.append(True)
        if len(self.finished) == 6:
            window.buttonexcel.setEnabled(True)




    def runeight(self):
        self.started.append(True)
        programBackside.fiyatliste.dataListRobolink()
        self.finished.append(True)
        if len(self.finished) == 6:
            window.buttonexcel.setEnabled(True)


    def runnine(self):
        self.started.append(True)
        programBackside.fiyatliste.dataListMtRobit()
        self.finished.append(True)
        if len(self.finished) == 6:
            window.buttonexcel.setEnabled(True)


    def runten(self):
        self.started.append(True)
        programBackside.fiyatliste.dataListRbtstn()
        self.finished.append(True)
        if len(self.finished) == 6:
            window.buttonexcel.setEnabled(True)


    def runeleven(self):
        self.started.append(True)
        programBackside.fiyatliste.dataListRobitshop()
        self.finished.append(True)
        if len(self.finished) == 6:
            window.buttonexcel.setEnabled(True)


    def listener(self):

        print("listener aktif")

        while True:

            rbcsayac = programBackside.excelokuyucu.sayacrbc
            rbtshpsayac = programBackside.excelokuyucu.sayacrbtshop
            rbtstnsayac = programBackside.excelokuyucu.sayacrbtstn
            direncsayac = programBackside.excelokuyucu.sayacdirenc
            motorobitsayac = programBackside.excelokuyucu.sayacmtrobit
            rblinksayac = programBackside.excelokuyucu.sayacrobolink
            totalsayac = rbcsayac+rbtshpsayac+rblinksayac+rbtstnsayac+direncsayac+motorobitsayac
            window.kalanlabel.setText(str(window.totalex)+"/"+str(totalsayac))


            self.sayac += 1
            time.sleep(2)


worker = Workerone()





class MainWindow(QMainWindow):


    def __init__(self):

        super().__init__()




        self.filepath = []

        layaout = QVBoxLayout()



        self.labelLink=QLabel(self)
        self.labelLink.move(30, 120)

        self.kalanlabel =QLabel(self)
        self.kalanlabel.move(350, 60)

        self.kalan = QLabel(self)
        self.kalan.move(285, 60)
        self.kalan.setText("İLERLEME | ")

        layaout.addWidget(self.labelLink)
        layaout.addWidget(self.kalan)
        layaout.addWidget(self.kalanlabel)


        self.setWindowIcon(QIcon("botimage.ico"))






        self.setAcceptDrops(True)


        self.setFixedSize(480,160)
        self.setWindowTitle("DataScraper V.0.3")

        self.buttoncalis = QPushButton("Botları Çalıştır",self)
        self.buttoncalis.setDisabled(True)
        self.buttonexcel = QPushButton("Excel Çıkart",self)
        self.buttoncalis.move(40, 60)
        self.buttonexcel.move(150, 60)
        self.buttonexcel.setDisabled(True)
        self.buttoncalis.clicked.connect(self.main)
        self.buttonexcel.clicked.connect(self.final)
        self.setLayout(layaout)
        self.messagefinished = QMessageBox(self)
        self.messagefinished.setWindowTitle("BİLGİ")
        self.messagefinished.setText("EXCEL Dosyası Belirttiğiniz Yere kopyalandı")



    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):

        self.kalanlabel.setText("0")

        lines = []
        for url in event.mimeData().urls():
            programBackside.excelokuyucu.excelfile.clear()
            lines.append(url.toLocalFile())

        FilePath=os.path.basename(lines[0])
        self.labelLink.setText(FilePath)
        programBackside.excelokuyucu.excelrbcliste.clear()
        programBackside.excelokuyucu.excelrobolinkliste.clear()
        programBackside.excelokuyucu.exceldirenclinkliste.clear()
        programBackside.excelokuyucu.excelrbtstnliste.clear()
        programBackside.excelokuyucu.excelmotorobitliste.clear()
        programBackside.excelokuyucu.excelrobitshopliste.clear()


        try:
            self.excelFile=pd.read_excel("{}".format(lines[0]), sheet_name="girdi", header=None,
                      skiprows=1, na_filter=False)
            programBackside.excelokuyucu.excelfile.append(self.excelFile)

        except Exception as e:
            print(e)



        try:

            programBackside.excelokuyucu.exceloku_ad()
            programBackside.excelokuyucu.exceloku_kod()
            programBackside.excelokuyucu.excelokurbc()
            programBackside.excelokuyucu.excelokurbtstn()
            programBackside.excelokuyucu.excelokudirenc()
            programBackside.excelokuyucu.excelokurblink()
            programBackside.excelokuyucu.excelokumtrobit()
            programBackside.excelokuyucu.excelokurbtshop()

            rbcex = len(programBackside.excelokuyucu.excelrbcliste)
            rblinkex = len(programBackside.excelokuyucu.excelrobolinkliste)
            direncex = len(programBackside.excelokuyucu.exceldirenclinkliste)
            rbtstnex = len(programBackside.excelokuyucu.excelrbtstnliste)
            mtrobitex = len(programBackside.excelokuyucu.excelmotorobitliste)
            rbtshpex = len(programBackside.excelokuyucu.excelrobitshopliste)
            self.totalex = rbcex + rblinkex + direncex + rbtstnex + mtrobitex + rbtshpex



        except Exception as e:
            print(e)

        window.buttoncalis.setEnabled(True)

    def main(self):
        self.buttoncalis.setDisabled(True)

        try:


                a=threading.Thread(target=worker.listener,daemon=True)
                b=threading.Thread(target=worker.run,daemon=True)
                c=threading.Thread(target=worker.runseven,daemon=True)
                d=threading.Thread(target=worker.runeight,daemon=True)
                e=threading.Thread(target=worker.runnine,daemon=True)
                f=threading.Thread(target=worker.runten,daemon=True)
                g = threading.Thread(target=worker.runeleven,daemon=True)
                a.start()
                b.start()
                c.start()
                d.start()
                e.start()
                f.start()
                g.start()








        except:
            print("hata")


    def final(self):

        file = str(QFileDialog.getExistingDirectory(self, "ÇIKIŞ SEÇ"))
        if len(file) > 3:

            excelolusturucu = programBackside.ExcelCreator(file)
            h=threading.Thread(target=excelolusturucu.FinalExcelCreator,daemon=True)
            h.start()
            worker.finished.clear()
            self.messagefinished.show()


        else:
            self.messagefinished.setText("Hedef Belirtmelisin!")
            self.messagefinished.show()







app = QApplication(sys.argv)

window = MainWindow()
window.show()

app.exec()