# -*- coding: utf-8 -*-
"""
Created on Fri Apr 10 21:08:04 2020

@author: yosi setiawan
"""

import sys,os,glob,time
from calendar import monthrange
from os.path import expanduser
from PyQt5 import QtCore, QtGui, uic, QtWidgets
from process import iaga2imfv,min2hour,format_excel
from step_remove import step_removal

#QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True) #enable highdpi scaling
#QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True) #use highdpi icons

# Set lokasi home folder
home = expanduser("~")
# Nama file GUI
qtCreatorFile = "gui.ui" # Enter file here.

# Load file GUI
Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)


class EmittingStream(QtCore.QObject):
    textWritten = QtCore.pyqtSignal(str)
    def write(self, text):
        self.textWritten.emit(str(text))
        
    def flush(self):
        pass

class MyApp(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        self.SearchButtonIAGA.clicked.connect(self.SearchFolderIAGA)
        self.ExecuteButton.clicked.connect(self.Execute)
        self.StepRemovalButton.clicked.connect(self.stepremoval)
        self.CalendarWidget.clicked[QtCore.QDate].connect(self.showDate)
        self.progressBar.setMinimum(0)
        sys.stdout = EmittingStream(textWritten=self.normalOutputWritten)
        
    def __del__(self):
        # Restore sys.stdout
        sys.stdout = sys.__stdout__
        
    # Fungsi untuk tombol Search Folder IAGA
    def SearchFolderIAGA(self):
        global dirIAGA_,numfile
        dirIAGA_ = QtWidgets.QFileDialog.getExistingDirectory(None, 'Select a folder:', home+'\\data', QtWidgets.QFileDialog.ShowDirsOnly)
        self.labelfolderIAGA.setText(dirIAGA_)
        numfile = len([name for name in os.listdir(str(dirIAGA_)) if name.endswith('.min') and os.path.isfile(os.path.join(str(dirIAGA_), name))])
        return dirIAGA_,numfile
    
    # Fungsi untuk menampilkan teks
    def normalOutputWritten(self, text):
        cursor = self.textEdit.textCursor()
        cursor.movePosition(QtGui.QTextCursor.End)
        cursor.insertText(text)
        self.textEdit.setTextCursor(cursor)
        self.textEdit.ensureCursorVisible()
        
    # Fungsi untuk tombol Execute
    def Execute(self):
        listfile = glob.glob(os.path.join(str(dirIAGA_), '*.min'))
        # Menentukan file data pertama untuk diambil tahun bulan tanggalnya
        first_data = os.path.basename(listfile[0])
        tahun1 = int(first_data[3:7])
        bulan1 = int(first_data[7:9])
        hari_bulan = monthrange(tahun1, bulan1)[1]
        # Baca data
        x = [['AM' for u in range(24)] for u in range(hari_bulan)]
        y = [['AM' for u in range(24)] for u in range(hari_bulan)]
        z = [['AM' for u in range(24)] for u in range(hari_bulan)]
        f = [['AM' for u in range(24)] for u in range(hari_bulan)]
        h = [['AM' for u in range(24)] for u in range(hari_bulan)]
        d = [['AM' for u in range(24)] for u in range(hari_bulan)]
        i = [['AM' for u in range(24)] for u in range(hari_bulan)]
        
        # Set progressbar
        value = self.progressBar.value()
        self.progressBar.setMaximum(len(listfile))
        
        # Print keterangan
        self.labelfolder_3.setText('Creating excel file, please wait...')
        
        # Konversi data dari IAGA ke IMFV dengan fungsi iaga2imfv
        for j in range(0,len(listfile)):
            data = min2hour(listfile[j])
            tanggal = data[7]-1
            x[tanggal] = data[0]
            y[tanggal] = data[1]
            z[tanggal] = data[2]
            f[tanggal] = data[3]
            h[tanggal] = data[4]
            d[tanggal] = data[5]
            i[tanggal] = data[6]
            iaga2imfv(listfile[j],dirIAGA_)
            print(os.path.basename(glob.glob(os.path.join(str(dirIAGA_), '*.min'))[j]))
            value = j+1
            self.progressBar.setValue(value)    
            
        # Konversi ke format excel    
        format_excel(x,y,z,f,h,d,i,dirIAGA_)
        # Selesai
        self.labelfolder_3.setText('Done')
        
    # Fungsi untuk program stepremoval
    def stepremoval(self):
        # Ambil threshold dari form
        threshold = self.lineEdit.displayText()
        # Ambil tanggal dari widget calendar
        date = self.CalendarWidget.selectedDate().toPyDate()
        print('Removing steps at '+date.strftime('%d %B %Y'))
        # Jalankan program step removal dari folder berikut
        #step_removal(date,float(threshold),'Z:\\PROCESS\\DATALEMIRAWFILTER\\',home+'\\data\\Provisional')
        step_removal(date,float(threshold),home+'\\data\\LEMIRAW',home+'\\data\\Provisional')
    
    # Ambil variabel tanggal dari widget calendar
    def showDate(self):     
        date = self.CalendarWidget.selectedDate()
        #self.label.setText(str(date.toPyDate()))

# Jalankan program
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)

    # Create and display the splash screen
    splash_pix = QtGui.QPixmap('splash.jpg')
    splash = QtWidgets.QSplashScreen(splash_pix, QtCore.Qt.WindowStaysOnTopHint)
    splash.setMask(splash_pix.mask())
    splash.show()
    app.processEvents()

    # Pause 2 detik
    time.sleep(2)

    # Program utama
    window = MyApp()
    window.show()
    splash.finish(window)
    sys.exit(app.exec_())