import sys
import os
import glob
from calendar import monthrange
from os.path import expanduser
from PyQt4 import QtCore, QtGui, uic
from process import iaga2imfv,min2hour,format_excel

home = expanduser("~")
qtCreatorFile = "gui.ui" # Enter file here.

Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)

class EmittingStream(QtCore.QObject):
    textWritten = QtCore.pyqtSignal(str)
    def write(self, text):
        self.textWritten.emit(str(text))
            
class MyApp(QtGui.QMainWindow, Ui_MainWindow):
    def __init__(self):
        QtGui.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        self.SearchButtonIAGA.clicked.connect(self.SearchFolderIAGA)
        self.ExecuteButton.clicked.connect(self.Execute)
        self.progressBar.setMinimum(0)
        sys.stdout = EmittingStream(textWritten=self.normalOutputWritten)
    def __del__(self):
        # Restore sys.stdout
        sys.stdout = sys.__stdout__    
    def SearchFolderIAGA(self):
        global dirIAGA_,numfile
        dirIAGA_ = QtGui.QFileDialog.getExistingDirectory(None, 'Select a folder:', home+'\\data', QtGui.QFileDialog.ShowDirsOnly)
        self.labelfolderIAGA.setText(dirIAGA_)
        numfile = len([name for name in os.listdir(str(dirIAGA_)) if name.endswith('.min') and os.path.isfile(os.path.join(str(dirIAGA_), name))])
        return dirIAGA_,numfile
    def normalOutputWritten(self, text):
        cursor = self.textEdit.textCursor()
        cursor.movePosition(QtGui.QTextCursor.End)
        cursor.insertText(text)
        self.textEdit.setTextCursor(cursor)
        self.textEdit.ensureCursorVisible()
    def Execute(self):
        listfile = glob.glob(os.path.join(str(dirIAGA_), '*.min'))
        first_data = os.path.basename(listfile[0])
        tahun1 = int(first_data[3:7])
        bulan1 = int(first_data[7:9])
        hari_bulan = monthrange(tahun1, bulan1)[1]
        x = [['AM' for u in range(24)] for u in range(hari_bulan)]
        y = [['AM' for u in range(24)] for u in range(hari_bulan)]
        z = [['AM' for u in range(24)] for u in range(hari_bulan)]
        f = [['AM' for u in range(24)] for u in range(hari_bulan)]
        h = [['AM' for u in range(24)] for u in range(hari_bulan)]
        d = [['AM' for u in range(24)] for u in range(hari_bulan)]
        i = [['AM' for u in range(24)] for u in range(hari_bulan)]
        value = self.progressBar.value()
        self.progressBar.setMaximum(len(listfile))
        self.labelfolder_3.setText('Creating excel file, please wait...')
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
            print os.path.basename(glob.glob(os.path.join(str(dirIAGA_), '*.min'))[j])
            value = j+1
            self.progressBar.setValue(value)        
        format_excel(x,y,z,f,h,d,i,dirIAGA_)
        self.labelfolder_3.setText('Done')
            
if __name__ == "__main__":
    app = QtGui.QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())
    
