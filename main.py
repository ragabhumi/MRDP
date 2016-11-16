import sys
import os
import glob
import numpy as np
from os.path import expanduser
from PyQt4 import QtCore, QtGui, uic
from process import min2hour
from process import format_excel

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
        self.SearchButtonIMFV.clicked.connect(self.SearchFolderIMFV)
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
    def SearchFolderIMFV(self):
        global dirIMFV_
        dirIMFV_ = QtGui.QFileDialog.getExistingDirectory(None, 'Select a folder:', home+'\\data', QtGui.QFileDialog.ShowDirsOnly)
        self.labelfolderIMFV.setText(dirIMFV_)
        return dirIMFV_
    def normalOutputWritten(self, text):
        cursor = self.textEdit.textCursor()
        cursor.movePosition(QtGui.QTextCursor.End)
        cursor.insertText(text)
        self.textEdit.setTextCursor(cursor)
        self.textEdit.ensureCursorVisible()
    def Execute(self):
        listfile = glob.glob(os.path.join(str(dirIAGA_), '*.min'))
        x = [[np.nan for u in range(24)] for u in range(len(listfile))]
        y = [[np.nan for u in range(24)] for u in range(len(listfile))]
        z = [[np.nan for u in range(24)] for u in range(len(listfile))]
        f = [[np.nan for u in range(24)] for u in range(len(listfile))]
        h = [[np.nan for u in range(24)] for u in range(len(listfile))]
        d = [[np.nan for u in range(24)] for u in range(len(listfile))]
        i = [[np.nan for u in range(24)] for u in range(len(listfile))]
        value = self.progressBar.value()
        self.progressBar.setMaximum(len(listfile))
        self.labelfolder_3.setText('Creating excel file, please wait...')
        for j in range(0,len(listfile)):
            data = min2hour(listfile[j])
            x[j] = data[0]
            y[j] = data[1]
            z[j] = data[2]
            f[j] = data[3]
            h[j] = data[4]
            d[j] = data[5]
            i[j] = data[6]
            print os.path.basename(glob.glob(os.path.join(str(dirIAGA_), '*.min'))[j])
            value = j+1
            self.progressBar.setValue(value)        
        format_excel(x,y,z,f,h,d,i,dirIAGA_,dirIMFV_)
        self.labelfolder_3.setText('Done')
            
if __name__ == "__main__":
    app = QtGui.QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())
    
