import sys
import os
from os.path import expanduser
from PyQt4 import QtCore, QtGui, uic
from process import save_excel

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
        self.progressBar.setMaximum(100)
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
        save_excel(dirIAGA_,dirIMFV_)
        while True:
            value = self.progressBar.value() + (100/numfile)
            self.progressBar.setValue(value)
            if (value >= self.progressBar.maximum()):
                self.progressBar.setValue(100)
                self.labelfolder_3.setText('Done')
                break
            
if __name__ == "__main__":
    app = QtGui.QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())
    
