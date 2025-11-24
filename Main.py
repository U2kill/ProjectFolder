
from pathlib import Path
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow
from ui import Ui_MainWindow  # импорт сгенерированного интерфейса
from PyQt5.QtWidgets import QFileDialog
import txtToSwct
from SWCT import Text

from PyQt5.QtCore import QRunnable, QThreadPool, QTimer, pyqtSlot



operation = "SWCR creator"
filePath = "value"
savePath = "value"
textPath = []
log = []
SWCTname = "Value"

class MyApp(QMainWindow):

    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        
        self.ui.comboBox.currentTextChanged.connect(self.comboBoxFunc)
        self.ui.pushButton.clicked.connect(self.openFileDialog)
        self.ui.pushButton_2.clicked.connect(self.openFileDialog2)
        self.ui.lineEdit.editingFinished.connect(self.saveName)
        # self.ui.pushButton_4.clicked.connect(self.saveName)
        self.ui.pushButton_3.clicked.connect(self.startSWCT)
        
        self.threadpool = QThreadPool()

    def comboBoxFunc(self, text):
        if self.ui.comboBox.currentText() == "SWCT Creator":
            self.ui.lineEdit.show()
         
        else:
            self.ui.lineEdit.hide()


        global operation
        operation = text


    def openFileDialog(self):
        if self.ui.comboBox.currentText() == "SWCT Creator":
            global textPath
            curFilePath, _ = QFileDialog.getOpenFileNames(self,"Выберите файл","",";Текстовые (*.txt);")
            if len(curFilePath) > 0:
                dir = Path(curFilePath[0])
                textPath = curFilePath[:]
                txtToSwct.PathList = textPath
                self.ui.pushButton.setText(f"{dir.parent}")
                L = "\n".join(textPath)
                log.append("Добавлены файлы:")
                log.append(f"{L}")
                self.logger(log)
                print(textPath)
                

    def openFileDialog2(self):
        global savePath
        curFilePath = QFileDialog.getExistingDirectory(parent=None, caption="Выберите папку", directory="/home/user")
        if len(curFilePath) > 2:
            savePath = curFilePath
            self.ui.pushButton_2.setText(savePath)
            log.append(f"Добавлена директория для сохранения документа: \n{savePath}")
            self.logger(log)

        
    def logger(self, log):
        for i in log:
            str(i)
        log = "\n".join(log)
        self.ui.textEdit.setText(log)

    def saveName(self):
        global SWCTname
        SWCTname = f"{self.ui.lineEdit.text().strip()}.xlsm"
        print(SWCTname)
        log.append(f"Добавлено имя файла: \n{SWCTname}")
        self.logger(log)


    def startSWCT(self):
        

        result = Text(textPath, SWCTname, savePath)
        self.threadpool.start(result)
        # result.signals.finished.connect(print("AYE"))
 

    
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()
    app.exec_()
