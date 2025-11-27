
from pathlib import Path
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow
from ui import Ui_MainWindow  # импорт сгенерированного интерфейса
from PyQt5.QtWidgets import QFileDialog
import Waste.txtToSwct as txtToSwct
from SWCT import Text
from PyQt5.QtCore import QRunnable, QThreadPool, QTimer, pyqtSlot
from Yamazumi import Yamazumi
from JES import Jes


filePath = None
savePath = None
textPath = []
log = []
SWCTname = "Value"


class MyApp(QMainWindow):

    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.operation = "SWCT creator"
        # self.operation = ""
        self.ui.comboBox.currentTextChanged.connect(self.comboBoxFunc)
        self.ui.pushButton.clicked.connect(self.openFileDialog)
        self.ui.pushButton_2.clicked.connect(self.openFileDialog2)
        self.ui.lineEdit.editingFinished.connect(self.saveName)
        self.ui.pushButton_3.clicked.connect(self.startOp)
        
        self.threadpool = QThreadPool()

    def nameChecker(self): # Добавить нормальную проверку имени файла
        if len(str(self.ui.lineEdit.text())) == 0:
            return False
        else:
            return True

    def filePathChecker(self):
        if self.operation == "SWCT creator":
            if textPath == None or savePath == None:
                return False
            else:
                return True

        if self.operation == "Yamazumi Creator" or self.operation == "JES Creator":
            if filePath == None or savePath == None:
                return False
            else:
                return True


    def namError(self):
        QApplication.beep()
        self.ui.lineEdit.setStyleSheet("background-color: pink;")

    def comboBoxFunc(self, text):
        if self.ui.comboBox.currentText() == "SWCT Creator":
            self.ui.pushButton.setText("Select text")
            self.ui.label_2.setText("Text files path:")
            self.ui.lineEdit.show()
         
        else:
            self.ui.pushButton.setText("Select SWCT")
            self.ui.label_2.setText("SWCT file path:")
            self.ui.lineEdit.hide()
        self.operation = text

    def openFileDialog(self):
        if self.ui.comboBox.currentText() == "SWCT Creator":
            global textPath
            curFilePath, _ = QFileDialog.getOpenFileNames(self,"Выберите файл","",";Текстовые (*.txt);")
            if len(curFilePath) > 0:
                dir = Path(curFilePath[0])
                textPath = curFilePath[:]
                self.ui.pushButton.setText(f"{dir.parent}")
                L = "\n".join(textPath)
                log.append(f"Добавлены файлы: \n{L}")
                self.logger(log)
                print(textPath)
        
        else:
            global filePath
            curFilePath, _ = QFileDialog.getOpenFileName(self,"Выберите файл","",";Текстовые (*.xlsm);")
            if len(curFilePath) > 0:
                dir = Path(curFilePath)
                filePath = curFilePath
                self.ui.pushButton.setText(dir.stem)
                log.append(f"Добавлен файл SWCT: \n{curFilePath}")
                self.logger(log)
                print(filePath)


    def openFileDialog2(self):
        global savePath
        curFilePath = QFileDialog.getExistingDirectory(parent=None, caption="Выберите папку", directory="/home/user")
        if len(curFilePath) > 2:
            savePath = curFilePath
            # print((savePath))
            self.ui.pushButton_2.setText(savePath)
            log.append(f"Добавлена директория для сохранения документа: \n{savePath}")
            self.logger(log)

    def logger(self, log):
        for i in log:
            str(i)
        log = "\n".join(log)
        self.ui.textEdit.setText(log)

    def saveName(self):
        if self.nameChecker() == True:
            self.ui.lineEdit.setStyleSheet("")
            global SWCTname
            SWCTname = f"{self.ui.lineEdit.text().strip()}.xlsm"
            print(SWCTname)
            log.append(f"Добавлено имя файла: \n{SWCTname}")
            self.logger(log)
        
    def startOp(self):
        if self.ui.comboBox.currentText() == "SWCT Creator":
            if self.nameChecker() == True and self.filePathChecker() == True:
                result = Text(textPath, SWCTname, savePath)
                log.append("Создание документа...")
                self.logger(log)
                self.threadpool.start(result)
                result.signals.progress.connect(lambda txt:(log.append(f"{txt}"), self.logger(log)))
                result.signals.finished.connect(lambda: (log.append(f"\nСоздан файл: {SWCTname}"), self.logger(log)))

            elif self.nameChecker() == False:
                self.namError()

        elif self.ui.comboBox.currentText() == "Yamazumi Creator":
            if self.filePathChecker() == True:
                log.append("Создание документа...")
                self.logger(log)
                result = Yamazumi(filePath, savePath)
                self.threadpool.start(result)
                result.signals.finished.connect(lambda: (log.append("\nСоздан файл: Yamazumi цеха.xlsx"), self.logger(log)))
    
        elif self.ui.comboBox.currentText() == "JES Creator":
            if self.filePathChecker() == True:
                log.append("Создание документа...")
                self.logger(log)
                result = Jes(filePath, savePath)
                self.threadpool.start(result)
                result.signals.finished.connect(lambda: (log.append("\nСоздан файл: Jes.xlsx"), self.logger(log)))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()
    app.exec_()
