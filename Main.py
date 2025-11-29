
from pathlib import Path
import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QFileDialog
from ui import Ui_MainWindow  # импорт сгенерированного интерфейса
from SWCT import Text
from PySide6.QtCore import QRunnable, QThreadPool, QTimer, Slot
from Yamazumi import Yamazumi
from JES import Jes
from typing import Union, List
import inspect

filePath = None
savePath = None
# textPath = []
log = []
SWCTname = "Value"



class MyApp(QMainWindow):

    def __init__(self):
        super().__init__()
        self.textPath = []
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.operation = "SWCT creator"
        self.appStat = AppStat(self.ui)
        self.changeUI = ChangeUi(self.ui)
        self.log = Logger(self.ui)
        self.ui.comboBox.currentTextChanged.connect(self.changeUI.comboBoxFunc)
        self.ui.pushButton.clicked.connect(self.openFileDialog)
        self.ui.pushButton_2.clicked.connect(self.openSaveFileDialog)
        self.ui.lineEdit.editingFinished.connect(self.saveName)
        self.ui.pushButton_3.clicked.connect(self.startOp)
        self.threadpool = QThreadPool()

    def openFileDialog(self):
        if self.ui.comboBox.currentText() == "SWCT Creator":
            # global textPath
            curFilePath, _ = QFileDialog.getOpenFileNames(self,"Выберите файл","",";Текстовые (*.txt);")
            if len(curFilePath) > 0:
                self.textPath = self.appStat.status(curFilePath)
                # textPath = curFilePath[:]
                self.ui.pushButton.setText(f"{Path(curFilePath[0]).parent}")
                self.log.addLog(f"Добавлены файлы: \n{"\n".join(curFilePath)}")
                # print(textPath)
        
        else:
            global filePath
            curFilePath, _ = QFileDialog.getOpenFileName(self, "Выберите файл", "", "Текстовые (*.xlsm)")
            if len(curFilePath) > 0:
                dir = Path(curFilePath)
                filePath = curFilePath
                self.ui.pushButton.setText(dir.stem)
                log.append(f"Добавлен файл SWCT: \n{curFilePath}")
                self.logger(log)


    def openSaveFileDialog(self):
        global savePath
        curFilePath = QFileDialog.getExistingDirectory(None, "Выберите папку", "/home/user")
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
        if self.appStat.nameChecker() == True:
            self.ui.lineEdit.setStyleSheet("")
            global SWCTname
            SWCTname = f"{self.ui.lineEdit.text().strip()}.xlsm"
            print(SWCTname)
            log.append(f"Добавлено имя файла: \n{SWCTname}")
            self.logger(log)
        
    def startOp(self):
        if self.ui.comboBox.currentText() == "SWCT Creator":
            if self.appStat.nameChecker() == True and self.appStat.filePathChecker(self.textPath, self.operation) == True:
                print(self.textPath)
                result = Text(self.textPath, SWCTname, savePath)
                log.append("Создание документа...")
                self.logger(log)
                self.threadpool.start(result)
                result.signals.progress.connect(lambda txt:(log.append(f"{txt}"), self.logger(log)))
                result.signals.finished.connect(lambda: (log.append(f"\nСоздан файл: {SWCTname}"), self.logger(log)))

            elif self.appStat.nameChecker() == False:
                self.appStat.namError()

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


class AppStat:
    def __init__(self, ui):
        self.ui = ui

    def status(self, value):
        callerFrame = inspect.stack()[1]
        callerName = callerFrame.function

        if callerName == "openFileDialog":
            try:
                if Path(value[0]).suffix == ".txt":
                    textPath = value
                    return textPath
            
            except:
                pass
            

    def nameChecker(self): # Добавить нормальную проверку имени файла
        if len(str(self.ui.lineEdit.text())) == 0:
            return False
        else:
            return True
    
    def filePathChecker(self, textPath, operation):
        if operation == "SWCT creator":
            if textPath == None or savePath == None:
                return False
            else:
                return True

        if operation == "Yamazumi Creator" or operation == "JES Creator":
            if filePath == None or savePath == None:
                return False
            else:
                return True    

    def namError(self):
        QApplication.beep()
        self.ui.lineEdit.setStyleSheet("background-color: pink;")


class ChangeUi:
    def __init__(self, ui):
        self.ui = ui

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
        return self.operation


class Logger:
    def __init__(self, ui):
        self.ui = ui    
        self.logger = []

    def addLog(self, log: Union[str, List[str]]):
        if isinstance(log, str):
            self.logger.append(log)
        
        elif isinstance(log, list):
            self.logger.extend(log)

        self.uiLog = "\n".join(self.logger)
        self.ui.textEdit.setText(self.uiLog)






# class FileManager:
    


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()
    app.exec_()

myString = True
