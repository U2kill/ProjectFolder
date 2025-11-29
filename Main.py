
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







class MyApp(QMainWindow):

    def __init__(self):
        super().__init__()

        self.filePath = None
        self.textPath = None
        self.operation = "SWCT creator"

        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.log = Logger(self.ui)
        self.appStat = AppStat(self.ui, self.log)
        self.longTask = LongTask(self.appStat, self.log)
        self.changeUI = ChangeUi(self.ui)

        self.ui.comboBox.currentTextChanged.connect(self.changeUI.comboBoxFunc)
        self.ui.pushButton.clicked.connect(self.openFileDialog)
        self.ui.pushButton_2.clicked.connect(self.openSaveFileDialog)
        self.ui.lineEdit.editingFinished.connect(self.saveName)
        self.ui.pushButton_3.clicked.connect(self.startOp)

        

    def openFileDialog(self):
        if self.ui.comboBox.currentText() == "SWCT Creator":
            curFilePath, _ = QFileDialog.getOpenFileNames(self,"Выберите файл","",";Текстовые (*.txt);")
            if len(curFilePath) > 0:
                self.textPath = self.appStat.status(curFilePath)
                self.ui.pushButton.setText(f"{Path(curFilePath[0]).parent}")
                self.log.addLog(f"Добавлены файлы: \n{"\n".join(curFilePath)}")
        
        else:
            curFilePath, _ = QFileDialog.getOpenFileName(self, "Выберите файл", "", "Текстовые (*.xlsm)")
            if len(curFilePath) > 0:
                self.filePath = self.appStat.status(curFilePath)
                self.ui.pushButton.setText(Path(curFilePath).stem)
                self.log.addLog(f"Добавлен файл SWCT: \n{curFilePath}")



    def openSaveFileDialog(self):
        curFilePath = QFileDialog.getExistingDirectory(None, "Выберите папку", "/home/user")
        if len(curFilePath) > 2:
            self.savePath = self.appStat.status(curFilePath)
            self.ui.pushButton_2.setText(self.savePath)
            self.log.addLog(f"Добавлена директория для сохранения документа: \n{self.savePath}")

    def saveName(self):
        if self.appStat.nameChecker() == True:
            self.ui.lineEdit.setStyleSheet("")
            self.SWCTname = f"{self.ui.lineEdit.text().strip()}.xlsm"
            self.log.addLog(f"Добавлено имя файла: \n{self.SWCTname}")
        
    def startOp(self):
        if self.ui.comboBox.currentText() == "SWCT Creator":
            self.longTask.createSWCT(self.textPath, self.savePath, self.SWCTname)
            # if self.appStat.nameChecker() == True and self.appStat.filePathChecker(self.textPath, self.savePath) == True:
            #     result = Text(self.textPath, self.SWCTname, self.savePath)
            #     self.log.addLog("Создание документа...")
            #     self.threadpool.start(result)
            #     result.signals.progress.connect(lambda txt:(self.log.addLog(f"{txt}")))
            #     result.signals.finished.connect(lambda: (self.log.addLog(f"\nСоздан файл: {self.SWCTname}")))

            # elif self.appStat.nameChecker() == False:
            #     self.appStat.namError()

        elif self.ui.comboBox.currentText() == "Yamazumi Creator":
            if self.appStat.filePathChecker(self.filePath, self.savePath) == True:
                self.log.addLog("Создание документа...")
                result = Yamazumi(self.filePath, self.savePath)
                self.threadpool.start(result)
                result.signals.finished.connect(lambda: (self.log.addLog("\nСоздан файл: Yamazumi цеха.xlsx")))
    
        elif self.ui.comboBox.currentText() == "JES Creator":
            if self.appStat.filePathChecker(self.filePath, self.savePath) == True:
                self.log.addLog("Создание документа...")
                result = Jes(self.filePath, self.savePath)
                self.threadpool.start(result)
                result.signals.finished.connect(lambda: (self.log.addLog("\nСоздан файл: Jes.xlsx")))


class AppStat:
    def __init__(self, ui, log):
        self.ui = ui
        self.log = log

    def status(self, value):
        callerFrame = inspect.stack()[1]
        callerName = callerFrame.function

        if callerName == "openFileDialog":
            return self.FileDialog(value)
        
        if callerName == "openSaveFileDialog":
            return value

    def FileDialog(self, value):
            try:
                if Path(value[0]).suffix == ".txt":
                    textPath = value
                    return textPath
            except:
                pass

            try: 
                if Path(value).suffix == ".xlsm":
                    filePath = value
                    return filePath
            except:
                pass        


    def nameChecker(self): # Добавить нормальную проверку имени светильника
        if len(str(self.ui.lineEdit.text())) == 0:
            return False
        else:
            return True

        
    def filePathChecker(self, *paths: Union[str, List[str]]):
        try:
            for path in paths:
                if len(path) == 0:
                    return False
            return True
        except TypeError:

            QApplication.beep()
            self.log.addLog("Ошибка: Не выбраны все пути файлов")

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

class LongTask:
    def __init__(self, stat, log):
        self.appStat = stat
        self.log = log
        self.threadpool = QThreadPool()

    def createSWCT(self, textPath: List[str], savePath: str, SWCTname: str):
        if self.appStat.nameChecker() == True and self.appStat.filePathChecker(textPath, savePath) == True:
            result = Text(textPath, SWCTname, savePath)
            self.log.addLog("Создание документа...")
            self.threadpool.start(result)
            result.signals.progress.connect(lambda txt:(self.log.addLog(f"{txt}")))
            result.signals.finished.connect(lambda: (self.log.addLog(f"\nСоздан файл: {SWCTname}")))

        elif self.appStat.nameChecker() == False:
            self.appStat.namError()
    


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()
    app.exec()

myString = True
