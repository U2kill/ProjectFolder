#from UI#
# from PyQt5 import uic
# from PyQt5.QtWidgets import QApplication
# Form, Window = uic.loadUiType("untitled.ui")
# app = QApplication([])
# window = Window()
# form = Form()
# form.setupUi(window)
# window.show()
# app.exec_()

from pathlib import Path
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow
from ui import Ui_MainWindow  # импорт сгенерированного интерфейса
from PyQt5.QtWidgets import QFileDialog
import txtToSwct


operation = "SWCR creator"
filePath = "value"
savePath = "value"
textPath = []

class MyApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.ui.comboBox.currentTextChanged.connect(self.comboBoxFunc)
        self.ui.pushButton.clicked.connect(self.openFileDialog)
        self.ui.pushButton_2.clicked.connect(self.openFileDialog2)


    def comboBoxFunc(self, text):
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
                

    def openFileDialog2(self):
        global savePath
        curFilePath = QFileDialog.getExistingDirectory(parent=None, caption="Выберите папку", directory="/home/user")
        if len(curFilePath) > 2:
            savePath = curFilePath
            self.ui.pushButton_2.setText(savePath)





if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())
