from PyQt5.QtWidgets import QApplication, QPushButton, QWidget, QVBoxLayout, QLabel
from PyQt5.QtCore import QThread, pyqtSignal
import time

class Worker(QThread):
    finished = pyqtSignal(str)

    def run(self):
        # Долгая операция
        time.sleep(5)
        self.finished.emit("Готово!")

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.label = QLabel("Жду...")
        self.button = QPushButton("Запустить")
        self.button.clicked.connect(self.start_task)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.button)
        self.setLayout(layout)

    def start_task(self):
        self.thread = Worker()
        self.thread.finished.connect(self.on_finished)
        self.thread.start()  # запускаем в отдельном потоке

    def on_finished(self, text):
        self.label.setText(text)

app = QApplication([])
window = MainWindow()
window.show()
app.exec()