# import os
# import PyQt5.QtCore

# os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = (
#     os.path.join(
#         os.path.dirname(PyQt5.QtCore.__file__),
#         "Qt5",
#         "plugins",
#         "platforms"
#     )
# )



from PyQt5 import uic
from PyQt5.QtWidgets import QApplication
Form, Window = uic.loadUiType("untitled.ui")
app = QApplication([])
window = Window()
form = Form()
form.setupUi(window)
window.show()
app.exec_()