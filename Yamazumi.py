##########################
###SWCT TO YAMAZUMI CEX###
##########################

from openpyxl import load_workbook
from pathlib import Path

from PyQt5.QtCore import (
    QObject,
    QRunnable,
    QThreadPool,
    QTimer,
    pyqtSignal,
    pyqtSlot,
)



# xlPath = Path("SWCT Светильник LINE ILF30-1,5W40-30H-150(50P02).xlsm")
# templateWorkshop = Path("/content/drive/MyDrive/Yamazumi/Yamazumi_Цеха.xlsx")


# activeWb = load_workbook(xlPath, keep_vba=True)
# templateWorkshop = load_workbook(templateWorkshop)

class Yamazumi():
  def __init__(self, filePath, SWCTname):
    # super().__init__()
    self.SWCTname = SWCTname
    self.filePath = filePath
    
###ПАРСЕР###
  def createOperationsList(counter, sheet):
    num = 9
    sitesList = []
    site = "Value"
    while num <= counter:
      cell_f = sheet[f"F{num}"]
      cell_m = sheet[f"M{num}"]

      if isinstance(cell_m.value, str) and isinstance(cell_f.value, str):
        site = sheet[f"M{num}"].value
        benefit = sheet[f"H{num}"].value
        loss = sheet[f"I{num}"].value
        # удаление всех технологических процессов
        if loss > 7200:
          loss = loss - 7200
        if loss > 25000:
          loss = 0
        sitesList.append({"Операция": sheet[f"F{num}"].value, "Участок": site, "Польза": benefit,"Потери": loss})

      elif cell_m.value is None and isinstance(cell_f.value, str):
          # удаление всех технологических процессов
          loss = sheet[f"I{num}"].value
          if loss > 7200:
              loss = loss - 7200
          if loss > 15000:
              loss = 0
          sitesList.append({"Операция":sheet[f"F{num}"].value, "Участок": site,"Польза": sheet[f"H{num}"].value, "Потери": loss})
      num += 1
    return sitesList


  ###ЗАПИСЬ В YAMAZUMI ЦЕХ###
  def writeInWorkshop(operationsList, sheet):

    siteCol = {"МЕХ СБ": [9,10,11], "ППМ": [14,15,16], "РЕГУЛИРОВКА":[19,20,21], "ПРСБ":[24,25,26], "ПРОГОН":[29,30,31], "ЛИНЗЫ":[34,35,36],"СБОРКА":[39,40,41], "УПАКОВКА":[44,45,46]}
    number = 16
    site = 0
    for i in operationsList:
      if site == 0:
        site = i.get('Участок')

      if site != i.get('Участок'):
        number = 16
        site = i.get('Участок')

      if siteCol.get(i.get('Участок')) != None and sheet.cell(row = number, column = siteCol[i.get('Участок')][1]).value == None:
        sheet.cell(row = number, column = siteCol[i.get('Участок')][0], value = i.get('Операция'))
        sheet.cell(row = number, column = siteCol[i.get('Участок')][1], value = i.get('Польза'))
        sheet.cell(row = number, column = siteCol[i.get('Участок')][2], value = i.get('Потери'))
        number += 1

      else:
        while sheet.cell(row = number, column = siteCol[i.get('Участок')][1]).value != None:
          number += 1

        sheet.cell(row = number, column = siteCol[i.get('Участок')][0], value = i.get('Операция'))
        sheet.cell(row = number, column = siteCol[i.get('Участок')][1], value = i.get('Польза'))
        sheet.cell(row = number, column = siteCol[i.get('Участок')][2], value = i.get('Потери'))
        number += 1
    
  def run(self):

    templateWorkshop = Path("Yamazumi.xlsx")
    templateWorkshop = load_workbook(templateWorkshop)

    activeWb = load_workbook(Path(self.filePath, self.SWCTname))
    sheet = activeWb["SWCT"]

    counter = 9
    for row in sheet["J9:J300"]:
      for cell in row:
        if cell.value != None:
          counter += 1

    operationsList = self.createOperationsList(counter)

    sheet = templateWorkshop["YAMAZUMI цеха"]

    self.writeInWorkshop(operationsList, sheet)

    templateWorkshop.save("Yamazumi цеха.xlsx")


result = Yamazumi(r"C:\Users\pisos\Downloads", "SWCT Светильник LINE ILF30-1,5W40-30H-150(50P02)1.xlsm")
# QThreadPool.globalInstance().start(result)
result.run()