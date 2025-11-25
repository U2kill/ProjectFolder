##########################
###SWCT TO YAMAZUMI CEX###
##########################

from openpyxl import load_workbook
from pathlib import Path





xlPath = Path("SWCT Светильник LINE ILF30-1,5W40-30H-150(50P02).xlsm")
templateSite = Path("/content/drive/MyDrive/Yamazumi/Yamazumi_Участка.xlsx")
templateWorkshop = Path("/content/drive/MyDrive/Yamazumi/Yamazumi_Цеха.xlsx")


activeWb = load_workbook(xlPath, keep_vba=True)
templateWorkshop = load_workbook(templateWorkshop)
templateSite = load_workbook(templateSite)

class Yamazumi:
###ПАРСЕР###
  def createOperationsList(counter):
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
    siteList = ["РЕГУЛИРОВКА","МЕХ СБ","ППМ","ЛИНЗЫ","ПРСБ","ПРОГОН","СБОРКА","УПАКОВКА"]
    sectionList=["I16:K35", "N16:P35", "S16:U35", "X16:Z35", "AC16:AE35", "AH16:AJ35", "AM16:AO35", "AR16:AT35"]
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

  sheet = activeWb["SWCT"]

  counter = 9
  for row in sheet["J9:J300"]:
    for cell in row:
      if cell.value != None:
        counter += 1

  countSec = f"M9:M{counter}"
  operationsList = createOperationsList(counter)

  sheet = templateWorkshop["YAMAZUMI цеха"]

  writeInWorkshop(operationsList, sheet)

  templateWorkshop.save("Yamazumi цеха.xlsx")