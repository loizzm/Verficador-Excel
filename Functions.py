import openpyxl
import re
import PySimpleGUI as pg
import matplotlib

matplotlib.use('Agg')
Verficador=[]
Verficando=[]
diff=[]


def open_Le():
  theFile = openpyxl.load_workbook('LE-6027SA-K-00001_Rev_B.xlsm')
  currentSheet = theFile["Equipamentos"]
  for row in range(14, currentSheet.max_row + 1):
        cell_name = "{}{}".format("A", row)
        if (currentSheet[cell_name].value != None and  currentSheet[cell_name].value != "-"):
          Verficando.append(formato(currentSheet[cell_name].value))

  return Verficando
  
  
def open_PQ():
  theFile = openpyxl.load_workbook('PQ-6027SA-K-00001_Rev_A.xlsx')
  allSheetNames = theFile.sheetnames
  for sheet in allSheetNames:
     currentSheet = theFile[sheet]
     for row in range(12, currentSheet.max_row + 1):
        cell_name = "{}{}".format("D", row)
        if (currentSheet[cell_name].value != None and  currentSheet[cell_name].value != "-"):
          aux=currentSheet[cell_name].value[:3] + str(sheet) + "-"+ currentSheet[cell_name].value[3:] 
          Verficador.append(formato(aux)) 
          
  return Verficador  


def formato(let):
  new_string=str()
  indice = re.finditer(pattern='-',string=let)
  ind=[index.start() for index in indice]
  if ( let[ind[1]+1] == '0' and let[ind[1]+2] != '0' ):
    new_string = let[:ind[1]+1] + let[ind[1]+2:]
  elif(let[ind[1]+1] == '0' and let[ind[1]+2] == '0'):
     new_string = let[:ind[1]+1] + let[ind[1]+3:] 
  else:
     new_string=let
  return new_string
    
   

def check():
  for element in Verficador:
    if ( not (element in Verficando)):
      diff.append(element)
  return diff
 
def open_popup():
  if(len(diff)==0):
      pg.theme("DarkAmber")  
      layout=[
          [pg.Text("Nenhuma Iconsistência encontrada")],
          [pg.Button("OK"), pg.Button("Cancel")]
      ]
      window= pg.Window("Tabela de Inconsistências",layout)
      while True:
          print(window.read())
          break
  else:
      pg.theme("DarkRed")  
      layout=[
          [pg.Text(f'Inconsistências Encontradas em: \n {diff}')],
          [pg.Button("OK"), pg.Button("Cancel")]
      ]
      window= pg.Window("Tabela de Inconsistências",layout)
      while True:
          print(window.read())
          break 
       
   
    

