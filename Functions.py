import openpyxl
import re
import PySimpleGUI as pg
import matplotlib


matplotlib.use('Agg')
Verficador={}
Verficando={}
diff1={}
diff={}

def Begin():
    pg.theme("DarkBlue")
    layout=[
              [pg.Text(f'Insira os caminhos dos arquivos:',font = ("Bold", 11))],
              [pg.Text('Lista:', size =(15, 1)), pg.InputText(key='-Lista-')],
              [pg.Text('PQ:', size =(15, 1)), pg.InputText(key='-Pq-')],
              [pg.Button("Submit"), pg.Button("Cancel")]
          ]
    window= pg.Window('Excel-Verificador',layout)
    event,values = window.read()
    window.close()
    if (check_Files(values['-Pq-']) and check_Files(values['-Lista-'])):
      open_Le(values['-Lista-'])
      open_PQ(values['-Pq-'])
      check_PQ_LE()
      check_LE_PQ()
      open_popup()

def check_Files(file):
    if(file.find('.xlsx') or file.find('.xlsm')):
       return True
       
    else:
       raise ValueError("Formato .xls não suportado")
   
def check_pattern(letter):
  pattern= '(\D{2,3})-(\w{5,10})-(\d{2,4})'
  match = re.search(pattern, letter)
  if (match != None):
     return True
  else:
     return False      

    

def open_Le(element):
  theFile = openpyxl.load_workbook(element)
  allSheetNames = theFile.sheetnames
  for sheet in allSheetNames:
     currentSheet = theFile[sheet]
     for row in range(14, currentSheet.max_row + 1):
          cell_name = "{}{}".format("A", row)
          if (currentSheet[cell_name].value != None and  currentSheet[cell_name].value != "-" and check_pattern(currentSheet[cell_name].value) == True):
            Verficando[formato(currentSheet[cell_name].value)]= f'Equipamentos:{cell_name}'

  return Verficando
  
  
def open_PQ(element):
  theFile = openpyxl.load_workbook(element)
  allSheetNames = theFile.sheetnames
  for sheet in allSheetNames:
     currentSheet = theFile[sheet]
     for row in range(12, currentSheet.max_row + 1):
        cell_name = "{}{}".format("D", row)
        if (currentSheet[cell_name].value != None and  currentSheet[cell_name].value != "-"):
          aux=currentSheet[cell_name].value[:3] + str(sheet) + "-"+ currentSheet[cell_name].value[3:]
          if (check_pattern(aux)): 
            Verficador[formato(aux)]= f'Sheet:{sheet}:{cell_name}'
          
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
    
def check_PQ_LE():
  for element, val in Verficador.items():
    if ( not (element in Verficando)):
      diff[element]=val
 
def check_LE_PQ():
    for element, val in Verficando.items():
      if ( (not (element in Verficador.keys()))):
        diff1[element]=val

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
      s = '***Itens não presentes na Lista***\n'
      s += '\n'.join([str(i) for i in diff.items()])
      s1 = '***Itens não presentes na PQ***\n'
      s1 += '\n'.join([str(i) for i in diff1.items()])
      text= s+'\n'+s1
      column = [[pg.Text(text, font=('Courier New', 12))]]
      layout=[
          [pg.Text(f'Inconsistências Encontradas em:', font=('Bold', 16))],
          [pg.Column(column, size=(800, 300), scrollable=True, key = "Column")],
          #[pg.Text(text)],
          [pg.Button("OK"), pg.Button("Cancel")]
      ]
      window= pg.Window("Tabela de Inconsistências",layout)
      while True:
          print(window.read())
          break 
       
   
    

