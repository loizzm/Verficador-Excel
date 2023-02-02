import openpyxl
import re
import PySimpleGUI as pg
import matplotlib


matplotlib.use('Agg')
Verficador={}
Verficando={}
diff1={}
diff={}
doppel={}

def Begin(): #Função principal, começa o programa e chama todas as outas
    pg.theme("DarkBlue")
    layout=[
              [pg.Text(f'Insira os caminhos dos arquivos:',font = ("Bold", 11))],    #Cria a tela incial
              [pg.Text('Lista:', size =(15, 1)), pg.InputText(key='-Lista-')],
              [pg.Text('PQ:', size =(15, 1)), pg.InputText(key='-Pq-')],
              [pg.Button("Submit"), pg.Button("Cancel")]
          ]
    window= pg.Window('Excel-Verificador',layout)
    event,values = window.read()
    window.close()
    if (check_Files(values['-Pq-']) and check_Files(values['-Lista-'])): # Checa se os caminhos inseridos estão completos e se são dos arquivos certos
      open_Le(values['-Lista-'])                                         # abre a Le e faz todas as operações necessárias
      open_PQ(values['-Pq-'])                                             #abre a PQ e faz todas as operações necessárias
      check_PQ_LE()                                                      #Checa as incosistências PQ -> LE
      check_LE_PQ()                                                        #Checa as incosistências LE -> PQ
      open_popup()                                                         #Feito as operações de comparação, é aberto a janela de inconsistências

def check_Files(file):
    if(file.find('.xlsx') or file.find('.xlsm')):                        #Verfica se os arquivos estão no formato correto
       return True
       
    else:
       raise ValueError("Formato .xls não suportado")
   
def check_pattern(letter):                                                # Verfica se o conteúdo da célula é um tag
  pattern= '(\D{2,3})-(\w{5,10})-(\d{2,4})'                               # \D{2,3}= entre 2 a 3 não dígitos -(\w{5,10})- = -entre 5 a 10 alfanúmericos-  \d{2,4} = entre 2 a 4 dígitos
  match = re.search(pattern, letter)
  if (match != None):
     return True
  else:
     return False      

    

def open_Le(element):                                                     # abre a LE e realiza as operações de iterar, testar valores, formatar e guardar o conteúdo em um dicionário
  aux=str()
  theFile = openpyxl.load_workbook(element)
  allSheetNames = theFile.sheetnames
  for sheet in allSheetNames:
     currentSheet = theFile[sheet]
     for row in range(14, currentSheet.max_row + 1):
          cell_name = "{}{}".format("A", row)
          if (currentSheet[cell_name].value != None and  currentSheet[cell_name].value != "-" and check_pattern(currentSheet[cell_name].value) == True):
            aux=formato(currentSheet[cell_name].value)
            if(check_duplicates_Le(aux,cell_name)):
              Verficando[aux]= f'Equipamentos:{cell_name}'

  return Verficando
  
  
def open_PQ(element):                                                     # abre a PQ e realiza as operações de iterar, testar valores, formatar e guardar o conteúdo em um dicionário
  theFile = openpyxl.load_workbook(element)
  allSheetNames = theFile.sheetnames
  for sheet in allSheetNames:
     currentSheet = theFile[sheet]
     for row in range(12, currentSheet.max_row + 1):
        cell_name = "{}{}".format("D", row)
        if (currentSheet[cell_name].value != None and  currentSheet[cell_name].value != "-"):
          aux=currentSheet[cell_name].value[:3] + str(sheet) + "-"+ currentSheet[cell_name].value[3:]
          if (check_pattern(aux)): 
            aux=formato(aux)
            if(check_duplicates_Pq(aux,cell_name)):
              Verficador[aux]= f'Sheet:{sheet}:{cell_name}'
          
  return Verficador  


def formato(let):                                                # formata os tags econtrados visto que podem haver 1 ou 2 zeros antes do número de fato, ex : PB-6027SA-04 ou IT-6027SA-008
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

def check_duplicates_Le(aux,cell):
   if (aux in Verficando):
      doppel[aux]=f'Le:{cell}'
      return False
   else:
      return True
   
def check_duplicates_Pq(aux,cell):
   if (aux in Verficador):
      doppel[aux]= f'PQ:{cell}'
      return False
   else:
      return True

def open_popup():
  if(len(diff)==0 and len(diff1) == 0 and len(doppel)==0):
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
      s2 =  '***Itens Duplicados***\n'
      s2 += '\n'.join([str(i) for i in doppel.items()])
      text= s+'\n'+s1+'\n'+s2
      column = [[pg.Text(text, font=('Courier New', 12))]]
      layout=[
          [pg.Text(f'Inconsistências Encontradas em:', font=('Bold', 16))],
          [pg.Column(column, size=(800, 300), scrollable=True, key = "Column")],
          [pg.Button("OK"), pg.Button("Cancel")]
      ]
      window= pg.Window("Tabela de Inconsistências",layout)
      while True:
          print(window.read())
          break 
       
   
    

