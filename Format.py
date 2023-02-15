import openpyxl
from openpyxl.styles import colors
import PySimpleGUI as pg 


def find_color(color,lista):
      if (color == '00FFCC00'):
          indices = [i for i, x in enumerate(lista) if (len(x)==1)]
          return indices[-1]
      elif(color == 'FFE68D'):
          indices = [i for i, x in enumerate(lista) if (len(x)==3)]
          return indices[-1]
      elif(color == 'FFFF99'):
          indices = [i for i, x in enumerate(lista) if (len(x)==5)]
          return indices[-1]

            

def c_by_v(color,lista):
    if(color=='00FFCC00' and len(lista)==0):
        lista.append("1")
        return str(lista[-1])
    
    elif(color=='00FFCC00' and len(lista)>0):
        value= lista[find_color(color,lista)]
        aux = int(value)
        aux +=1
        string= str(aux)
        lista.append(string)
        return str(lista[-1]) 
     
    elif(color=='FFE68D' and len(lista[-1])==1):
        value = float(lista[-1])
        value += 0.1
        value= round(value,2)
        value=str(value)
        lista.append(value)
        return value
    
    elif(color=='FFE68D' and len(lista[-1])>1):
      value= lista[find_color(color,lista)]
      aux= float(value)
      aux += 0.1
      aux=round(aux,2)
      string=str(aux)
      lista.append(string)
      return str(lista[-1])

    elif(color == 'FFFF99'and len(lista[-1])==3):
        value = (lista[-1])
        value = value + '.1'
        lista.append(value)
        return value
    
    elif(color == 'FFFF99'and len(lista[-1])>3):
        value= lista[find_color(color,lista)]
        aux = int(value[4])
        aux +=1
        string= str(aux)
        value= value[:len(value)-1]+string
        lista.append(value)
        return(value)
    
    elif(color == '000000'and len(lista) >= 3):
      if(len(lista[-1])==5):
        value = (lista[-1])
        value = value + '.1'
        lista.append(value)
        return value
      else:
         value = (lista[-1])
         value=int(value[-1])
         value +=1
         value = str(value)
         aux= lista[-1]
         tam=len(aux)-1
         aux= aux[:tam]+value
         lista.append(aux)
         return aux


def Begin():
  pg.theme("DarkBlue")
  layout=[                                                                                       
              [pg.Text('Planilha de Quantidades: ', size=(13,2))],
              [pg.Input(), pg.FileBrowse(key='-Pq-',file_types=(("Excel", "*.xlsx"),("Excel", "*.xlsm")))],
              [pg.Button("Submit"), pg.Button("Cancel")]
   ]
  window= pg.Window('Verificador',layout, size=(450,150))
  event,values = window.read()
  window.close()
  exit=0
  lista=[]
  theFile = openpyxl.load_workbook(values['-Pq-'])
  allSheetNames = theFile.sheetnames
  for sheet in allSheetNames:
      lista.clear()
      currentSheet = theFile[sheet]
      if(exit==1):
          break
      for row in range(1, currentSheet.max_row + 1):
            cell_name = "{}{}".format("E", row)
            cell_exit= "{}{}".format("I", row)
            cell = currentSheet[cell_name]
            if (cell.fill.start_color.type == 'rgb'):
              color_hex = cell.fill.start_color.rgb[2:] # Remove o caractere '#' do início da string
            else:
              color_hex = colors.COLOR_INDEX[cell.fill.start_color.indexed]
            if(currentSheet[cell_exit].value == None and len(lista) != 0):
                exit=1
                break
            cell.value= c_by_v(color_hex,lista)
  theFile.save(values['-Pq-'])
