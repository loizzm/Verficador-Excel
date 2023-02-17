import openpyxl
from openpyxl.styles import colors
import PySimpleGUI as pg 

allowed_colors=['FF9900','00FF9900','FFCC00','00FFCC00','FFE68D','FFFF99','000000']  ##Cores permitidas pela SPE

def find_color(color,lista):                                            ## Serve para procurar a última vez que uma cor foi encontrada, visto que normalmente existe itens em branco antes das cores principais
      if (color=='FF9900' or color=='00FF9900'):                        ## Nõa há branco descrito nessa função, pois nunca o branco virá entre duas cores diferentes de branco
          indices = [i for i, x in enumerate(lista) if (len(x)==1)]     ## retorna os indíces de todas as vezes que essa cor foi encontrada 
          return indices[-1]                                            ## retorna o último índide onde a cor foi encontrada
      elif(color=='FFCC00'or color=='00FFCC00'):
          indices = [i for i, x in enumerate(lista) if (len(x)==3)]
          return indices[-1]
      elif(color == 'FFE68D'):
          indices = [i for i, x in enumerate(lista) if (len(x)==5)]
          return indices[-1]
      elif(color == 'FFFF99'):
          indices = [i for i, x in enumerate(lista) if (len(x)==7)]
          return indices[-1]

            

def c_by_v(color,lista):                                                 ## Retorna o valor destinado a célula baseado na cor e na posição dela na lista
    if(len(lista)==0):
        if(color=='FF9900' or color=='00FF9900'):                        ## Para a priemeira vez que a cor laranja for encontrada
            lista.append("1")
            return str(lista[-1])
    else:
        if(color=='FF9900' or color=='00FF9900'):                        ## Para a segunda vez que a cor laranja for encontrada, antecipada por uma célula branca
            value= lista[find_color(color,lista)]                        ## usada pra pegar a última aprição de uma cor laranja e pegar seu valor
            aux = int(value)                                             ## Transfomra "1" em 1
            aux +=1                                                      ## acrescenta
            string= str(aux)                                             ## Transforma em string dnv
            lista.append(string)                                         ## acrescenta na lista
            return str(lista[-1]) 
        
        elif((color=='FFCC00'or color=='00FFCC00') and len(lista[-1])==1):   ##  a primeira aparição de uma cor laranja mais amarelada, sempre é antecipada pelo laranja forte por isso o len(lista[-1])==1
            value = float(lista[-1])                                         ## transforma em float o último valor da lista
            value += 0.1                                                     ## Faz 1.0 virar 1.1
            value= round(value,2)                                            ## arredonda para duas casas decimais já que transfomações pra float podem ter erros de arredondamneto                             
            value=str(value)                                                 ## Transforma em string novamente
            lista.append(value)
            return value
        
        elif((color=='FFCC00'or color=='00FFCC00') and len(lista[-1])>=9):    ## para uma aparição da cor laranja mais amarelado depois de um branco, por isso o len do úlitmo elemneto regisrado na lista tem que ser maior que 9
            value= lista[find_color(color,lista)]                            ## retorna o úlitmo indice da lista onde essa cor foi encontrada
            aux= float(value)                                                ## Transfoma em float
            aux += 0.1                                                       ## soma 0,1 n0 valor anterior, ou seja "1.1 + 0.1 = 1.2"
            aux=round(aux,2)                                                 ## Arredonda para 2 casas decimais 
            string=str(aux)                                                  ## Transforma em string
            lista.append(string)
            return str(lista[-1])
        
        elif(color=='FFE68D' and len(lista[-1])==3):                        ## Primeira aparição do amarelo, logo depois de um laranja mais amarelad0, por isso o len do úlitmo item da lista tem que ter tamanho 3("1.1")
            value = (lista[-1])
            value = value + '.1'
            lista.append(value)
            return value
        
        elif(color=='FFE68D' and len(lista[-1])>=9):                        ## Para um aparição depois de uma célula branca, por isso o len do úlitmo item da lista tem que ter tamanho maior que 9
            value= lista[find_color(color,lista)]
            aux = int(value[4])                                            ##Pega o último valor da última aparição registrada dessa cor, por isso o 4 ("1.1.1")
            aux +=1
            string= str(aux)
            value= value[:len(value)-1]+string
            lista.append(value)
            return(value)

        elif(color == 'FFFF99'and len(lista[-1])==5):                     ##Para uma primeira aparição do amarelo mais fraco, vindo depois de um amrelo, por isso o len ==5 (1.1.1)
            value = (lista[-1])
            value = value + '.1'
            lista.append(value)
            return value
        
        elif(color == 'FFFF99'and len(lista[-1])>=9):                   ## Para uma aparição do amareloa mais fraco depois de um branco, por isso o len > 9
            value= lista[find_color(color,lista)]
            aux = int(value[6])
            aux +=1
            string= str(aux)
            value= value[:len(value)-1]+string
            lista.append(value)
            return(value)
        
        elif(color == '000000'and len(lista) >= 4):                  ## Para uma parição de células Brancas
            if(len(lista[-1])==7):
                value = (lista[-1])
                value = value + '.1'
                lista.append(value)
                return value
            else:
                value = (lista[-1])
                if(len(value)==9):
                    value=int(value[-1])
                    value +=1
                    value = str(value)
                    aux= lista[-1]
                    tam=len(aux)-1
                    aux= aux[:tam]+value
                    lista.append(aux)
                    return aux
                elif(len(value)==10):
                    value=int(value[8:])
                    value +=1
                    value = str(value)
                    aux= lista[-1]
                    tam=len(aux)-2
                    aux= aux[:tam]+value
                    lista.append(aux)
                    return aux
                else:
                    value=int(value[8:])
                    value +=1
                    value = str(value)
                    aux= lista[-1]
                    tam=len(aux)-3
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
  lista=[]
  theFile = openpyxl.load_workbook(values['-Pq-'])
  allSheetNames = theFile.sheetnames
  for sheet in allSheetNames:
      lista.clear()                                    ## Reseta a lista para ser preenchida novamente na próxima sheet
      currentSheet = theFile[sheet]
      if (currentSheet.sheet_state == 'hidden' or currentSheet.sheet_state == 'veryHidden'):
          continue
      for row in range(1, currentSheet.max_row + 1):
            cell_name = "{}{}".format("A", row)
            cell_exit= "{}{}".format("I", row)
            cell = currentSheet[cell_name]
            if (type(cell).__name__ != 'MergedCell'):
                if (cell.fill.start_color.type == 'rgb'):
                    color_hex = cell.fill.start_color.rgb[2:] # Remove o caractere '#' do início da string
                else:
                    if (cell.fill.start_color.type == 'indexed'):
                        indexed_color = cell.fill.start_color.indexed
                        color_hex = openpyxl.styles.colors.COLOR_INDEX[indexed_color]
                if(currentSheet[cell_exit].value == None and len(lista) != 0):                          ##Para resolver o problema do código escrever onde não devia
                    break
                if(color_hex in allowed_colors):
                    cell.value= c_by_v(color_hex,lista)
                else:
                    pg.popup_error(f'Cores divergentes da SPE, favor verificar\n Planilha:{currentSheet} célula:A{row}')
                    break
  pg.popup_auto_close("Feito!!")
  theFile.save(values['-Pq-'])
