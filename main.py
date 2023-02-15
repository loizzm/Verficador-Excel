import PySimpleGUI as pg 
import Functions
import MoreFunctions
import Format


pg.theme("DarkBlue")
layout=[                                                                                      
              [pg.Text('Escolha a funcionalidade:', size=(13,2))],
              [pg.Combo(['Verificar LE/PQ', 'Verificar Lista de demanda', 'Formatar PQ'],default_value='Verificar LE/PQ', key='board')],
              [pg.Button("Submit"), pg.Button("Cancel")] 
       ]
window= pg.Window('Verificador',layout, size=(450,150), element_justification='c')
event,values = window.read()
if (event == 'Submit'):
    if(values['board']=="Verificar LE/PQ"):
        Functions.Begin()
    elif(values['board']=="Formatar PQ"):
        Format.Begin()
    else:
        MoreFunctions.Begin()
window.close()