# -*- coding: utf-8 -*-
"""
Created on Tue Nov 30 10:42:22 2021

@author: edjperez
"""

###############################################    Librerias    ###################################################################################################

import powerfactory
import time
import pandas as pd
import os



###############################################    Funciones    ###################################################################################################




def manejoDeContingencias(elemento, estadoInicial, accion):
    
    if accion == 'encender':
        elemento.outserv = estadoInicial
    else:
        elemento.outserv = 1







###############################################    Codigo    ###################################################################################################

app = powerfactory.GetApplication()
start_time = time.time()
app.EchoOff()
script = app.GetCurrentScript()        
app.GetAllUsers(1)
reloj = app.GetFromStudyCase('*.SetTime')
Ldf = app.GetFromStudyCase('*.ComLdf')
Ldf.iopt_check = 0
ruta = script.Ruta
dataFrame = pd.DataFrame(columns=['Caso de estudio', 'Escenario', 'Perdidas'])



#obtiene los objetos de los sets del powerFactory      
for i in script.GetContents():    
    if i.loc_name == 'Casos de estudio':
        set_Casos_de_estudio = i.All()
    elif i.loc_name == 'Contingencias':
        set_Contingencias = i.All()
    elif i.loc_name == 'Elementos_a_monitorear':   
        set_Elementos_a_monitorear = i.All()
        
        
for casoDeEstudio in set_Casos_de_estudio:
    
    
    
    # Activar caso de estudio
    casoDeEstudio.Activate()
    
    
    # guardar nombre del caso de estudio
    casoDeEstudioName = casoDeEstudio.loc_name
    
    # Escenario operativo 
    escenarioActivo = app.GetActiveScenario() 
    
    # variables para posteriormente ser guardadas en el dataframe
    index = 0
    indice = []
    sumaPerdidas = []
    temp = []
    
    # flujo de cargas
    Ldf.Execute()    
    
    # Suma de las perdidas por las lineas de transmision
    perdidasTotales = 0    
    for elemento in set_Elementos_a_monitorear:
        
        try:
            perdidasTotales += elemento.GetAttribute('c:Ploss')
        except:
            pass
    
    # Agregar los valores hallados a los vectores
    indice.append((casoDeEstudioName, 'Caso Base'))
    sumaPerdidas.append(perdidasTotales)
    dataFrame.loc[index] = [casoDeEstudioName, 'Caso Base', perdidasTotales]
    index += 1
    # imprime valor de las perdidas del caso base
    app.PrintPlain(str(perdidasTotales) + ' Caso base')
    
    # contingencias 
    for contingencia in set_Contingencias:
        
        # Hace la contingencia (saca el elemento de servivio)
        estadoInicial = contingencia.outserv
        try:
            
            manejoDeContingencias(contingencia, estadoInicial, 'apagar')
           
            # ejecuta flujo de cargas
            Ldf.Execute()
            
            # Suma de las perdidas por las lineas de transmision
            perdidasTotales = 0
            for elemento in set_Elementos_a_monitorear:
                try:
                    perdidasTotales += elemento.GetAttribute('c:Ploss')
                except:
                    pass
            
            # Agregar los valores hallados a los vectores
            indice.append((casoDeEstudioName, elemento.loc_name))
            sumaPerdidas.append(perdidasTotales)
            dataFrame.loc[index] = [casoDeEstudioName, contingencia.loc_name, perdidasTotales]
            index += 1
            
            
            
            # imprime valor de las perdidas de la contingencia
            app.PrintPlain(str(perdidasTotales) + ' ' + contingencia.loc_name)
            
            # Vuleve y pone en servicio el elemento
            manejoDeContingencias(contingencia, estadoInicial, 'encender')
            
        except:
            pass
        
    escenarioActivo.DiscardChanges()
            
# app.PrintPlain(set_Elementos_a_monitorear)


dataFrame.to_excel(os.path.join(ruta, 'prueba.xlsx'))
app.EchoOn()
app.PrintPlain(str((time.time() - start_time)) + ' segundos de ejecucion')  



        


    
    
    

