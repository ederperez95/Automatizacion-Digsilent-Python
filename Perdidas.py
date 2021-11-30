# -*- coding: utf-8 -*-
"""
Created on Tue Nov 30 10:42:22 2021

@author: edjperez
"""

###############################################    Librerias    ###################################################################################################

import powerfactory
import time
import pandas as pd



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
ruta = script.ruta
dataFrame = pd.DataFrame()


#obtiene los objetos de los sets del powerFactory      
for i in script.GetContents():    
    if i.loc_name == 'Casos de estudio':
        set_Casos_de_estudio = i.All()
    elif i.loc_name == 'Contingencias':
        set_Contingencias = i.All()
    elif i.loc_name == 'Elementos_a_monitorear':   
        set_Elementos_a_monitorear = i.All()
        
        
for casoDeEstudio in set_Casos_de_estudio:
    
    casoDeEstudio.Activate()
    
    casoDeEstudioName = casoDeEstudio.loc_name
    for contingencia in set_Contingencias:
        
        estadoInicial = contingencia.outserv
        
        perdidasTotales = 0
        
        Ldf.Execute()
        
        for elemento in set_Elementos_a_monitorear:
            
            try:
                perdidasTotales += elemento.GetAttribute('c:Ploss')
            except:
                pass
        
        try:
            manejoDeContingencias(contingencia, estadoInicial, 'apagar')
            
            perdidasTotales = 0
            
            Ldf.Execute()
            
            for elemento in set_Elementos_a_monitorear:
                try:
                    perdidasTotales += elemento.GetAttribute('c:Ploss')
                except:
                    pass
               
            app.PrintPlain(perdidasTotales)
            
            
            manejoDeContingencias(contingencia, estadoInicial, 'encender')
        except:
            pass
            
# app.PrintPlain(set_Elementos_a_monitorear)



app.EchoOn()
app.PrintPlain(str((time.time() - start_time)) + ' segundos de ejecucion')  


        


    
    
    

