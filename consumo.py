# Lee un archivo CSV con una señal y calcula el area bajo su curva. 

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import pdb
import os

from sklearn.metrics import auc

def leer(filename):
    df = pd.read_csv(filename)
    return df

def area(x,y):
    return auc(x,y) 

def calcular_consumo(cliente,mes):  
    
    formato_fecha = '%d-%m-%Y'
    carpeta = f'D:/01 Findero/{mes}/{cliente}/Datos'
    finderos = os.listdir(carpeta)
    finderos = [item for item in finderos if '.CSV' in item[-4:] or '.csv' in item[-4:]]
    
    resultados = dict()
    inicios = []
    finales = []
    periodos = []
    
    for fin in finderos:
        archivo  = fin
        
        filename = carpeta +'/'+ archivo
        
        df = leer(filename)
        if '00' in archivo:
            df['Datetime'] = df['Date']+' '+df['Time']
        else:
            df['Datetime'] = df['Date']+df['Time']
            
        
        inicial = 0
        final = -1
        inicio_fin = df.iloc[[inicial, final]]
        inicio_fin.reset_index(drop=True, inplace = True)
#        pdb.set_trace()
        inicio = pd.to_datetime(inicio_fin['Datetime'][0] , format = formato_fecha+' %H:%M:%S')
        final = pd.to_datetime(inicio_fin['Datetime'][1] , format = formato_fecha+' %H:%M:%S')
        duracion = final - inicio
        
        mediciones = df.shape[0]
        
        try:
            tiempo_muestreo = df['Milis'].diff()[df['Milis'].diff()>0].mean()/1000
        
        except:
            continue
        
#        segundos = duracion.total_seconds() 
#        frecuencia_muestreo = mediciones/segundos
#        tiempo_muestreo = 1/frecuencia_muestreo

        segundos = mediciones*tiempo_muestreo
        
        horas=round(segundos/3600,2)

#        columna_segundos = np.arange(0,segundos,tiempo_muestreo)
        columna_segundos = np.arange(0,segundos,tiempo_muestreo)

#        pdb.set_trace()
        
#        try:
#            df["Segundos"] = columna_segundos
#        except:
#            df["Segundos"] = columna_segundos[:-1]
        
        contador = 0
        cargas_findero = np.ones(12)
        
        for column in df:
            
            if 'Date' in column or 'Time' in column or 'Segundos' in column or 'Milis' in column:
                continue
            
#            xx=df["Segundos"].values
            xx = columna_segundos
            yy = df[column].values
            
            a = area(xx,yy) #en Joules
            
            kWh = round(a/3600000,2)
                                           
            cargas_findero[contador] = round(kWh,1)
            contador += 1
        
        resultados[fin] = cargas_findero
        
        fecha_inicio = str(inicio)[:-9]
        fecha_fin = str(final)[:-9]
        
        inicios.append(fecha_inicio)
        finales.append(fecha_fin)
        periodos.append(horas)
        
    return resultados,horas,fecha_inicio,fecha_fin,inicios,finales,periodos
        

  
if __name__=='__main__':
    cliente = '11 Gonzalo Celorio'
    mes='07 Julio'

    resultados,horas,fecha_inicio,fecha_fin,inicios,finales,periodos = calcular_consumo(cliente,mes)