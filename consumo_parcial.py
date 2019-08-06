# Lee un archivo CSV con una señal y calcula el area bajo su curva. 

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import pdb
import os
import plotly.graph_objs as go
from plotly.offline import download_plotlyjs, init_notebook_mode, plot, iplot


from sklearn.metrics import auc

def leer(filename):
    df = pd.read_csv(filename)
    try:
        df['Datetime'] = pd.to_datetime(df['Date']+df['Time'], format='%d/%m/%Y  %H:%M:%S')
    except ValueError:
    
        try:
            df['Datetime'] = pd.to_datetime(df['Date']+df['Time'], format='%d-%m-%Y  %H:%M:%S')
        except ValueError:
            df['Datetime'] = pd.to_datetime(df['Date']+df['Time'], format='%d-%m-%y  %H:%M:%S')
    return df

def area(x,y,dx):
    return auc(x,y) 

def calcular_consumo(cliente,mes,findero,columna,inicial,final):  
    
    formato_fecha = '%d-%m-%Y'
    carpeta = 'D:/01 Findero/'+mes+'/'+cliente+'/Datos'
    filename = carpeta +'/'+ findero
    
    df = leer(filename)
    df['datetime']=df['Date']+df['Time']
    largo = df.shape[0]
    idx_inicio = df.index[df['datetime']==inicial].tolist()[0]
    idx_final = df.index[df['datetime']==final].tolist()[0]
    
    saltar = [i for i in range(1,idx_inicio)]
    saltar.extend([i for i in range(idx_final,largo-1)])
    
    df.drop(saltar, inplace=True)
    df.reset_index(inplace=True)
    
    mediciones = df.shape[0]
    
    df.drop([0,mediciones-1], inplace=True)
    mediciones = df.shape[0]
    
    inicio = pd.to_datetime(inicial, format=formato_fecha+' %H:%M:%S')
    final = pd.to_datetime(final, format=formato_fecha+' %H:%M:%S')
    duracion = final - inicio
    
    segundos = duracion.total_seconds() 
    frecuencia_muestreo = mediciones/segundos
    tiempo_muestreo = 1/frecuencia_muestreo
    horas=round(segundos/3600,2)
        
    xx = np.arange(0,segundos,tiempo_muestreo)
        
    if mediciones != len(xx):
        xx = xx[:-1]
            
    yy=df[columna].values
    plt.plot(xx,yy)
    a=area(xx,yy,tiempo_muestreo) #en Joules
    
    kWh=round(a/3600000,2)    
    
    print()
    print(f'Duró {horas} horas')
    print()
    print(f'Consumo: {kWh} kWh')                    
       

#        

  
if __name__=='__main__':
    cliente = '10 Patricia Velasco'
    mes='05 Mayo'
    findero = 'DATALOG_COM20.CSV'
    columna='L'+str(1)
    inicial = '04/05/2019  04:02:23'
    final = '04/05/2019  13:41:54'
#    r,h,i,f = calcular_consumo(cliente,mes,findero,columna,inicial,final)
    calcular_consumo(cliente,mes,findero,columna,inicial,final)