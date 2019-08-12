# Lee un archivo CSV con una señal y calcula el area bajo su curva. 

import numpy as np
import pandas as pd
import pdb
import os
from sklearn.metrics import auc


def calcular_consumo(cliente,mes):  
    
    formato_fecha = '%d-%m-%Y'
    carpeta = f'D:/01 Findero/{mes}/{cliente}/Datos'
    finderos = os.listdir(carpeta)
    finderos = [item for item in finderos if '.CSV' in item[-4:] or '.csv' in item[-4:]]
    
    consumos = dict()
    inicios = []
    finales = []
    periodos = []
    
    for archivo in finderos:
        
        filename = f'{carpeta}/{archivo}'
        
        df = pd.read_csv(filename)
                 
        
        inicial = 0
        final = -1
        inicio_fin = df.iloc[[inicial, final]][['Date','Time']]
        inicio_fin.reset_index(drop=True, inplace = True)
#        pdb.set_trace()
        inicio_fin['Datetime'] = inicio_fin['Date']+inicio_fin['Time']
        
        inicio = pd.to_datetime(inicio_fin['Datetime'][0] , format = formato_fecha+' %H:%M:%S')
        final = pd.to_datetime(inicio_fin['Datetime'][1] , format = formato_fecha+' %H:%M:%S')
       
        duracion = final - inicio
        duracion = duracion.total_seconds()
        
        mediciones = df.shape[0]
        
        try:
            tiempo_muestreo = df['Milis'].diff()[df['Milis'].diff()>0].mean()/1000    
        except:
            tiempo_muestreo = duracion/mediciones
        

        segundos = mediciones*tiempo_muestreo
        
        horas = round(segundos/3600,2)

        arreglo_segundos = np.linspace(0,segundos,mediciones)
        
        cargas_puertos = []
        for column in df:
            
            if 'Date' in column or 'Time' in column or 'Milis' in column:
                continue
            
            tiempo = arreglo_segundos
            potencia = df[column].values
            
            
            area_curva = auc(tiempo,potencia)
            
            kWh = round(area_curva/3600000,1)
                                           
            cargas_puertos.append(kWh)

        
        consumos[archivo] = cargas_puertos
        
        fecha_inicio = str(inicio)[:-9]
        fecha_fin = str(final)[:-9]
        
        inicios.append(fecha_inicio)
        finales.append(fecha_fin)
        periodos.append(horas)
        
    return consumos,horas,fecha_inicio,fecha_fin,inicios,finales,periodos
        

  
if __name__=='__main__':
    cliente = '29 Diego Medina'
    mes='07 Julio'

    resultados,horas,fecha_inicio,fecha_fin,inicios,finales,periodos = calcular_consumo(cliente,mes)