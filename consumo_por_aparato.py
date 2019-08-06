# ************************
# Script para calcular el consumo total de un solo aparato. 
# Se basa en encontrar la parte de la señal que está en un rango deseado
# y luego calcular su área. 
# ************************

import pandas as pd
import pdb

def lector(mes, cliente, archivo, puerto, condicion=None):
    
    df = pd.read_csv(f'D:/01 Findero/{mes}/{cliente}/Datos/{archivo}')
    
    seleccion = df[df[df.columns[puerto+2]]>condicion[0]]
    seleccion = seleccion[seleccion[seleccion.columns[puerto+2]]<condicion[1]]
    
    aparato = seleccion[seleccion.columns[puerto+2]]
    aparato.reset_index(inplace=True, drop=True)
    
    tiempo_muestras = .480
    largo_muestras = aparato.shape[0]
    muestras = aparato.index.values+1
    
    tiempo = tiempo_muestras * muestras
    tiempo_total = largo_muestras*tiempo_muestras
    
    potencia = 225
    
    energia = potencia/1000 * tiempo_total/3600
    
#    aparato.plot()
#    print(aparato.head())    
    
    
    return energia, aparato, seleccion
    

if __name__ == '__main__':
    mes = '07 Julio'
    cliente = '11 Gonzalo Celorio'
    archivo = 'DATALOG_F13_Celorio.CSV'
    condicion = [200,270]
    puerto = 4
    
    energia, aparato, data = lector(mes, cliente, archivo, puerto, condicion)
