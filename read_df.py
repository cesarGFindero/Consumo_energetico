import pandas as pd
import pdb

def lector(mes, cliente, archivo, puerto, condicion=None):
    
    df = pd.read_csv(f'D:/01 Findero/{mes}/{cliente}/Datos/{archivo}')
#    pdb.set_trace()
    if condicion is not None:
        seleccion = df[df[df.columns[puerto+2]]>condicion[0]]
        seleccion = seleccion[seleccion[seleccion.columns[puerto+2]]<condicion[1]]
    else:
        seleccion = df[df.columns[puerto+2]]
        
    seleccion[df.columns[puerto+2]].plot()
    print(seleccion.head())    
    
    return seleccion
    

if __name__ == '__main__':
    mes = '07 Julio'
    cliente = '11 Gonzalo Celorio'
    archivo = 'DATALOG_F13_Celorio.CSV'
    condicion = [200,270]
    puerto = 4
    
    data = lector(mes, cliente, archivo, puerto, condicion)
    