#*********************
# Interfaz para hacer la limpieza de datos
# usando el acomodo de clases y objetos
#*********************

import tkinter as tk

import os
from datetime import datetime
from consumo import *
from functools import partial
 
class Ventana(tk.Tk):
    
    def __init__(self, master):
        tk.Tk.__init__(self, master)
        self.master = master
        self.title('Consumo por circuito')
        self.iconbitmap('favicon.ico')
        self.geometry('1000x900+500+60')
        self.resizable(0,0)
        menu = tk.Menu(self)
        self.config(menu=menu)

        subMenu = tk.Menu(menu,tearoff=False)
        menu.add_cascade(label = 'Archivo', menu = subMenu)
        subMenu.add_command(label = 'Cerrar', command=self.cerrar)
        
        self.mainWidgets()
        
       
    def mainWidgets(self):
        self.cuerpo = Cuerpo(self)
        self.cuerpo.grid(row=0, column=0)
        
#        self.titulo = Titulo(self)
#        self.titulo.grid(row=2, column=1)
     
    def cerrar(self):
        self.destroy()


class Cuerpo(tk.Frame):
    
    carpeta_in = 'D:/01 Findero'
    meses_lista = {1:'01 Enero',2:'02 Febrero',3:'03 Marzo',
                   4:'04 Abril',5:'05 Mayo',6:'06 Junio',
                   7:'07 Julio',8:'08 Agosto',9:'09 Septiembre',
                   10:'10 Octubre',11:'11 Noviembre',12:'12 Diciembre'}
    
    def __init__(self, master):
        tk.Frame.__init__(self,master)
        self.master = master
        self.widgets()        

    def widgets(self):

        self.label_encabezado = tk.Label(self,  text = 'Consumo por circuito',
                                         font='Helvetica 18 bold' )
        self.label_encabezado.grid(row=0, column=1, columnspan = 2)
        
        self.label_mes = tk.Label(self, text ='Seleccionar el mes:')
        self.label_mes.grid(row=1, column=0, padx=20, pady=20, sticky=tk.E)
        
        self.label_cliente = tk.Label(self, text = 'Seleccionar el cliente:')
        self.label_cliente.grid(row=2, column=0, padx=20, pady=20, sticky=tk.E)

        self.tk_mes = tk.StringVar()
        self.tk_mes.set(self.meses_lista[datetime.now().month])        
        self.meses = {m for m in os.listdir(self.carpeta_in) if m in self.meses_lista.values()}
        
        self.mes_desplegable = tk.OptionMenu(self, self.tk_mes, *self.meses)
        self.mes_desplegable.config(width=20)
        self.mes_desplegable.grid(row=1,column=1)
        
        self.tk_cliente = tk.StringVar()        
        self.clientes = os.listdir(self.carpeta_in+'/'+ self.tk_mes.get())
        self.tk_cliente.set(self.clientes[-1])
        self.tk_mes.trace('w',self.update)
        
        self.cliente_desplegable = tk.OptionMenu(self, self.tk_cliente, *self.clientes)
        self.cliente_desplegable.config(width=20)
        self.cliente_desplegable.grid(row=2,column=1)
        
        
        self.finderos = [item for item in os.listdir(self.carpeta_in+'/'+ self.tk_mes.get()+'/'+self.tk_cliente.get()+'/Datos') if '.CSV' in item[-4:] or '.csv' in item[-4:]]

        self.boton_enviar = tk.Button(self, text = 'Calcular consumo', command=self.enviar)
        self.boton_enviar.grid(row=4, column=0, columnspan = 2)
        
        self.seleccion = tk.StringVar()
        self.seleccion.set('')
        self.seleccion.trace('w',self.update_radio)       

                      
        self.grid_rowconfigure(5, minsize=40)
        self.grid_rowconfigure(3, minsize=20)
        self.grid_columnconfigure(2, minsize=150)
        self.grid_columnconfigure(4, minsize=20)
        self.grid_rowconfigure(8, minsize=50)
        self.grid_rowconfigure(10, minsize=40)
        self.grid_rowconfigure(12, minsize=40)
        self.grid_rowconfigure(10, minsize=40)
        
        
            
    def enviar(self):
        self.finderos = [item for item in os.listdir(self.carpeta_in+'/'+ self.tk_mes.get()+'/'+self.tk_cliente.get()+'/Datos') if '.CSV' in item[-4:] or '.csv' in item[-4:]]

        self.diccionario, self.horas, self.inicio, self.final = calcular_consumo(self.tk_cliente.get(), self.tk_mes.get())
        
        self.label_nombre = tk.Label(self, text=self.tk_cliente.get()[3:], 
                                     font='Helvetica 20 bold')
        self.label_nombre.grid(row=6, column=1, columnspan = 2)


        self.label_horas = tk.Label(self, text= 'Duración:')
        self.label_horas.grid(row=1, column=3, columnspan = 2, sticky=tk.E)
        self.label_horas = tk.Label(self, text= self.horas)
        self.label_horas.grid(row=1, column=5,sticky=tk.W)       
        
        self.label_inicio = tk.Label(self, text= 'Inicio:')
        self.label_inicio.grid(row=2, column=3, columnspan = 2, sticky=tk.E)
        self.label_inicio = tk.Label(self, text= self.inicio)
        self.label_inicio.grid(row=2, column=5,sticky=tk.W)

        self.label_final = tk.Label(self, text= 'Fin:')
        self.label_final.grid(row=3, column=3, columnspan = 2, sticky=tk.E)
        self.label_final = tk.Label(self, text= self.final)
        self.label_final.grid(row=3, column=5,sticky=tk.W)
        
        
        nombres = self.finderos
        self.seleccion.set(nombres[0])
        i=1
        for name in nombres:
            b = tk.Radiobutton(self, text=name[8:-4], variable=self.seleccion, value=name)
            b.grid(row=7, column=i)
            i+=1
        
        
        
    def update(self, *args):
        client = os.listdir(self.carpeta_in+'/'+ self.tk_mes.get())
        try:
            self.tk_cliente.set(client[-1])
        except:
            self.tk_cliente.set('')      
        
        
        menu = self.cliente_desplegable['menu']
        menu.delete(0,'end')
    
        for name in client:
            menu.add_command(label=name, command=lambda nuevo=name: self.tk_cliente.set(nuevo))
    
    
    
    def update_radio(self, *args):
        
        seleccionado = self.diccionario[self.seleccion.get()]
        
        self.label_findero = tk.Label(self, text='Findero:  '+ self.seleccion.get()[8:11], font = 'Helvetica 15 bold')
        self.label_findero.grid(row=8,column=0,columnspan = 2)
        
        self.label_c1 = tk.Label(self, text='Circuito 1', bg='light gray')
        self.label_c2 = tk.Label(self, text='Circuito 2', bg='light gray')
        self.label_c3 = tk.Label(self, text='Circuito 3', bg='light gray')
        self.label_c4 = tk.Label(self, text='Circuito 4', bg='light gray')
        self.label_c5 = tk.Label(self, text='Circuito 5', bg='light gray')
        self.label_c6 = tk.Label(self, text='Circuito 6', bg='light gray')
        self.label_c7 = tk.Label(self, text='Circuito 7', bg='light gray')
        self.label_c8 = tk.Label(self, text='Circuito 8', bg='light gray')
        self.label_c9 = tk.Label(self, text='Circuito 9', bg='light gray')
        self.label_c10 = tk.Label(self, text='Circuito 10', bg='light gray')
        self.label_c11 = tk.Label(self, text='Circuito 11', bg='light gray')
        self.label_c12 = tk.Label(self, text='Circuito 12', bg='light gray')
        
        self.label_c1_consumo = tk.Label(self, text=seleccionado[0], bg='white')
        self.label_c2_consumo = tk.Label(self, text=seleccionado[1], bg='white')
        self.label_c3_consumo = tk.Label(self, text=seleccionado[2], bg='white')
        self.label_c4_consumo = tk.Label(self, text=seleccionado[3], bg='white')
        self.label_c5_consumo = tk.Label(self, text=seleccionado[4], bg='white')
        self.label_c6_consumo = tk.Label(self, text=seleccionado[5], bg='white')
        self.label_c7_consumo = tk.Label(self, text=seleccionado[6], bg='white')
        self.label_c8_consumo = tk.Label(self, text=seleccionado[7], bg='white')
        self.label_c9_consumo = tk.Label(self, text=seleccionado[8], bg='white')
        self.label_c10_consumo = tk.Label(self, text=seleccionado[9], bg='white')
        self.label_c11_consumo = tk.Label(self, text=seleccionado[10], bg='white')
        self.label_c12_consumo = tk.Label(self, text=seleccionado[11], bg='white')
        
        self.label_c1.grid(row=9,column=0)
        self.label_c2.grid(row=9,column=2)
        self.label_c3.grid(row=9,column=4)
        self.label_c4.grid(row=9,column=6)
        
        self.label_c5.grid(row=11,column=0)
        self.label_c6.grid(row=11,column=2)
        self.label_c7.grid(row=11,column=4)
        self.label_c8.grid(row=11,column=6)
        
        self.label_c9.grid(row=13,column=0)
        self.label_c10.grid(row=13,column=2)
        self.label_c11.grid(row=13,column=4)
        self.label_c12.grid(row=13,column=6)
        
        
        self.label_c1_consumo.grid(row=9,column=1)
        self.label_c2_consumo.grid(row=9,column=3)
        self.label_c3_consumo.grid(row=9,column=5)
        self.label_c4_consumo.grid(row=9,column=7)
        
        self.label_c5_consumo.grid(row=11,column=1)
        self.label_c6_consumo.grid(row=11,column=3)
        self.label_c7_consumo.grid(row=11,column=5)
        self.label_c8_consumo.grid(row=11,column=7)
        
        self.label_c9_consumo.grid(row=13,column=1)
        self.label_c10_consumo.grid(row=13,column=3)
        self.label_c11_consumo.grid(row=13,column=5)
        self.label_c12_consumo.grid(row=13,column=7)
        
        
        
class Resultados(tk.Frame):
    
    def __init__(self, master):
        tk.Frame.__init__(self,master)
        self.master = master
        self.widgets() 

        
consumo_ = Ventana(None)

consumo_.mainloop()