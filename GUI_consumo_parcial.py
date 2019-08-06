#*******************************
# Interfaz grafica para calcular el consumo parcial 
#********************************


import tkinter as tk

import os
from datetime import datetime
import consumo_parcial as cp
import pdb

class Ventana(tk.Tk):
    
    def __init__(self, master):
        tk.Tk.__init__(self, master)
        self.master = master
        self.title("Consumo parcial")
        self.iconbitmap('favicon.ico')
        self.geometry('500x600+600+200')
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

        
#        self.mensajes = Mensajes(self)
#        self.mensajes.grid(row=3, column=0,columnspan = 3, sticky=tk.W)
        
    
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
        
        self.tk_mes = tk.StringVar()
        self.tk_mes.set(self.meses_lista[datetime.now().month])
        self.tk_cliente = tk.StringVar()
        
        self.label_encabezado = tk.Label(self,  text = '         Consumo Parcial',
                                         font='Helvetica 18 bold' )
        self.label_encabezado.grid(row=1, column=1, columnspan = 2)
        
        self.label_mes = tk.Label(self, text ='Mes:')
        self.label_mes.grid(row=3, column=1, padx=20, pady=20, sticky=tk.E)
        
        self.label_cliente = tk.Label(self, text = 'Cliente:')
        self.label_cliente.grid(row=4, column=1, padx=20, pady=20, sticky=tk.E)
        
        self.meses = {m for m in os.listdir(self.carpeta_in) if m in self.meses_lista.values()}
        
        self.mes_desplegable = tk.OptionMenu(self, self.tk_mes, *self.meses)
        self.mes_desplegable.config(width=20)
        self.mes_desplegable.grid(row=3,column=2)
        
        self.clientes = os.listdir(self.carpeta_in+'/'+ self.tk_mes.get())
        self.tk_cliente.set(self.clientes[-1])
        self.tk_mes.trace('w',self.update)
        
        self.cliente_desplegable = tk.OptionMenu(self, self.tk_cliente, *self.clientes)
        self.cliente_desplegable.config(width=20)
        self.cliente_desplegable.grid(row=4,column=2)
        
        
        self.label_findero = tk.Label(self, text = 'Findero:')
        self.label_findero.grid(row=5, column=1, padx=20, pady=20, sticky=tk.E)
        self.tk_findero = tk.StringVar()        
        self.finderos = [item[8:-4] for item in os.listdir(self.carpeta_in+'/'+ self.tk_mes.get()+'/'+self.tk_cliente.get()+'/Datos')
                                                                    if '.CSV' in item[-4:] or '.csv' in item[-4:]]
        
        self.tk_findero.set(self.finderos[0])
        self.tk_cliente.trace('w',self.update_findero)
        
        self.findero_desplegable = tk.OptionMenu(self, self.tk_findero, *self.finderos)
        self.findero_desplegable.config(width=20)
        self.findero_desplegable.grid(row=5,column=2)
        
        self.label_columna = tk.Label(self, text='Puerto:')
        self.label_columna.grid(row=6, column=1, padx=20, pady=20, sticky=tk.E)
        
        columna_ = tk.StringVar(self, value='1')
        self.entrada_columna = None
        self.entrada_columna = tk.Entry(self,textvariable = columna_)
        self.entrada_columna.grid(row=6,column=2)
        self.columna = 'L'+self.entrada_columna.get()
        
        self.label_inicio = tk.Label(self, text='Inicio:')
        self.label_inicio.grid(row=7, column=1, padx=20, pady=20, sticky=tk.E)
        
        self.inicial = tk.StringVar(self,value='20-07-2019  08:30:00')
        self.entrada_inicio = tk.StringVar(self, value='')
        self.entrada_inicio = tk.Entry(self, textvariable=self.inicial )
        self.entrada_inicio.grid(row=7,column=2)
        
        self.label_final = tk.Label(self, text='Final:')
        self.label_final.grid(row=8, column=1, padx=20, pady=20, sticky=tk.E)
        
        self.final = tk.StringVar(self,value='20-07-2019  14:00:00')
        self.entrada_final = None
        self.entrada_final = tk.Entry(self, textvariable=self.final )
        self.entrada_final.grid(row=8,column=2)
        
       
        self.boton_enviar = tk.Button(self, text = 'Calcular', command=self.enviar)
        self.boton_enviar.grid(row=10, column=1, columnspan = 2, ipadx=15)
        
        self.grid_rowconfigure(0, minsize=20)
        self.grid_rowconfigure(2, minsize=20)
        self.grid_rowconfigure(9, minsize=20)
        
        self.grid_columnconfigure(0, minsize=55)
        
    def enviar(self):
#        pdb.set_trace()
        self.entrada_inicio = tk.Entry(self, textvariable=self.inicial )
        self.entrada_final = tk.Entry(self, textvariable=self.final )
        self.columna = 'L'+self.entrada_columna.get()
        findero = 'DATALOG_'+self.tk_findero.get()+'.CSV' 
        cp.calcular_consumo( self.tk_cliente.get(), self.tk_mes.get(),findero,self.columna,
                                self.entrada_inicio.get(),self.entrada_final.get())
        
        del self.entrada_inicio
        del self.entrada_final

        
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
                 
    
    def update_findero(self, *args):
        finderos = [item[8:-4] for item in os.listdir(self.carpeta_in+'/'+ self.tk_mes.get()+'/'+self.tk_cliente.get()+'/Datos')
                                                                    if '.CSV' in item[-4:] or '.csv' in item[-4:]]
        try:
            self.tk_findero.set(finderos[0])
        except:
            self.tk_findero.set('')      
        
        
        menu = self.findero_desplegable['menu']
        menu.delete(0,'end')
    
        for name in finderos:
            menu.add_command(label=name, command=lambda nuevo=name: self.tk_findero.set(nuevo))
                
    

        

parcial = Ventana(None)

parcial.mainloop()