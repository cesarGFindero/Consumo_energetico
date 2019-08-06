# ****************************
# En este programa se rellena automaticamente una hoja en excel
#*****************************

import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from xlsxwriter.utility import xl_cell_to_rowcol
import pdb 

def sumatoria(cargas):
    suma = '='
    for celda in cargas:
        suma = suma + celda + '+'
    return suma[:-1]

def suma_varias(*args):
    suma = '='
    for celda in args:
        suma = suma + celda + '+'
    return suma[:-1]

def producto_varias(*args):
    producto = '='
    for celda in args:
        producto = producto + celda + '*'
    return producto[:-1]

def reacomodo_fechas(fecha):
    nueva = fecha[-2:] + '-' + fecha[-5:-3] + '-' + fecha[0:4]
    return nueva
        

def crear_vlookup(filas,aparato):
    formula = '='
    for f in range(filas):
        for c in range(0,12):
            esquina_1 = xl_rowcol_to_cell(10+11*f,1+4*c)
            esquina_2 = xl_rowcol_to_cell(13+11*f,2+4*c)
            formula+='IFERROR(VLOOKUP("'+aparato+'",Detalles!'+esquina_1+':'+esquina_2+',2,FALSE),0)'+'+'
    return formula[:-1] 
    

def general(datos,horas,precio,inicio,final,cliente,mes,nombre,workbook):
    
    worksheet = workbook.add_worksheet('General')
    
    celda_titulo = 'B3'
    
    titulo = workbook.add_format({'bold': True,'font_size':20})
    
    worksheet.write(celda_titulo,'Desciframiento del consumo de ' + nombre, titulo)


def detalles(datos,horas,precio,inicio,final,cliente,mes,nombre,workbook,incios,finales,periodos):
    
    worksheet = workbook.add_worksheet('Detalles')   
    
    celda_titulo = 'A1'
    celda_periodo = 'H2'
    celda_inicio_titulo = 'H3'
    celda_final_titulo = 'H4'
    celda_horas = 'I2'
    celda_inicio = 'I3'
    celda_final = 'I4'
    columna_nombre_finderos = 0
    fila_primer_findero = 4
    encabezado = 2
    espacio_vertical = 11
    espacio_horizontal = 4
    
    bold = workbook.add_format({'bold': True})
    bold_titulo = workbook.add_format({'bold': True,'font_size':15})
    centrado = workbook.add_format({'align':'center'})
    encabezados = workbook.add_format({'bold': True,'align':'center'})
    formato_kWh = workbook.add_format({'num_format': '0.0','align':'center'})
    money = workbook.add_format({'num_format': '_-$* #,##0.00_-;-$* #,##0.00_-;_-$* "-"??_-;_-@_-'})
    money_miles = workbook.add_format({'num_format': '$#,##0.00'})
    porcentaje = workbook.add_format({'num_format': '0.00%','align':'center'})
    condicional_amarillo = workbook.add_format({'bg_color':   '#FFEB9C',
                               'font_color': '#9C6500'})
    condicional_rojo = workbook.add_format({'bg_color':   '#FFC7CE',
                               'font_color': '#9C0006'})
                                       
    worksheet.write(celda_titulo,'Detalles del consumo de ' + nombre, bold_titulo)
    worksheet.write(celda_periodo,'Periodo:')
    worksheet.write(celda_horas,horas)
    
    worksheet.write(celda_inicio_titulo,'Inicio:')
    worksheet.write(celda_inicio,reacomodo_fechas(inicio))
    
    worksheet.write(celda_final_titulo,'Final:')
    worksheet.write(celda_final,reacomodo_fechas(final))

    
    
    celdas_finderos = []
    for indice,findero in enumerate(list(datos.keys())):
        celdas_cargas = []
        worksheet.merge_range(fila_primer_findero+indice*espacio_vertical,columna_nombre_finderos,
                              fila_primer_findero+indice*espacio_vertical,columna_nombre_finderos+encabezado,'Findero: '+findero[8:-4],encabezados)
        worksheet.write(fila_primer_findero+indice*espacio_vertical,columna_nombre_finderos+3,'Periodo:')
        worksheet.write(fila_primer_findero+indice*espacio_vertical,columna_nombre_finderos+5,'Inicio:')
        worksheet.write(fila_primer_findero+indice*espacio_vertical,columna_nombre_finderos+7,'Final:')
        
        worksheet.write(fila_primer_findero+indice*espacio_vertical,columna_nombre_finderos+4,periodos[indice])
        worksheet.write(fila_primer_findero+indice*espacio_vertical,columna_nombre_finderos+6,incios[indice])
        worksheet.write(fila_primer_findero+indice*espacio_vertical,columna_nombre_finderos+8,finales[indice])        
        
        for indice_,columna in enumerate(datos[findero]):
            worksheet.write(fila_primer_findero+2+indice*espacio_vertical,0+indice_*espacio_horizontal,'Puerto '+str(indice_+1), bold)
            worksheet.merge_range(fila_primer_findero+2+indice*espacio_vertical,0+indice_*espacio_horizontal,
                                  fila_primer_findero+2+indice*espacio_vertical,0+indice_*espacio_horizontal+encabezado,
                                  'Puerto '+str(indice_+1), encabezados)
            worksheet.write(fila_primer_findero+3+indice*espacio_vertical,0+indice_*espacio_horizontal,'kWh',centrado)
            worksheet.write(fila_primer_findero+3+indice*espacio_vertical,1+indice_*espacio_horizontal,'Señal',centrado)
            worksheet.write(fila_primer_findero+3+indice*espacio_vertical,2+indice_*espacio_horizontal,'%',centrado)
            worksheet.write(fila_primer_findero+4+indice*espacio_vertical,0+indice_*espacio_horizontal, columna, formato_kWh)
            worksheet.write(fila_primer_findero+4+indice*espacio_vertical,1+indice_*espacio_horizontal,'-', centrado)
            
            celda_consumo = xl_rowcol_to_cell(fila_primer_findero+4+indice*espacio_vertical,0+indice_*espacio_horizontal)
            celda_total = xl_rowcol_to_cell(2+fila_primer_findero+len(list(datos.keys()))*espacio_vertical,1+columna_nombre_finderos)
#            pdb.set_trace()
            worksheet.write_formula(fila_primer_findero+4+indice*espacio_vertical,2+indice_*espacio_horizontal,'='+celda_consumo+'/$'+celda_total[0]+'$'+celda_total[1:],porcentaje)
            
            worksheet.conditional_format(fila_primer_findero+4+indice*espacio_vertical,2+indice_*espacio_horizontal,
                                         fila_primer_findero+4+indice*espacio_vertical,2+indice_*espacio_horizontal,
                                         {'type':     'cell',
                                        'criteria': 'between',
                                        'minimum':    .04,
                                        'maximum':    .09,
                                        'format':   condicional_amarillo})
            
            worksheet.conditional_format(fila_primer_findero+4+indice*espacio_vertical,2+indice_*espacio_horizontal,
                                         fila_primer_findero+4+indice*espacio_vertical,2+indice_*espacio_horizontal,
                                         {'type':     'cell',
                                        'criteria': '>=',
                                        'value':    .09,
                                        'format':   condicional_rojo})
    
            worksheet.write_formula(fila_primer_findero+6+indice*espacio_vertical,2+indice_*espacio_horizontal,
                                    '='+xl_rowcol_to_cell(fila_primer_findero+6+indice*espacio_vertical,0+indice_*espacio_horizontal)+
                                    '/$'+celda_total[0]+'$'+celda_total[1:], porcentaje)
            
            celdas_cargas.append(xl_rowcol_to_cell(fila_primer_findero+4+indice*espacio_vertical,0+indice_*espacio_horizontal)) 
            
            
            if indice_+1 == len(datos[findero]):
                
                suma_cargas = sumatoria(celdas_cargas)
                
                idx_ = indice_+1
                worksheet.write(fila_primer_findero+2+indice*espacio_vertical,0+idx_*espacio_horizontal,'Total findero', bold)
                worksheet.write(fila_primer_findero+3+indice*espacio_vertical,0+idx_*espacio_horizontal,'kWh', centrado)
                worksheet.write(fila_primer_findero+3+indice*espacio_vertical,1+idx_*espacio_horizontal,'%', centrado)
                worksheet.write_formula(fila_primer_findero+4+indice*espacio_vertical,0+idx_*espacio_horizontal,suma_cargas)
                worksheet.write_formula(fila_primer_findero+4+indice*espacio_vertical,1+idx_*espacio_horizontal,
                                        '='+xl_rowcol_to_cell(fila_primer_findero+4+indice*espacio_vertical,0+idx_*espacio_horizontal)+'/'+celda_total,porcentaje)

                celdas_finderos.append(xl_rowcol_to_cell(fila_primer_findero+4+indice*espacio_vertical,0+idx_*espacio_horizontal))

                
        if indice+1 == len(list(datos.keys())):
            idx = indice+1
            
            suma_finderos = sumatoria(celdas_finderos) 
            
            worksheet.merge_range(fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos,
                              fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+encabezado-1,'Periodo',encabezados)
            worksheet.merge_range(fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+3,
                              fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+encabezado-1+3,'Bimestre',encabezados)
            
            worksheet.write(1+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos,'Horas:')
            worksheet.write_formula(1+fila_primer_findero+idx*espacio_vertical,1+columna_nombre_finderos,'=I2')
            
            worksheet.write(-1+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos,'DAC:')
            worksheet.write(-1+fila_primer_findero+idx*espacio_vertical,1+columna_nombre_finderos,precio,money)
            celda_precio = xl_rowcol_to_cell(-1+fila_primer_findero+idx*espacio_vertical,1+columna_nombre_finderos)
            
            worksheet.write(2+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos,'Consumo:')
            worksheet.write_formula(2+fila_primer_findero+idx*espacio_vertical,1+columna_nombre_finderos,suma_finderos,formato_kWh)
            celda_consumo = xl_rowcol_to_cell(2+fila_primer_findero+idx*espacio_vertical,1+columna_nombre_finderos)
            
            worksheet.write(3+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos,'Costo:')
            worksheet.write(3+fila_primer_findero+idx*espacio_vertical,1+columna_nombre_finderos,producto_varias(celda_precio,celda_consumo),money_miles)
            celda_costo = xl_rowcol_to_cell(3+fila_primer_findero+idx*espacio_vertical,1+columna_nombre_finderos)

            worksheet.write(4+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos,'Periodos al bimestre:')
            worksheet.write(4+fila_primer_findero+idx*espacio_vertical,1+columna_nombre_finderos,'=(60*24)/I2',formato_kWh)
            celda_periodos = xl_rowcol_to_cell(4+fila_primer_findero+idx*espacio_vertical,1+columna_nombre_finderos)

            
            worksheet.write(1+fila_primer_findero+idx*espacio_vertical,3+columna_nombre_finderos,'Consumo:')
            worksheet.write(1+fila_primer_findero+idx*espacio_vertical,4+columna_nombre_finderos,'='+celda_periodos+'*'+celda_consumo,formato_kWh)
            
            worksheet.write(2+fila_primer_findero+idx*espacio_vertical,3+columna_nombre_finderos,'Costo:')
            worksheet.write(2+fila_primer_findero+idx*espacio_vertical,4+columna_nombre_finderos,'='+celda_periodos+'*'+celda_costo,money_miles)
             
            global celda_consumo_bim
            celda_consumo_bim = xl_rowcol_to_cell(1+fila_primer_findero+idx*espacio_vertical,4+columna_nombre_finderos)
    
    
def desciframiento(datos,horas,precio,inicio,final,cliente,mes,nombre,workbook):
    
    titulos = ['Consumo (kWh)', 'Gasto', 'Ubicación', 'Equipo', 'Proporción (%)', 'Consumo (kWh)', 'Gasto', 'Notas']
    anchos = [12,12,15,19,15,15,10,80,10]
    celda_titulo = 'A1'
    celdas = 16
    filas = len(datos.keys())


    bold_1 = workbook.add_format({'bold': True,'font_size':12})
    bold_2 = workbook.add_format({'bold': True,'font_size':15})
    bold_3 = workbook.add_format({'bold': True,'bg_color':'#E1E4EB','align':'center','border':1})
    columna_gris = workbook.add_format({'bold': True,'bg_color':'#F2F2F2','border':1})
    columna_blanca_1 = workbook.add_format({'align':'center','num_format': '0.0 %','border':1})                                        
    columna_blanca_2 = workbook.add_format({'align':'center','num_format': '#','border':1})
    columna_blanca_3 = workbook.add_format({'align':'center','num_format': '$  #,###     ','border':1})
    dinero_1 = workbook.add_format({'align':'center','num_format': '$   #,###.##   ','border':1,'align':'center'}) 

                                     
    worksheet = workbook.add_worksheet('Desciframiento') 
    
    worksheet.write(celda_titulo,'Desciframiento del consumo de ' + nombre + ' (cifras bimestrales)', bold_2)
    
    for idx,titulo in enumerate(titulos):
        skip = 0
        reset = 0
        if idx>1:
            skip = 4
            reset = 2
        worksheet.set_column(idx, idx, anchos[idx])
        worksheet.write(xl_rowcol_to_cell(3+skip,idx-reset+2),titulo, bold_3)
        
    
    for i in range(0,celdas):
        worksheet.write_blank(xl_rowcol_to_cell(8+i,2),'', columna_gris)
        worksheet.write_blank(xl_rowcol_to_cell(8+i,3),'', columna_gris)
        worksheet.write_blank(xl_rowcol_to_cell(8+i,4),'', columna_blanca_1)
        worksheet.write_formula(xl_rowcol_to_cell(8+i,5),'=' + xl_rowcol_to_cell(8+i,4) + '*$C$5', columna_blanca_2)
        worksheet.write_formula(xl_rowcol_to_cell(8+i,6),'=' + xl_rowcol_to_cell(8+i,4) + '*$D$5', columna_blanca_3)
        worksheet.write_blank(xl_rowcol_to_cell(8+i,7),'', columna_blanca_1)
    
    worksheet.write(xl_rowcol_to_cell(8+celdas,3),'Total',bold_3)
    worksheet.write(xl_rowcol_to_cell(8+celdas+1,3),'Total de Fugas',bold_3)
    
    worksheet.write_formula(xl_rowcol_to_cell(8+celdas,4),'=SUM('+xl_rowcol_to_cell(8,4)+
                                                                ':'+xl_rowcol_to_cell(8+celdas-1,4)+
                                                                                ')',columna_blanca_1)
    
    worksheet.write_formula(xl_rowcol_to_cell(8+celdas+1,4),'=SUM('+xl_rowcol_to_cell(8+6,4)+
                                                                ':'+xl_rowcol_to_cell(8+10,4)+
                                                                                ')',columna_blanca_1)
    
    worksheet.write('F4','Tarifa DAC:',bold_3)
    worksheet.write_number('G4', 5.626, dinero_1)    
    worksheet.write_formula('D5','=C5*G4',columna_blanca_3)
    worksheet.write_formula('C5','=Detalles!'+celda_consumo_bim,columna_blanca_2)
    
    worksheet.write('D9','Refrigerador', columna_gris)
    worksheet.write('D11','Bomba de agua', columna_gris)
    worksheet.write('D12','Centro de lavado', columna_gris)
    worksheet.write('D13',"Tv's", columna_gris)
    worksheet.write('D21','Luces', columna_gris)
    worksheet.write('D22','Cómputo y cargadores', columna_gris)
    worksheet.write('D23','Sin Identificar', columna_gris)
    
    #=MID(LEFT(FORMULATEXT(I11),FIND("*",FORMULATEXT(I11))-1),FIND("=",FORMULATEXT(I11))+1,LEN(FORMULATEXT(I11)))

    for i in range(5):
#        pdb.set_trace()
        row_fuga = 'ROW(INDIRECT(MID(FORMULATEXT('+xl_rowcol_to_cell(14+i,4)+'),FIND("!",FORMULATEXT('+xl_rowcol_to_cell(14+i,4)+'))+1,256)))'
        column_fuga = 'COLUMN(INDIRECT(MID(FORMULATEXT('+xl_rowcol_to_cell(14+i,4)+'),FIND("!",FORMULATEXT('+xl_rowcol_to_cell(14+i,4)+'))+1,256)))-2'
        argumento = 'FORMULATEXT(INDIRECT(ADDRESS('+row_fuga+','+column_fuga+',,,"Detalles")))'
        worksheet.write_formula(xl_rowcol_to_cell(14+i,3),'IFERROR(CONCATENATE("Fuga ",'+'MID(LEFT('+argumento+',FIND("*",'+argumento+')-1),FIND("=",'+argumento+')+1,LEN('+argumento+'))),"Fuga")', columna_gris)
        
    worksheet.write_formula('E11',crear_vlookup(filas,'Bomba'), columna_blanca_1)  
    worksheet.write_formula('E12',crear_vlookup(filas,'Lavado'), columna_blanca_1)
    worksheet.write_formula('E13',crear_vlookup(filas,'TV'), columna_blanca_1)
    
    
    worksheet.write_formula('E21',crear_vlookup(filas,'Luces'), columna_blanca_1)
    worksheet.write_formula('E22',crear_vlookup(filas,'Computo'), columna_blanca_1)
    worksheet.write_formula('E23',crear_vlookup(filas,'Sin ID'), columna_blanca_1)
    
    
def ahorro(datos,horas,precio,inicio,final,cliente,mes,nombre,workbook):  
    
    celda_titulo = 'A1'
    
    bold_1 = workbook.add_format({'bold': True,'font_size':15})
    bold_2 = workbook.add_format({'bold': True,'align':'center'})
    bold_3 = workbook.add_format({'bold': True,'top':1})
    blanco = workbook.add_format({'font_color':'#ffffff'})
    dinero_1 = workbook.add_format({'num_format': '$      #,###      '})
    dinero_2 = workbook.add_format({'num_format': '$      #,###      ','top':1})
    kWh = workbook.add_format({'num_format': '#','align':'center'})
    kWh_2 = workbook.add_format({'num_format': '#','top':1})
                               
    worksheet = workbook.add_worksheet('Ahorro') 
    worksheet.set_column(0,0,45)
    worksheet.set_column(2, 2, 15)
    worksheet.set_column(3, 3, 15)
    
    worksheet.write(celda_titulo,'Potencial de ahorro de ' + nombre,bold_1)
    worksheet.write('A4','Acción', bold_2)
    worksheet.write('A5','Eliminar fugas')
    worksheet.write('C4','Ahorro (DAC)',bold_2)
    worksheet.write('D4','Ahorro en kWh',bold_2)
    worksheet.write('B9','Total: ', bold_3)
    worksheet.write_formula('C9','SUM(C5:C8)',dinero_2) 
    worksheet.write_formula('D9','SUM(D5:D8)',kWh_2)
    worksheet.write_formula('D5','IF(SUM(Desciframiento!F15:F19)-40<0,0,SUM(Desciframiento!F15:F19)-40)',kWh)
    worksheet.write_formula('C5','D5*Desciframiento!G4',dinero_1)
    
    worksheet.write('A13','Cambio de tarifa',bold_2)
    worksheet.write_formula('I2','IF(D15<500,1,0)',blanco)
    
    worksheet.write_formula('A14','IF(I2=1,"Sí es posible bajar de tarifa","No es posible bajar de tarifa")')
    
    worksheet.write('C13','DAC:',bold_2)
    
    worksheet.write('C14','Nuevo recibo',bold_2)
    worksheet.write_formula('C15','Desciframiento!D5-C9',dinero_1)
    
    worksheet.write('D14','Nuevo consumo',bold_2)
    worksheet.write_formula('D15','Desciframiento!C5-D9',kWh)
    
    worksheet.write_formula('C17','IF(I2=1,"Subsidiada:","")',bold_2)
    
    worksheet.write_formula('C18','IF(I2=1,"Nuevo recibo","")',bold_2)
    worksheet.write_formula('D18','IF(I2=1,"Ahorro total","")',bold_2)
    
    worksheet.write_formula('C19','IF(I2=1,IF(D15>=280,150*.808+130*.976+(D15-280)*2.8,IF(D15>=150,150*.808+(D15-150)*.976,D15*.808)),"")',dinero_1)
    worksheet.write_formula('D19','IF(I2=1,Desciframiento!D5-C19,"")',dinero_1)
    
    
def fugas(datos,horas,precio,inicio,final,cliente,mes,nombre,workbook):
    
    worksheet = workbook.add_worksheet('Fugas')
    worksheet.set_column(1,1,2)
    worksheet.set_column(4,4,2)
        
    bold_1 = workbook.add_format({'bold': True,'font_size':15})
    bold_2 = workbook.add_format({'bold': True,'align':'center'})
    center = workbook.add_format({'num_format': '0.0','align':'center'})
    center_2 = workbook.add_format({'align':'center'})
    
    
    worksheet.write('A1','Resumen de fugas de ' + nombre,bold_1)
    worksheet.write('A4','Circuito',bold_2)
    worksheet.write('C4','Fuga (W)',bold_2)
    worksheet.write('D4','A',bold_2)
    worksheet.write('F4','Notas',bold_2)
    
    for i in range(0,5):
        worksheet.write_formula(xl_rowcol_to_cell(5+2*i,3),'=IFERROR('+xl_rowcol_to_cell(5+2*i,2)+'/127,0)',center)
        worksheet.write_formula(xl_rowcol_to_cell(5+2*i,0),'=IFERROR(MID(Desciframiento!'+xl_rowcol_to_cell(14+i,2)
                                                            +',FIND(" ",Desciframiento!'+xl_rowcol_to_cell(14+i,2)+')+1,256)," ")',center_2)
        worksheet.write(xl_rowcol_to_cell(5+2*i,2),'=IFERROR(MID(Desciframiento!'+xl_rowcol_to_cell(14+i,3)
                                                            +',FIND(" ",Desciframiento!'+xl_rowcol_to_cell(14+i,3)+')+1,256)," ")',center_2)
        
    
    
def excel(datos,horas,precio,inicio,final,cliente,mes,inicios,finales,periodos):
    
    nombre = cliente[3:]
    nombre_ = cliente[3:].replace(' ','_')
    direccion = 'D:/01 Findero/'+mes+'/'+cliente+'/Resultados/ResultadosGenerales_'+nombre_+'.xlsx'
    workbook = xlsxwriter.Workbook(direccion)
    
    general(datos,horas,precio,inicio,final,cliente,mes,nombre,workbook)   
    detalles(datos,horas,precio,inicio,final,cliente,mes,nombre,workbook,inicios,finales,periodos)
    desciframiento(datos,horas,precio,inicio,final,cliente,mes,nombre,workbook)
    ahorro(datos,horas,precio,inicio,final,cliente,mes,nombre,workbook)
    fugas(datos,horas,precio,inicio,final,cliente,mes,nombre,workbook)
            
    workbook.close()

if __name__ == '__main__':
    precio = 5.626
    horas = 166
    inicio = '2019-05-03'
    final = '2019-05-11'
    finderos = ['F03','F05','F21']
    datos = {finderos[0]: [i for i in range(0,12)],
             finderos[1]: [3*i for i in range (5,17)],
             finderos[2]: [.15*i for i in range (4,16)] }
    excel(datos,precio,horas,inicio,final)
    
