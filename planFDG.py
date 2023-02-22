#Faltan subir archivos de las funciones particulares que se llaman en el código. Próximamente...

from ConcDiluidasTrun import truncate, dilucion
from Decay import calcula_Tiempo_Decay, activity_decay
from Vol_numViales import calc_num_viales, calc_volActHijas, busca_VolumenMenor
import xlsxwriter
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import matplotlib.pyplot as plt
import math as mt
import datetime as date
from datetime import datetime as dt
from datetime import timedelta

fAct = dt.now()
f_prod = fAct+timedelta(days=1)
hora = "{}:{}:{}".format(fAct.hour, fAct.minute, fAct.second)
fecha = "{}/{}/{}".format(fAct.day, fAct.month, fAct.year)
fecha_prod = "{}/{}/{}".format(f_prod.day, f_prod.month, f_prod.year)

archivoExcel = xlsxwriter.Workbook('C:/~/test.xlsx')
hoja_3 = archivoExcel.add_worksheet('Encabezado')

imagen = Image('logo.png')

#FORMATOS
#Tamaños
hoja_3.set_column('A:A',14)
hoja_3.set_column('B:B',13)
hoja_3.set_column('C:C',12)
hoja_3.set_column('D:D',16)
hoja_3.set_column('E:E',12)
hoja_3.set_column('F:F',20)
hoja_3.set_column('G:G',22)
hoja_3.set_column('H:H',13)
hoja_3.set_column('I:I',20)
hoja_3.set_column('J:J',20)
hoja_3.set_row(0,25)
hoja_3.set_row(3,25)

cabecera = archivoExcel.add_format({'bold':1,'align':'center','valign':'vcenter','fg_color':'#C0C0C0','size':'11','font_name':'Century Gothic'})
negrita = archivoExcel.add_format({'bold':True,'align':'center','valign':'vcenter'})
variable = archivoExcel.add_format({'align':'center','border':2,'border_color':'#000000','valign':'vcenter','size':'11','font_name':'Century Gothic'})
negrita1 = archivoExcel.add_format({'bold':True,'align':'center','border':2,'border_color':'#000000','valign':'vcenter','size':'11','font_name':'Century Gothic'})
negrita2 = archivoExcel.add_format({'bold':True,'align':'center','border':2,'border_color':'#000000','valign':'vcenter','size':'10','font_name':'Century Gothic'})
negrita3 = archivoExcel.add_format({'bold':True,'align':'center','fg_color':'#C0C0C0','border':2,'border_color':'#000000','valign':'vcenter','size':'11','font_name':'Century Gothic'})
negrita4 = archivoExcel.add_format({'bold':True,'align':'center','fg_color':'#C0C0C0','border':2,'border_color':'#000000','valign':'vcenter','size':'11','font_name':'Century Gothic','text_wrap': True})

#Datos de Planilla
cabecera.set_border(2), cabecera.set_border_color('#000000')

hoja_3.merge_range('C1:I1','REGISTRO',negrita1)
hoja_3.merge_range('C2:I2','PLANEAMIENTO: ACTIVIDADES DISPENSADAS PARA CLIENTES',negrita1)
hoja_3.merge_range('A1:B2',' ',negrita1)
hoja_3.write('J1','R1-Planea-01',negrita1)
hoja_3.merge_range('A3:C3','Fecha de vigencia:',negrita1)
hoja_3.merge_range('D3:E3','05/12/2022',negrita1)
hoja_3.merge_range('F3:H3','Fecha de próxima versión:',negrita1)
hoja_3.merge_range('I3:J3','05/12/2024',negrita1)
negrita3.set_border(2), negrita3.set_border_color('#000000')
hoja_3.write('J2','Página 1 de 1',negrita2)

hoja_3.write(3,0,'Nº de Pedido',negrita3)
hoja_3.write(3,1,'Cliente',negrita3)
hoja_3.write(3,2,'Fecha',negrita3)
hoja_3.write(3,3,'Act. Req., mCi',negrita3)
hoja_3.write(3,4,'Hra. Req.',negrita3)
hoja_3.write(3,5,'Act. Calib., mCi',negrita3)
hoja_3.write(3,6,'Act. Calib. corr., mCi',negrita3)
hoja_3.write(3,7,'Hra. Calib.',negrita3)
hoja_3.write(3,8,'Vol. Dispens., mL',negrita3)
hoja_3.write(3,9,'Nº lote',negrita3)

#LOGO OULTON RADIOFARMACOS
hoja_3.insert_image('A1:B2','logo.png',{'x_offset':1,'y_offset':0,'x_scale':0.27,'y_scale':0.21,'positioning':2})

#Solicita número de pedidos de clientes
numPedidos = input("Ingrese numero de pedido: ")
numPedidos = int(numPedidos)
actAcumulada = 0; cuentaOrden = 0
listActMadre = []; listActMadrecorr = []
listCuentaOrden = []; listCliente = []
x=0; k=0

while x < numPedidos:
    try:
        print("*************************")
        print("**** DATOS PEDIDO {} ****".format(x+1))
        print("*************************")
        cliente = input("Nombre del Cliente: ")
        actividadRequerida = float(input("Ingrese actividad requerida (mCi): "))
        actividadRequerida = actividadRequerida*1.10  #Agrego a la Actividad Requerida un 10% de cobertura.
        if x == 0:
            numLote = input("Ingrese número de lote: ")
            if len(numLote) != 15 or numLote[13] != "-" or numLote[4:7] != "FDG":
                raise ValueError
        tiempoDecay, hraCalib, hraEnt = calcula_Tiempo_Decay()
        actividadMadre = activity_decay(actividadRequerida, 1.82, tiempoDecay)
        actividadMadrecorr = actividadMadre*1.20   
        listActMadre.append(actividadMadre); listActMadrecorr.append(actividadMadrecorr)
        listCliente.append(cliente)
        x+=1; falla = 0
    except ValueError:
        print("*************************")
        print("Debes Ingresar los datos en los siguientes formatos: ")
        print("Actividad requerida ---> nº entero en mCi")
        print("Nº lote de producción ---> según formato indicado en POE RF-003")
        print("Fecha y Hora de Calibración ---> aaaa/mm/dd hh:mm")
        print("Hora de Entrega ---> hh:mm")
        falla = 1
    except TypeError:
        print("*************************")
        print("Debes Ingresar la hora de entrega en el siguiente formato hh:mm")
        falla = 1
    except NameError:
        print("*************************")
        print("Alguna de las variables quedó mal definida. Vuelva a introducir los datos en formato pedido.")
        falla = 1
    if falla == 0:
        cuentaOrden = int(cuentaOrden)   #Transforma en entero cuentaOrden para poder incrementar en una unidad.
        cuentaOrden+=1
        cuentaOrden = str(cuentaOrden)   #Transforma en cadena cuentaOrden para que aparezcan cuadro dígitos en la columna de Excel.
        cuentaOrden = "000" + cuentaOrden
        hoja_3.write(4+k,0,cuentaOrden,variable); hoja_3.write(4+k,1,cliente,variable); hoja_3.write(4+k,2,fecha,variable)
        hoja_3.write_number(4+k,3,actividadRequerida,variable); hoja_3.write(4+k,4,hraEnt.isoformat("minutes"),variable)
        hoja_3.write_number(4+k,5,actividadMadre,variable) #Actividades calculadas por vial en columna.
        hoja_3.write_number(4+k,6,actividadMadrecorr,variable)
        hoja_3.write(4+k,7,hraCalib.isoformat("minutes"),variable); hoja_3.write(4+k,9,numLote,variable)
        listCuentaOrden.append(cuentaOrden)
        k+=1   #Comienza a contar k, 4+k es el ugar donde comienza el registro en las filas.
    else:
        pass
      
lista_volVial, lista_f_Vm, volTotal = calc_volActHijas(listActMadre,0)
print(len(listActMadre), len(listActMadrecorr))

#Volumen de solucion madre
lista_concActHija = []
lista_concActHijacorr = []
for k in range(len(listActMadre)):
    concActHija = listActMadre[k]/lista_volVial[k]
    concActHijacorr = listActMadrecorr[k]/lista_volVial[k]
    lista_concActHija.append(concActHija)
    lista_concActHijacorr.append(concActHijacorr)

c = 0
for z in range(len(lista_concActHija)):      # Se calcula la máxima concentración de actividad
    if lista_concActHija[c] < lista_concActHija[z]:  # Para tomar como referencia la máxima actividad de la madre en el menor volumen (dará la concentración máxima necesaria)
        c = z

concActMadre = lista_concActHija[c]  #Conc. Act., mCi/mL
concActMadrecorr = lista_concActHijacorr[c]   #Conc. Act. corr., mCi/mL


#Volumen solución madre (concentrada)
lista_Vm_corr = []
for j in range(len(listActMadre)):
    fraccionVolMadre = float(listActMadre[j])/float(concActMadre)
    lista_Vm_corr.append(fraccionVolMadre)
for k in range(numPedidos):
    hoja_3.write_number(4+k,8,truncate(lista_Vm_corr[k],3),variable)


#Buscar el menor volumen fuera del rango (3 - 5 mL).
listaVolRemanente = []
listaOrdenesPreparados = []

try:
    for j in range(len(lista_Vm_corr)):
        if lista_Vm_corr[j] >= 3:
            listaVolRemanente.append(lista_Vm_corr[j])  #Saca los volumenes que se pueden dispensar a la concentración calculada.
            listaOrdenesPreparados.append(listCuentaOrden[j])  #Saca los número de pedidos que se pueden preparar a la concentración calculada.
    vol_Remanente = sum(listaVolRemanente)  #Volumen total dispensado en la primera predilución. 
    min_Vol_Conc, indice = busca_VolumenMenor(lista_Vm_corr,listCliente)  #Selecciona el menor volumen y el índice que le corresponde (de acuerdo al orden de ingreso de datos).
    #Recalcula la concentración de la solución madre para preparar el vial que tenía el menor volumen.
    #El 3 es debido a la restricción de (3 - 5) mL.
    conc_Madre_Diluida = listActMadrecorr[indice]/3  
    print(concActMadrecorr, listActMadre[indice])
    
    #Factor de dilución.                                             
    factor_Diluc = concActMadrecorr/conc_Madre_Diluida
    #factor_Diluc = concActMadre/conc_Madre_Diluida
    #Se suma vol_extra por cobertura. Se resta vol remanente por haberse preparado los viales sin diluir.
    volumen_Final = factor_Diluc*truncate(sum(lista_Vm_corr)-vol_Remanente-lista_Vm_corr[0],3)
    vol_SolFisio_Agregar = volumen_Final - truncate(sum(lista_Vm_corr)-vol_Remanente-lista_Vm_corr[0],3)

    #Predilucion
    hoja_3.merge_range('A'+str(13+k)+':'+'C'+str(13+k),'Factor de Dilución',negrita3)
    hoja_3.write_number(k+12,3,factor_Diluc,negrita1)

    #Viales afectados por dilución.
    hoja_3.merge_range('F'+str(13+k)+':'+'J'+str(13+k),'Viales afectados por Dilución',negrita1)
    hoja_3.write(k+13,5,'Nº de Pedido',negrita3)
    hoja_3.write(k+13,6,'Cliente',negrita3)
    hoja_3.merge_range('H'+str(14+k)+':'+'I'+str(14+k),'Act. Calib. corr., mCi',negrita3)
    hoja_3.write(k+13,9,'Vol. Madre Dil., mL',negrita3)
    lista_NumPed_dil = []
    lista_Cliente_dil = []
    lista_ActMadcorr = []
    lista_NumPed_dil.append(listCuentaOrden[indice])
    lista_Cliente_dil.append(listCliente[indice])
    lista_ActMadcorr.append(listActMadrecorr[indice])

    for j in range(len(listCuentaOrden)):
        if lista_Vm_corr[j] > lista_Vm_corr[indice] and lista_Vm_corr[j] < 3:
            lista_NumPed_dil.append(listCuentaOrden[j])
            lista_Cliente_dil.append(listCliente[j])
            lista_ActMadcorr.append(listActMadrecorr[j])

    for j in range(len(lista_NumPed_dil)):
        hoja_3.write(k+14+j,5,lista_NumPed_dil[j],variable)
        hoja_3.write(k+14+j,6,lista_Cliente_dil[j],variable)
        hoja_3.merge_range('H'+str(15+k+j)+':'+'I'+str(15+k+j),lista_ActMadcorr[j],variable)
        hoja_3.write_number(k+14+j,9,lista_ActMadcorr[j]/conc_Madre_Diluida,variable)

    #Registro de viales preparados.
    hoja_3.merge_range('E'+str(9+k)+':'+'G'+str(11+k),'Se preparan los pedidos {}\nRealizar predilución de los pedidos {}.'.format(listaOrdenesPreparados,lista_NumPed_dil),negrita4)

    hoja_3.merge_range('A'+str(14+k)+':'+'C'+str(14+k),'Volumen de Fisiológica',negrita3)
    hoja_3.write_number(k+13,3,vol_SolFisio_Agregar,negrita1)
    hoja_3.merge_range('A'+str(15+k)+':'+'C'+str(15+k),'Concentración Madre Diluida',negrita3)
    hoja_3.write_number(k+14,3,conc_Madre_Diluida,negrita1)

except:
    print("¡No es necesario diluir!")
    #Registro de viales preparados.
    hoja_3.merge_range('E'+str(9+k)+':'+'G'+str(11+k),'Se preparan los pedidos {}\n¡No es necesario realizar prediluciones!.'.format(listaOrdenesPreparados),negrita4)

    
#Actividad total
hoja_3.merge_range('A'+str(8+k)+':'+'C'+str(8+k),'Conc. Act. de Síntesis, mCi/mL',negrita3)
conc_Fin_Sintesis = 2*concActMadrecorr
hoja_3.write(5+k+2,3,truncate(conc_Fin_Sintesis,3),negrita1)
primera_Predil = dilucion(conc_Fin_Sintesis,concActMadrecorr,10)
hoja_3.merge_range('A'+str(9+k)+':'+'C'+str(9+k),'Conc. Act., mCi/mL',negrita3)
hoja_3.write(5+k+3,3,truncate(concActMadre,3),negrita1) #Actividad total
hoja_3.merge_range('A'+str(10+k)+':'+'C'+str(10+k),'Conc. Act. corr., mCi/mL',negrita3)
hoja_3.write(5+k+4,3,truncate(concActMadrecorr,3),negrita1) #Actividad total (teniendo el cuenta un 10% extra en cada vial)

#Volumen total
hoja_3.write(k+6,4,'Totales',negrita3)
hoja_3.write_number(k+6,5,(sum(listActMadre)),negrita1)
hoja_3.write_number(k+6,6,(sum(listActMadrecorr)),negrita1)
hoja_3.write_number(k+6,8,(truncate(sum(lista_Vm_corr),3)),negrita1)
hoja_3.merge_range('A'+str(11+k)+':'+'C'+str(11+k),'Volumen total (sugerido), mL',negrita3)
vol_disp_total = truncate(sum(lista_Vm_corr))

#Volumen total sugerido.
hoja_3.write_number(k+10,3,primera_Predil+10,negrita1)

archivoExcel.close()
