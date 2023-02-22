
import xlsxwriter
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import matplotlib.pyplot as plt
import math as mt
import datetime as date
from datetime import datetime as dt


def calc_num_viales(act_total, max_Act, min_Act, cliente):
    if act_total < max_Act:
        print("Necesitas solo un vial")
    else:
        listAct=[max_Act]
        while act_total > min_Act:
            act_total =  act_total - max_Act
            if act_total > min_Act and act_total > max_Act:
                listAct.append(max_Act)
            elif act_total > min_Act and act_total < max_Act:
                listAct.append(act_total)
        print("Necesitas", len(listAct), "viales para el cliente", cliente)
        for k in range(len(listAct)):
            pacientes = (listAct[k]/max_Act)*4
            pacientes = round(pacientes)
            print("Vial", k+1, listAct[k],"mCi","Rendimiento:", pacientes,"pacientes")


def calc_volActHijas(listaActMadre,volSolSalina):
    lista_volVial = []
    lista_f_Vm = []
    k = 0
    Vd = volSolSalina   # Este volumen sería el que se utilizaría de sol. fisiológica para llegar al vol vial.
    volVial = float(input("Ingrese el volumen final en el vial (3-5 mL):"))
    for actHija in listaActMadre:
        lista_volVial.append(volVial)
        f_Vm = volVial-Vd
        lista_f_Vm.append(f_Vm)
        k=k+1
    volTotal = 0
    for j in range(k):
        volTotal = lista_f_Vm[j] + volTotal
    return lista_volVial, lista_f_Vm, volTotal

def busca_VolumenMenor(listaVolumenes, listaClientes):
    minVol = []
    listaCopy = listaVolumenes
    for v in range(len(listaVolumenes)):
        if listaVolumenes[v] < 3 and listaClientes[v] != "CC":
            minVol.append(listaVolumenes[v])
    if len(minVol) != 0:
        min = minVol[0]
        for k in range(len(minVol)):
            if minVol[k] < min:
                min = minVol[k]
        return min, listaVolumenes.index(min)
