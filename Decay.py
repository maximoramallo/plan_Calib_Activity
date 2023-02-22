#import xlsxwriter
#from openpyxl import load_workbook
#from openpyxl.drawing.image import Image
#import matplotlib.pyplot as plt
import math as mt
import datetime as date
from datetime import datetime as dt


def calcula_Tiempo_Decay():
    horaCalib = input("Ingrese fecha y hora de calibración aaaa/mm/dd hh:mm =  ")
    horaEntrega = input("Ingrese hora de entrega hh:mm =  ")
    anoCal = int(horaCalib[0:4]); mesCal = int(horaCalib[5:7])
    diaCal = int(horaCalib[8:10]); hraCal = int(horaCalib[11:13]); minCal = int(horaCalib[14:16])
    hraEnt = int(horaEntrega[0:2]); minEnt = int(horaEntrega[3:5])
    hraCalib = date.time(hraCal,minCal)
    horaEnt = date.time(hraEnt,minEnt)
    horaCal = dt(anoCal,mesCal,diaCal,hraCal,minCal)
    horaEntrega = dt(anoCal,mesCal,diaCal,hraEnt,minEnt)
    dif = horaEntrega - horaCal
    tiempoDecay = dif.total_seconds()/(60*60)
    return tiempoDecay, hraCalib, horaEnt

def activity_decay(actividad, semiperiodo, tiempo):     #Función para calcular actividad en funcion de t.
    factorDecay = mt.exp((-mt.log(2)*tiempo)/semiperiodo)
    actividadMadre = int(actividad/factorDecay)
    return actividadMadre
