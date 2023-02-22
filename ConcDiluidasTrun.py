def truncate(n, decimals = 0):
    multiplier = 10 ** decimals
    return int(n * multiplier) / multiplier


def dilucion_Act(minActividad,conMaxIni,volRemanente):
    concDil = minActividad/3
    volumenDil = (conMaxIni/concDil)*volRemanente
    volumenAdd = volumenDil - volRemanente
    return volumenAdd

def dilucion(concMad,concHija,vol):
    fd = concMad/concHija
    volumenFinal = fd*vol
    volumenFisiologica = volumenFinal - vol
    return volumenFisiologica
