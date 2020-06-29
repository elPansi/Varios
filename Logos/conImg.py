# -*- coding: utf-8 -*-
import time, sys, logging, os
from datetime import date, timedelta, datetime
from pathlib import Path

nomFich = "start_16_16"

longByte = 8
fichEscritura = open(nomFich + "_OK.txt", "w")
lin = 1
lineaTmp =""
f = open(nomFich + ".txt", "r")
for l in f:
    nuevaLin=""
    posIni=0
    posFin = posIni + longByte
    while(posFin <= len(l)):
        nuevaLin = nuevaLin + "B" + l[posIni:posFin] + ", "
        print "Nueva Linea: " + str(nuevaLin) + "\r\n"
        posIni = posFin #+ 1
        posFin = posIni + longByte
    fichEscritura.write(nuevaLin + "\n")
    #fichEscritura.write(nuevaLin)
    print "Linea: " + str(lin) + " Posicion Final: " + str(posFin)
    print "Linea Antigua: " + str(l) + "\n"
    print "Linea Nueva:   " + str(nuevaLin) + "\n"
    lin=lin+1
	
f.close()
fichEscritura.close()