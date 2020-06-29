# -*- coding: utf-8 -*-

import time
from datetime import date, timedelta
import logging
import xlwt, xlrd
import sys, ast
import urllib2
import requests
from libAccess import consultaAccess
from yandex.Translater import Translater

reload(sys)
sys.setdefaultencoding("utf-8")

nombreExcel="TraduceEPLAN"
rutaLog = ".\\Log\\"
rutaInformes = ".\\Informes\\"
cadTiempo = time.strftime("%Y")+  time.strftime("%m") + time.strftime("%d") +  "_" + time.strftime("%H")  + time.strftime("%M") +  time.strftime("%S")

f = open(rutaLog + "Log_"  + nombreExcel + "_" + cadTiempo + ".txt", "w")
cadArchivo = rutaInformes + cadTiempo + "_" + nombreExcel + ".xls"

consultaBD = """SELECT id, description1
FROM tblPart
where description1 like '*[??]_[??]*'
order by id;"""

consultaBD = "Select id, note, description1, description2 from tblPart where description1 like '%%??_??%%' and id > 400 order by id;"
#consultaBD = "Select id, note, description1, description2 from tblPart where description1 like '*[??]_[??]*' order by id;"
actualizaCampoDescription1 = "update tblPart set description1 = '{CADENA}' where id={ID};"
actualizaCampoNote = "update tblPart set [note] = '{CADENA}' where id={ID};"
actualizaCampoDescription2 = "update tblPart set description2 = '{CADENA}' where id={ID};"

consultaCT = "SELECT id, technicalcharacteristics FROM tblVariant where id > 630 order by id;"
consultaCT = "SELECT id, technicalcharacteristics FROM tblVariant order by id;"
actualizaCT = "update tblVariant set technicalcharacteristics = '{CADENA}' where id={ID};"


def traduceCarTec(cadenaEntrada):
    apiTranslateEplan = """trnsl.1.1.20190308T091207Z.xxx"""
    tr = Translater()
    tr.set_key(apiTranslateEplan)
    tr.set_from_lang('es')
    tr.set_to_lang('en')
    f.write("===> TraduceCarTec: " + str(cadenaEntrada) + str(type(cadenaEntrada)) )
    tr.set_text(cadenaEntrada)
    cadEng = tr.translate()
    return (cadEng)


def traduceEplan(cadenaEntrada):
    """esta funcion va a devolver la cadena compuesta y traducida a partir de una cadena de entrada"""
    #print "----> " + str(cadenaEntrada)
    cadenaEntrada=str(cadenaEntrada)
    cadEs="es_ES@" + cadenaEntrada + ';'
    #Primero traduccion al ingles
    apiTranslateEplan = """trnsl.1.1.20190308T091207Z.xxx"""
    tr = Translater()
    tr.set_key(apiTranslateEplan)
    tr.set_from_lang('es')
    tr.set_to_lang('en')
    tr.set_text(cadenaEntrada)
    cadEng="en_US@" + tr.translate() + ";"
    # #Frances
    tr.set_to_lang('fr')
    tr.set_text(cadenaEntrada)
    cadFr = "fr_FR@" + tr.translate() + ";"
    cadFr = cadFr.replace("'", "''")
    #Aleman
    tr.set_to_lang('de')
    tr.set_text(cadenaEntrada)
    cadAle = "de_DE@" + tr.translate() + ";"
    # #Polaco
    # tr.set_to_lang('pl')
    # tr.set_text(cadenaEntrada)
    # cadPol = "pl_PL@" + tr.translate() + ";"
    cadenaSalida = cadEs + cadEng + cadFr + cadAle #+ cadPol
    return (cadenaSalida)

#Tabla tblPart
#Obtenemos las cadenas a traducir
salida = consultaAccess(consultaBD)

#f.write("Cabeceras -> " + str(cabeceras) + "\r\n")
f.write("Salida tblPart-> " + str(salida) + "\r\n")

#procesamos la salida
for lin in salida:
    #f.write("Linea: " + str(lin) + "\r\n")
    id = lin[0]
    #TRADUCE Note
    texto = lin[1]
    cadOrig = texto[6:-1].decode('utf-8','ignore')
    print "--> " + str(id) + " - "  #+ cadOrig
    f.write("Cad Original: " + str(id) + " - " + str(cadOrig) + "\r\n")
    cadOriginal = cadOrig.replace("'","''")
    cadTraducida=''
    if cadOrig:
        cadTraducida = traduceEplan(cadOriginal)
        f.write("Cad Traducida: " + str(cadTraducida) + "\r\n")
    sqlActualiza=actualizaCampoNote.replace("{CADENA}",cadTraducida)
    sqlActualiza=sqlActualiza.replace("{ID}",str(id))
    #sqlActFin = sqlActualiza.replace("'","''")
    sqlActFin = sqlActualiza
    f.write("Consulta Actualizacion: " + str(sqlActFin) + "\r\n")
    #intentamos insertarla
    try:
        salAct = consultaAccess(sqlActFin)
    except:
        f.write (":::::::::::::::>>> ERROR NOTE <<<<:::::::::::::::::" + "\r\n")


    #Traducir Description1
    texto = lin[2]
    cadOrig = texto[6:-1].decode('utf-8','ignore')
    print "--> " + str(id) + " - " #+ cadOrig
    f.write("Cad Original: " + str(id) + " - " + str(cadOrig) + "\r\n")
    cadOriginal = cadOrig.replace("'","''")
    cadTraducida = ''
    if cadOrig:
        cadTraducida = traduceEplan(cadOriginal)
        f.write("Cad Traducida: " + str(cadTraducida) + "\r\n")
    sqlActualiza=actualizaCampoDescription1.replace("{CADENA}",cadTraducida)
    sqlActualiza=sqlActualiza.replace("{ID}",str(id))
    #sqlActFin = sqlActualiza.replace("'","''")
    sqlActFin = sqlActualiza
    f.write("Consulta Actualizacion: " + str(sqlActFin) + "\r\n")
    #intentamos insertarla
    try:
        salAct = consultaAccess(sqlActFin)
    except:
        f.write (":::::::::::::::>>> ERROR DESCRIPTION1 <<<<:::::::::::::::::" + "\r\n")

    #Traducir Description2
    texto = lin[3]
    cadOrig = texto[6:-1].decode('utf-8','ignore')
    print "--> " + str(id) + " - " #+ cadOrig
    f.write("Cad Original: " + str(id) + " - " + str(cadOrig) + "\r\n")
    cadOriginal = cadOrig.replace("'","''")
    cadTraducida = ''
    if cadOrig:
        cadTraducida = traduceEplan(cadOriginal)
        f.write("Cad Traducida: " + str(cadTraducida) + "\r\n")
    sqlActualiza=actualizaCampoDescription2.replace("{CADENA}",cadTraducida)
    sqlActualiza=sqlActualiza.replace("{ID}",str(id))
    #sqlActFin = sqlActualiza.replace("'","''")
    sqlActFin = sqlActualiza
    f.write("Consulta Actualizacion: " + str(sqlActFin) + "\r\n")
    #intentamos insertarla
    try:
        salAct = consultaAccess(sqlActFin)
    except:
        f.write (":::::::::::::::>>> ERROR DESCRIPTION2 <<<<:::::::::::::::::" + "\r\n")

#Tabla tblVariant
salida = consultaAccess(consultaCT)
f.write("Salida tblVariant -> " + str(salida) + "\r\n")
#Procesamos la salida
for lin in salida:
    id = lin[0]
    texto = lin[1]
    f.write("Cad. Orig: " + str(id) + " - " + str(texto) + "\r\n")
    if texto:
        texto = texto.replace("'", "''")
        texto = texto.decode('utf-8', 'ignore')
        cadTraducida = traduceCarTec( str(texto) )
        f.write("Cad. Trad: " + str(id) + " - " + str(cadTraducida) + "\r\n")
    else:
        cadTraducida = ''
    sqlVariant = actualizaCT.replace("{CADENA}", cadTraducida)
    sqlVariant = sqlVariant.replace ("{ID}", str(id) )
    #intentamos insertarla
    try:
        salAct = consultaAccess(sqlVariant)
    except:
        f.write (":::::::::::::::>>> ERROR DESCRIPTION2 <<<<:::::::::::::::::" + "\r\n")


f.close()