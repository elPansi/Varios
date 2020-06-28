#!/usr/local/bin/python
# -*- coding: utf-8 -*-
import time
import string
import sys
from winReg import *

rutaLog = ".\\Log\\"

cadLog = time.strftime("%d")+ "_" +  time.strftime("%m") + "_" + time.strftime("%Y") +  "_" + time.strftime("%H") +  "_" + time.strftime("%M") +  "_" + time.strftime("%S")
f = open(rutaLog + "appLang_Log_" + cadLog + ".txt","w")

registro = ConnectRegistry(None,HKEY_CURRENT_USER)
key = OpenKey(registro, r'Control Panel\International', 0, KEY_WRITE)

#Cargamos los datos de configuracion
conf = open("confLang.ini","r")
for l in conf:
    if l[:7] == "Locale=":
        locale = l[7:] ##Obtenemos Locale
        #numSuper = numSuper[:-1]
        f.write ("Locale: " + str(locale) + "\r\n")
    elif l[:6] == "iTime=":
        iTime = l[6:] ##Obtenemos la iTime
        #numSeccion = numSeccion[:-1]
        f.write ("iTime: " + iTime + "\r\n" )
    elif l[:6] == "iDate=":
        iDate = l[6:] ##Obtenemos iDate
        #numBalanza = numBalanza[:-1]
        f.write ("iDate: " + iDate + "\r\n" )
    elif l[:11] == "sShortDate=":
        sShortDate = l[11:] ##Obtenemos sShortDate
        #ip = ip[:-1]
        f.write ("sShortDate: " + sShortDate + "\r\n" )
    elif l[:11] == "LocaleName=":
        localeName = l[11:] ##Obtenemos localeName
        #ip = ip[:-1]
        f.write ("LocaleName: " + localeName + "\r\n" )
    elif l[:9] == "sCountry=":
        sCountry = l[9:] ##Obtenemos sCountry
        #ip = ip[:-1]
        f.write ("sCountry: " + sCountry + "\r\n" )
    elif l[:9] == "iCountry=":
        iCountry = l[9:] ##Obtenemos iCountry
        #ip = ip[:-1]
        f.write ("iCountry: " + iCountry + "\r\n" )
    elif l[:10] == "sLanguage=":
        sLanguage = l[10:] ##Obtenemos sLanguage
        #ip = ip[:-1]
        f.write ("sLanguage: " + sLanguage + "\r\n" )
    elif l[:6] == "sList=":
        sList = l[6:] ##Obtenemos sList
        #ip = ip[:-1]
        f.write ("sList: " + sList + "\r\n" )
    elif l[:12] == "sTimeFormat=":
        sTimeFormat = l[12:] ##Obtenemos sTimeFormat
        #ip = ip[:-1]
        f.write ("sTimeFormat: " + sTimeFormat + "\r\n" )
    elif l[:11] == "sShortTime=":
        sShortTime = l[11:] ##Obtenemos sShortTime
        #ip = ip[:-1]
        f.write ("sShortTime: " + sShortTime + "\r\n" )
##Fin Carga

try:
    SetValueEx(key, "Locale", 0, REG_SZ, locale)
    SetValueEx(key, "iTime", 0, REG_SZ, iTime)
    SetValueEx(key, "iDate", 0, REG_SZ, iDate)
    SetValueEx(key, "sShortDate", 0, REG_SZ, sShortDate)
    SetValueEx(key, "LocaleName", 0, REG_SZ, localeName)
    SetValueEx(key, "sCountry", 0, REG_SZ, sCountry)
    SetValueEx(key, "iCountry", 0, REG_SZ, iCountry)
    SetValueEx(key, "sLanguage", 0, REG_SZ, sLanguage)
    SetValueEx(key, "sList", 0, REG_SZ, sList)
    SetValueEx(key, "sTimeFormat", 0, REG_SZ, sTimeFormat)
    SetValueEx(key, "sShortTime", 0, REG_SZ, sShortTime)
except EnvironmentError:
    print "Error al escribir en el registro ..."


