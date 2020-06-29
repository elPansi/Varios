# -*- coding: utf-8 -*-

import time
from datetime import date, timedelta
import logging
import xlwt, xlrd
import sys, ast
import urllib2
import requests

reload(sys)
sys.setdefaultencoding("utf-8")

nombreExcel="ExpAPI"
rutaLog = ".\\Log\\"
rutaInformes = ".\\Informes\\"
cadTiempo = time.strftime("%Y")+  time.strftime("%m") + time.strftime("%d") +  "_" + time.strftime("%H")  + time.strftime("%M") +  time.strftime("%S")

f = open(rutaLog + "Log_"  + nombreExcel + "_" + cadTiempo + ".txt", "w")
cadArchivo = rutaInformes + cadTiempo + "_" + nombreExcel + ".xls"


tituloExcel = u'<Listado de facturas> '

rutaApi = """http://xxx.syltek.com/Api"""
rutaPck = """http://xxx.syltek.com/pck"""
peticionSesion = """/newsession?apikey=xxx"""
peticionVentas="""/voucherPayments?idPaymentMethod=978&start={FECHAINI}&end={FECHAFIN}&idsession={IDSESION}"""
pPagos="""/payments/all?start={FECHAINI}&end={FECHAFIN}&idsession={IDSESION}"""
cadVentas =rutaPck  + pPagos.replace("{FECHAINI}","01/08/2018")
cadVentas=cadVentas.replace("{FECHAFIN}","18/10/2018")
#f.write("Cad Ventas: " + str(cadVentas) + "\r\n")
f.write("Pet. Sesion: " + str(rutaApi + peticionSesion) + "\r\n")

#Hacemos la peticion de la sesion
sesion = requests.get(rutaApi + peticionSesion)
resp = sesion.content
f.write("Respuesta: " + str(resp) + "\r\n")
print resp
dic = ast.literal_eval(resp)
#Obtenemos el ID de sesion
idSesion = dic.get("sessionId")
cadVentas = cadVentas.replace("{IDSESION}",idSesion)


f.write("Cad. Ventas: " + str(cadVentas) + "\r\n")

cab = requests.head(cadVentas)
f.write("Cabeceras: " + str(cab))
#print(response['headers']['Accept'])

datos = requests.get( cadVentas)
f.write("Datos Bruto: " + str(datos) + "\r\n")
cabeceras = datos.headers
salida = datos.json()
#print salida['ticketnumber'] + "\r\n"

f.write("Cabeceras: " + str(cabeceras) + "\r\n")

#f.write("Ticket : " + salida['headers']['ticketnumber'] + "\r\n")
#f.write("ID Sesion: " + str(idSesion) + " || Contenido: " + str(salida) + "\r\n")
f.write("Datos: "+ str(salida))

cabeceras = salida['Columns']
for i,c in enumerate(cabeceras):
    print str(i) + " - " + c

info = salida['Rows']


#Imprimimos los campos en el Log
for j,i in enumerate(info):
    f.write("Columna: " + str(j) + " | Dia " + i[3] + " | Num. Ticket : " + i[9] + " | Producto: " + i[13].encode('utf-8') +"\r\n")


nExcel = xlwt.Workbook(encoding='latin-1')
#nExcel = xlwt.Workbook(encoding='UTF8')
nHoja = nExcel.add_sheet(cadTiempo)

nHoja.write(0,0, tituloExcel ) #Insertamos las cabeceras

for i,c in enumerate(cabeceras,1):
#     f.write (str(c))
     nHoja.write(1,i-1 , c)
#
f.write("Insertamos las Cabeceras en el Excel" + "\r\n")
#
nFila = 2
nCol = 0
#
#Procesamos la consulta al excel
for i,rc in enumerate(info):
    for j,c in enumerate(cabeceras):
        if rc[j]:
            nHoja.write(nFila+i, j, rc[j])
        f.write(u"Linea añadida: " + str(i) + '\r\n')

f.write("Antes de guardar Excel "+ '\r\n')
nExcel.save(cadArchivo)
f.write("Guardando Excel "+ '\r\n')

f.close()