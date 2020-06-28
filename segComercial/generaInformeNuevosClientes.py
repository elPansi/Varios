# -*- coding: cp1252 -*- 
#import sys
#reload(sys)
#sys.setdefaultencoding('utf8')

from libBDWebempresa import consultaMysql
#from libBD import consultaMSSQL
import time, codecs
import xlrd, xlwt
#import locale

rutaLog = ".\\Log\\"
rutaInformes = ".\\Informes\\"
cadTiempo = time.strftime("%d")+ "_" +  time.strftime("%m") + "_" + time.strftime("%Y") +  "_" + time.strftime("%H") +  "_" + time.strftime("%M") +  "_" + time.strftime("%S")
f = open(rutaLog + "Log_" + cadTiempo + "_NuevosClientesCompo" + ".txt", "w")
#m = open(rutaInformes + "Macro_" + cadTiempo + ".txt", "w")

print rutaInformes + cadTiempo
#locale.setlocale(locale.LC_ALL, "es_ES.utf-8")


sqlNuevosClientes = """SELECT Usuario, Cliente, IDTemp, Resultado, Contacto, Euros, Comentarios,
/*latitud, longitud, Cobros, mcc, mnc, cid, lac  */
Cobros
FROM VISITAS
WHERE CLIENTE LIKE '%Nuevo Cliente%'
and year(IDTemp) = {ANO}
and month(IDTemp) = {MES}
ORDER BY Fecha DESC"""


mes = raw_input("Numero del mes: ")
ano = 2018

consNuevosClientes = sqlNuevosClientes.replace("{ANO}", str(ano) )
consNuevosClientes = consNuevosClientes.replace("{MES}", str(mes) )

f.write("Consulta --> " + str(consNuevosClientes) )
respuesta = consultaMysql(consNuevosClientes)
nExcel = xlwt.Workbook()
nHoja = nExcel.add_sheet("NuevosClientes")

#Insertamos las cabeceras
nHoja.write(0,0, "Comercial")
nHoja.write(0,1, "Cliente")
nHoja.write(0,2, "IDTemp")
nHoja.write(0,3, "Resultado")
nHoja.write(0,4, "Contacto")
nHoja.write(0,5, "Euros")
nHoja.write(0,6, "Comentarios")
nHoja.write(0,7, "Cobros")

f.write("Insertamos las Cabeceras en el Excel" + "\r\n")

nFila = 1
nCol = 0


if not respuesta:
    f.write("No existen Clientes en el mes")
    sys.exit(0)

for r in respuesta:
    comercial = r[0]
    cliente = r[1]
    idTemp = r[2]
    resultado = r[3]
    contacto = r[4]
    euros = r[5]
    comentarios = r[6]
    cobros = r[7]
    f.write("Comercial: " + str(comercial) + " | Cliente: " + str(cliente) + " | Tiempo: " + str(
        idTemp) + " | Contacto: " + str(contacto) + " | Comentarios: " + str(comentarios) + "\r\n")

    nHoja.write(nFila, 0, comercial.decode('utf-8', 'ignore') )
    nHoja.write(nFila, 1, cliente.decode('utf-8', 'ignore') )
    nHoja.write(nFila, 2, str(idTemp) )
    nHoja.write(nFila, 3, resultado.decode('utf-8', 'ignore') )
    nHoja.write(nFila, 4, contacto.decode('utf-8', 'ignore') )
    nHoja.write(nFila, 5, euros.decode('utf-8', 'ignore') )
    nHoja.write(nFila, 6, comentarios.decode('utf-8', 'ignore') )
    nHoja.write(nFila, 7, cobros.decode('utf-8', 'ignore') )
    nFila = nFila + 1;


# guardamos el excel
nExcel.save(rutaInformes + cadTiempo + "_NuevosClientesCompo" + ".xls")
f.write("--> Guardamos el Excel\r\n")
f.close()
	


