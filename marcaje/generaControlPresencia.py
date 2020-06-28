# -*- coding: cp1252 -*- 
#import sys
#reload(sys)
#sys.setdefaultencoding('utf8')

import sys

reload(sys)
sys.setdefaultencoding('utf8')


from libBDSport import consultaMysql
from libBDSport import consultaMSSQL
import time, codecs
import xlrd, xlwt
#import locale

rutaLog = ".\\Log\\"
rutaInformes = ".\\Informes\\"
cadTiempo = time.strftime("%d")+ "_" +  time.strftime("%m") + "_" + time.strftime("%Y") +  "_" + time.strftime("%H") +  "_" + time.strftime("%M") +  "_" + time.strftime("%S")
f = open(rutaLog + "Log_" + cadTiempo + ".txt", "w")
#m = open(rutaInformes + "Macro_" + cadTiempo + ".txt", "w")

print rutaInformes + cadTiempo
#locale.setlocale(locale.LC_ALL, "es_ES.utf-8")

consDistUsuarios = u"""SELECT LEFT(T.NOMBRE,3) + '_' + LEFT(T.Apellido1,3) + '_' + LEFT(T.Apellido2,3) TRABAJADOR,
T.Nombre Nombre, T.Apellido1 Apellido, T.Apellido2 Apellido2
FROM people.DBO.People T INNER JOIN PEOPLE.DBO.People_Grupos PG ON T.Id = PG.Id_People
INNER JOIN controltornos.DBO.Tokens_Accesos CA ON CA.Id_People = t.Id
WHERE PG.Id_Grupo = 1 AND CA.Fecha_Hora BETWEEN GETDATE() -31 AND GETDATE()
GROUP BY LEFT(T.NOMBRE,3) + '_' + LEFT(T.Apellido1,3) + '_' + LEFT(T.Apellido2,3), T.Nombre, T.Apellido1, T.Apellido2
order BY LEFT(T.NOMBRE,3) + '_' + LEFT(T.Apellido1,3) + '_' + LEFT(T.Apellido2,3)"""

consDetInforme = """SELECT
CONVERT(NVARCHAR(12), CA.Fecha_Hora, 103) DIA,
(
SELECT CONVERT(NVARCHAR(12), MIN(Fecha_Hora), 108)
FROM controltornos.DBO.Tokens_Accesos
WHERE Id_People = CA.Id_People and Tipo ='entrada'
and (
(CONVERT(NVARCHAR(12), Fecha_Hora, 112) = ( CONVERT(NVARCHAR(12), CA.Fecha_Hora, 112)
) ) ) ) ENTRADA,
(
SELECT ISNULL(CONVERT(NVARCHAR(12), MAX(Fecha_Hora), 108),0)
FROM controltornos.DBO.Tokens_Accesos
WHERE Id_People = CA.Id_People and Tipo ='entrada'
and (
(CONVERT(NVARCHAR(12), Fecha_Hora, 112) = ( CONVERT(NVARCHAR(12), CA.Fecha_Hora, 112)
) ) ) ) SALIDA
------------------------------- TABLA TEMPORAL --------------------------
FROM people.DBO.People T INNER JOIN PEOPLE.DBO.People_Grupos PG ON T.Id = PG.Id_People
INNER JOIN controltornos.DBO.Tokens_Accesos CA ON CA.Id_People = t.Id
WHERE PG.Id_Grupo = 1 AND CA.Fecha_Hora BETWEEN GETDATE() -31 AND GETDATE()
AND Nombre = '{NOMBRE}' AND Apellido1 = '{APELLIDO1}' AND Apellido2 = '{APELLIDO2}'
GROUP BY CONVERT(NVARCHAR(12),CA.Fecha_Hora, 112 ),
CONVERT(NVARCHAR(12), CA.Fecha_Hora, 103), CA.Id_People
order by CONVERT(NVARCHAR(12),CA.Fecha_Hora, 112  ) DESC
-------------------------------------------------------------------------
"""



usuarios = consultaMSSQL(consDistUsuarios)
nExcel = xlwt.Workbook()

for i,u in enumerate(usuarios):
    #Para cada usuario generaremos una Hoja
    nFila = 0
    nCol = 0
    usuario = u[0]
    nombre = u[1]
    apellido1 = u[2]
    apellido2 = u[3]
    if str(nombre) == 'None':
        nombre = ""
    if str(apellido1) == 'None':
        apellido1 = ""
    if str(apellido2) == 'None':
        apellido2 = ""
    #nHoja = nExcel.add_sheet(str(usuario).encode('utf8'))
    nHoja = nExcel.add_sheet(unicode(usuario, errors='replace') )
    #nHoja = nExcel.add_sheet(str(i))

    #Insertamos el nombre del comercial como titulo
    nHoja.write(nFila, 1, "Control de presencia - " + unicode(nombre ,errors='replace' ) + " " + unicode(apellido1, errors='replace' ) + " " + unicode(apellido2, errors='replace' ) )
    #Dejamos un hueco
    nFila = nFila + 2
    #Insertamos las cabeceras
    nHoja.write(nFila, nCol, "DIA")
    nCol = nCol + 1
    nHoja.write(nFila, nCol, "ENTRADA")
    nCol = nCol + 1
    nHoja.write(nFila, nCol, "SALIDA")
    nCol = 0
    nFila = nFila + 1;
    #Fin de cabeceras
	
    sqlDetalle = consDetInforme.replace("{NOMBRE}", str(nombre) )
    sqlDetalle = sqlDetalle.replace("{APELLIDO1}", str(apellido1) )
    sqlDetalle = sqlDetalle.replace("{APELLIDO2}", str(apellido2) )

    f.write("--> Cons. SQL: " + str(sqlDetalle) + "\r\n")  
    detInforme = consultaMSSQL(sqlDetalle)
    f.write("--> Generando Pestaña: " + usuario + cadTiempo + "\r\n" )
    for d in detInforme:
        #FECHA
        #f.write("--> Anadimos Fecha \r\n");
        nHoja.write(nFila, nCol, str(d[0]).encode('utf8'))
        nCol = nCol + 1;
        #ENTRADA
        nHoja.write(nFila, nCol, str(d[1]).encode('utf8'))
        nCol = nCol + 1;
        #SALIDA
        nHoja.write(nFila, nCol, str(d[2]).encode('utf8'))
        nCol = 0
        nFila = nFila + 1;

# guardamos el excel
f.write("--> Ruta informes: " + str(rutaInformes) + "\r\n")
f.write("--> Cadena Tiempo: " + str(cadTiempo) + "\r\n")
f.write("--> Ruta completa: " + str(rutaInformes) + "hola" + ".xls" + "\r\n")
nExcel.save(str(rutaInformes) + str(cadTiempo) + ".xls")

f.write("--> Guardamos el Excel\r\n")
#m.close()
f.close()
	


