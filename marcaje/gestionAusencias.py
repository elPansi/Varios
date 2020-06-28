# -*- coding: utf-8 -*-

import time
from datetime import date, timedelta
import logging
import xlwt, xlrd
from libBD import consultaFB
import smtplib
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.MIMEBase import MIMEBase
from email import encoders

reload(sys)
sys.setdefaultencoding("utf-8")

#Preparamos el EMAIL
emailEnvio = 'email@envio.com'
emailRecepcion = ['email@envio.com']
#emailRecepcion = ['email@envio.com']
servidorSMTP = 'smtp.envio.com'
usuarioSMTP = emailEnvio
passSMTP = 'passEmail'


mail = MIMEMultipart()
mail['To'] = ", ".join(emailRecepcion)
mail['From'] = emailEnvio
mail['Subject'] = '52. Gestion de Ausencias'


sqlAusencias = """SELECT U.USERNAME "NOMBRE CORTO", U.FIRSTNAME NOMBRE, U.LASTNAME APELLIDO, D.NAME DEPARTAMENTO
FROM USERS U INNER JOIN DEPARTMENTS D ON U.DEPARTMENTID = D.ID
WHERE U.ID 
NOT IN ( select distinct(A.USERID) from ATTENDANT A 
where 
extract(year from A."WHEN") = extract(year from cast('Now' as date))
and extract(month from A."WHEN") = extract(month from cast('Now' as date))
and extract(day from A."WHEN") = extract(day from cast('Now' as date))
) AND U.STATEID = 1 AND U.ID <> 1
AND D.ID <> 9 --NO COMPOPLAST 
AND U.ID NOT IN --LOS QUE ESTAN EN OBRA
(SELECT r.ENTITYID
FROM PLANNING r
WHERE r.WORKCODE = 10 AND 
datediff(second, cast('01/01/1970 00:00:00.0000' as timestamp),  cast('Now' as timestamp) ) 
between r.STARTDT and r.ENDDT) 
AND U.ID NOT IN -- LOS QUE ESTAN DE BAJA
(SELECT r.ENTITYID
FROM PLANNING r
WHERE r.WORKCODE = 7 AND 
datediff(second, cast('01/01/1970 00:00:00.0000' as timestamp),  cast('Now' as timestamp) ) 
between r.STARTDT and r.ENDDT) -- FALTA COMPROBAR VACACIONES Y DIA FESTIVO"""


rutaLog = ".\\Log\\"
rutaInformes = ".\\Informes\\"
cadTiempo = time.strftime("%Y")+  time.strftime("%m") + time.strftime("%d") +  "_" + time.strftime("%H")  + time.strftime("%M") +  time.strftime("%S")
f = open(rutaLog + "Log_Ausencias_" + cadTiempo + ".txt", "w")
cadArchivo = rutaInformes + cadTiempo + "_Ausencias.xls"

f.write ("Consulta: " + str(sqlAusencias) + "\r\n")

resCons = consultaFB(sqlAusencias)


nExcel = xlwt.Workbook(encoding='latin-1')
nHoja = nExcel.add_sheet(cadTiempo)


#Insertamos las cabeceras
nHoja.write(0,0, u'Detalle de Ausencias: ')
nHoja.write(1,0, u"Nombre Usuario")
nHoja.write(1,1, "Nombre")
nHoja.write(1,2, "Apellido")
nHoja.write(1,23, "Departamento")


f.write("Insertamos las Cabeceras en el Excel" + "\r\n")

nFila = 2
nCol = 0

for rc in resCons:
    alias = rc[0]
    nombre = rc[1]
    apellido = rc[2]
    departamento = rc[3]

    f.write("Escribiendo detalle --> Nombre Corto: " + alias.decode('utf-8', 'ignore' ) +  " | Nombre: " + nombre.decode('utf-8', 'ignore') + " | Apellido: " + apellido.decode('utf-8', 'ignore') + '\r\n')
    nHoja.write(nFila, 0, alias.decode('utf-8', 'ignore' ) )#str() )
    nHoja.write(nFila, 1, nombre.decode('utf-8', 'ignore')  )
    nHoja.write(nFila, 2, apellido.decode('utf-8', 'ignore')  )
    nHoja.write(nFila, 3, departamento.decode('utf-8', 'ignore') )
    nFila = nFila + 1
    f.write(u"Linea añadida" + '\r\n')

f.write("Antes de guardar Excel "+ '\r\n')
nExcel.save(cadArchivo)
f.write("Guardando Excel "+ '\r\n')


archExcel = open(cadArchivo,'rb')
#adjunto = MIMEBase('multipart', 'encrypted')
adjunto = MIMEBase('application', "octet-stream")
adjunto.set_payload(archExcel.read())
archExcel.close()
encoders.encode_base64(adjunto)

adjunto.add_header('Content-Disposition', 'attachment', filename = cadTiempo + "_Ausencias.xls")

mail.attach(adjunto)
servidor = smtplib.SMTP(servidorSMTP, 587)
servidor.starttls()
servidor.ehlo()
servidor.login(usuarioSMTP, passSMTP)
#enviamos el Email
servidor.sendmail(emailEnvio, emailRecepcion, mail.as_string())
f.write ("Correo enviado ...")
servidor.quit()
f.close()
