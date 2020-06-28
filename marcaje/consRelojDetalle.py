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
servidorSMTP = 'smtp.envio.com'
usuarioSMTP = emailEnvio
passSMTP = 'passEmail'


mail = MIMEMultipart()
mail['To'] = ", ".join(emailRecepcion)
mail['From'] = emailEnvio
mail['Subject'] = '51. Detalles de marcajes del Reloj'


sqlReloj = """select extract(year from A."WHEN") AS Ano,
extract(month from A."WHEN") AS Mes,
extract(day from A."WHEN") AS Dia, U.ID,
U.FIRSTNAME as Nombre, U.LASTNAME as Apellido, A.INOUT as SalidaEntrada,
( right(cast((extract(hour from A."WHEN")+100) as varchar(3)),2)||':'||right(cast((extract(minute from A."WHEN")+100) as varchar(3)),2) ) As Hora --, A.ID
from ATTENDANT A INNER JOIN USERS U ON A.USERID = U.ID
where extract(day from A."WHEN") in ( extract(day from cast('Now' as date)) ) -- ,extract(day from cast('Now' as date))-1
AND extract( month from A."WHEN") = extract(month from cast('Now' as date)) 
AND extract( YEAR from A."WHEN") = extract(YEAR from cast('Now' as date)) 
AND A.UPDATEINOUT IS NULL
group by extract(year from A."WHEN"), extract(month from A."WHEN") ,extract(day from A."WHEN"), extract(hour from A."WHEN"), extract(minute from A."WHEN"), U.ID,
u.FIRSTNAME, U.LASTNAME, A.INOUT
order by extract(year from A."WHEN"), extract(month from A."WHEN") ,extract(day from A."WHEN"), U.ID """


rutaLog = ".\\Log\\"
rutaInformes = ".\\Informes\\"
cadTiempo = time.strftime("%Y")+  time.strftime("%m") + time.strftime("%d") +  "_" + time.strftime("%H")  + time.strftime("%M") +  time.strftime("%S")
f = open(rutaLog + "Log_Det_Marcajes_" + cadTiempo + ".txt", "w")
cadArchivo = rutaInformes + cadTiempo + "_Detalle_Marcajes.xls"

f.write ("Consulta: " + str(sqlReloj) + "\r\n")

resCons = consultaFB(sqlReloj)


nExcel = xlwt.Workbook(encoding='latin-1')
nHoja = nExcel.add_sheet(cadTiempo)


#Insertamos las cabeceras
nHoja.write(0,0, u'Detalle de Marcajes: ')
nHoja.write(1,0, u"Año")
nHoja.write(1,1, "Mes")
nHoja.write(1,2, "Dia")
nHoja.write(1,3, "ID Usuario")
nHoja.write(1,4, "Nombre")
nHoja.write(1,5, "Apellido")
nHoja.write(1,6, "Entrada/Salida")
nHoja.write(1,7, "Hora")
#nHoja.write(1,8, "ID")

f.write("Insertamos las Cabeceras en el Excel" + "\r\n")

nFila = 2
nCol = 0

for rc in resCons:
    ano = rc[0]
    mes = rc[1]
    dia = rc[2]
    idu = rc[3]
    nom = rc[4]
    ape = rc[5]
    eos = rc[6]
    hor = rc[7]
    #id = rc[8]
    f.write("Escribiendo detalle --> Dia: " + str(dia) +  " | Mes: " + str(mes) + " | Nombre: " + nom.decode('utf-8', 'ignore') + " | Detalle: " + str(hor) + '\r\n')
    nHoja.write(nFila, 0, ano ) #str() )
    nHoja.write(nFila, 1, mes) #str(idCli) )
    nHoja.write(nFila, 2, dia) #str(idCli) )
    nHoja.write(nFila, 3, idu) #str(idCli) )
    nHoja.write(nFila, 4, nom.decode('utf-8', 'ignore')  ) #nomCli.decode('utf-8') )
    nHoja.write(nFila, 5, ape.decode('utf-8', 'ignore')  ) #nomCli.decode('utf-8') )
    nHoja.write(nFila, 6, eos) #str(impTotal) )
    nHoja.write(nFila, 7, hor )
    #nHoja.write(nFila, 8, id )
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

adjunto.add_header('Content-Disposition', 'attachment', filename = cadTiempo + "_Detalle_Marcajes.xls")

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





