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
emailRecepcion = ['email@envio.com' ]
servidorSMTP = 'smtp.envio.com'
usuarioSMTP = emailEnvio
passSMTP = 'passEmail'


mail = MIMEMultipart()
mail['To'] = ", ".join(emailRecepcion)
mail['From'] = emailEnvio
mail['Subject'] = '50. Marcajes del Reloj'




sqlReloj = """select extract(year from A."WHEN") AS Ano,
extract(month from A."WHEN") AS Mes,
extract(day from A."WHEN") AS Dia, U.ID,
U.FIRSTNAME as Nombre, U.LASTNAME as Apellido, A.INOUT as SalidaEntrada,  count(*) as Veces
from ATTENDANT A INNER JOIN USERS U ON A.USERID = U.ID
where extract(day from A."WHEN") in ( extract(day from cast('Now' as date)), extract(day from cast('Now' as date))-1  )
group by extract(year from A."WHEN"), extract(day from A."WHEN"),
extract(month from A."WHEN"), U.ID, U.FIRSTNAME, U.LASTNAME, A.INOUT"""

rutaLog = ".\\Log\\"
rutaInformes = ".\\Informes\\"
cadTiempo = time.strftime("%Y")+  time.strftime("%m") + time.strftime("%d") +  "_" + time.strftime("%H")  + time.strftime("%M") +  time.strftime("%S")
f = open(rutaLog + "Log_Marcajes_" + cadTiempo + ".txt", "w")
cadArchivo = rutaInformes + cadTiempo + "_Marcajes.xls"

f.write ("Consulta: " + str(sqlReloj) + "\r\n")

resCons = consultaFB(sqlReloj)


nExcel = xlwt.Workbook(encoding='latin-1')
nHoja = nExcel.add_sheet(cadTiempo)


#Insertamos las cabeceras
nHoja.write(0,0, u'Resumen de Marcajes: ')
nHoja.write(1,0, u"Año")
nHoja.write(1,1, "Mes")
nHoja.write(1,2, "Dia")
nHoja.write(1,3, "ID")
nHoja.write(1,4, "Nombre")
nHoja.write(1,5, "Apellido")
nHoja.write(1,6, "Entrada/Salida")
nHoja.write(1,7, "Veces")

f.write("Insertamos las Cabeceras en el Excel" + "\r\n")


nFila = 2
nCol = 0

for rc in resCons:
    ano = rc[0]
    mes = rc[1]
    dia = rc[2]
    id = rc[3]
    nom = rc[4]
    ape = rc[5]
    eos = rc[6]
    vec = rc[7]
    f.write("Escribiendo detalle --> Dia: " + str(dia) +  " | Mes: " + str(mes) + " | Nombre: " + nom.decode('utf-8', 'ignore') + " | Veces: " + str(vec) + '\r\n')
    nHoja.write(nFila, 0, ano ) #str() )
    nHoja.write(nFila, 1, mes) #str(idCli) )
    nHoja.write(nFila, 2, dia) #str(idCli) )
    nHoja.write(nFila, 3, id) #str(idCli) )
    nHoja.write(nFila, 4, nom.decode('utf-8', 'ignore')  ) #nomCli.decode('utf-8') )
    nHoja.write(nFila, 5, ape.decode('utf-8', 'ignore')  ) #nomCli.decode('utf-8') )
    nHoja.write(nFila, 6, eos) #str(impTotal) )
    nHoja.write(nFila, 7, vec )
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

adjunto.add_header('Content-Disposition', 'attachment', filename = cadTiempo + "_Marcajes.xls")

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





