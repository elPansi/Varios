# -*- coding: utf-8 -*-

import time
from datetime import date, timedelta
import logging
import xlwt, xlrd
from libBD2 import consultaMSSQL
import smtplib
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.MIMEBase import MIMEBase
from email import encoders
from smb.SMBConnection import SMBConnection
from nmb.NetBIOS import NetBIOS
import os, shutil

reload(sys)
sys.setdefaultencoding("utf-8")

#Preparamos el EMAIL
emailEnvio = 'email@origen.com'
emailRecepcion = ['email2@destino.com', 'email2@destino.com']
servidorSMTP = 'smtp.origen.com'
usuarioSMTP = emailEnvio
passSMTP = 'passCorreo'

mail = MIMEMultipart()
mail['To'] = ", ".join(emailRecepcion)
mail['From'] = emailEnvio
mail['Subject'] = 'Exportacion de consulta SQL'

nombreExcel="ExpSQL_"
tituloExcel = u'<Listado de facturas> '


#fechas formato YYYY/MM/DD
fini = raw_input("""Fecha Inicial (formato mm/dd/aaaa): """)
fechaIniFmt = fini
ffin = raw_input("""Fecha Final (formato mm/dd/aaaa): """)
fechaFinFmt = ffin

sqlConsulta=u"""select * from clientes"""

sqlConsulta=sqlConsulta.replace("{FECHAINI}",fechaIniFmt)
sqlConsulta=sqlConsulta.replace("{FECHAFIN}",fechaFinFmt)

rutaLog = ".\\Log\\"
rutaInformes = ".\\Informes\\erp\\"
cadTiempo = time.strftime("%Y")+  time.strftime("%m") + time.strftime("%d") +  "_" + time.strftime("%H")  + time.strftime("%M") +  time.strftime("%S")

f = open(rutaLog + "Log_"  + nombreExcel + "_" + cadTiempo + ".txt", "w")
cadArchivo = rutaInformes + cadTiempo + "_" + nombreExcel + ".xls"

f.write ("Consulta: " + str(sqlConsulta) + "\r\n")

resCons, cabeceras = consultaMSSQL(sqlConsulta)

nExcel = xlwt.Workbook(encoding='latin-1')
nHoja = nExcel.add_sheet(cadTiempo)

#nHoja.write(0,0, tituloExcel )
#Insertamos las cabeceras

for i,c in enumerate(cabeceras,1):
    f.write (str(c))
    nHoja.write(0,i-1 , c[0])



f.write("Insertamos las Cabeceras en el Excel" + "\r\n")

nFila = 1
nCol = 0

#Procesamos la consulta al excel
for i,rc in enumerate(resCons):
    linea = ''
    for j,c in enumerate(cabeceras):
        nHoja.write(nFila+i, j, rc[j])
    f.write(u"Lin. " + str(i) + " - " + linea + '\r')



f.write("Antes de guardar Excel "+ '\r\n')
nExcel.save(cadArchivo)
f.write("Guardando Excel "+ '\r\n')


archExcel = open(cadArchivo,'rb')
#adjunto = MIMEBase('multipart', 'encrypted')
adjunto = MIMEBase('application', "octet-stream")
adjunto.set_payload(archExcel.read())
archExcel.close()
encoders.encode_base64(adjunto)

adjunto.add_header('Content-Disposition', 'attachment', filename = cadTiempo + "_" + nombreExcel + ".xls")

mail.attach(adjunto)
servidor = smtplib.SMTP(servidorSMTP, 587)
servidor.starttls()
servidor.ehlo()
servidor.login(usuarioSMTP, passSMTP)
#enviamos el Email
servidor.sendmail(emailEnvio, emailRecepcion, mail.as_string())
f.write ("Correo enviado ...")
servidor.quit()

source = cadArchivo
f.close()