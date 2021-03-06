# -*- coding: utf-8 -*-

import time
from datetime import date, timedelta
import logging
import xlwt, xlrd
from libBD2 import consultaPostgreSQL
import smtplib
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.MIMEBase import MIMEBase
from email import encoders

reload(sys)
sys.setdefaultencoding("utf-8")

#Preparamos el EMAIL
usuarioSMTP = 'email@envio.com'
emailRecepcion = ['email@recepcion']
servidorSMTP = 'smtp.envio.com'
emailEnvio = usuarioSMTP
passSMTP = 'passEmail'


mail = MIMEMultipart()
mail['To'] = ", ".join(emailRecepcion)
mail['From'] = emailEnvio
mail['Subject'] = 'Empleados que no han ticado hoy'

tituloExcel = "Ausencias"
nombreExcel="Odoo_" + str(tituloExcel)


sqlConsulta="""select to_char(now(), 'DD/MM/YYYY') Fecha, name
from hr_employee where id not in
(
select 
--e.name, t.check_in
e.id
from hr_attendance t inner join hr_employee e on e.id = t.employee_id
where DATE_PART('day', t.check_in ) = DATE_PART ('day', now() )  
and DATE_PART('month', t.check_in ) = DATE_PART ('month', now() )  
and DATE_PART('year', t.check_in ) = DATE_PART ('year', now() )  
) and active = 'true'
order by name"""
#Falta comprobar que no esté en las vacaciones

rutaLog = ".\\Log\\"
rutaInformes = ".\\Informes\\"
cadTiempo = time.strftime("%Y")+  time.strftime("%m") + time.strftime("%d") +  "_" + time.strftime("%H")  + time.strftime("%M") +  time.strftime("%S")

f = open(rutaLog + "Log_"  + nombreExcel + "_" + cadTiempo + ".txt", "w")
cadArchivo = rutaInformes + cadTiempo + "_" + nombreExcel + ".xls"

f.write ("Consulta: " + str(sqlConsulta) + "\r\n")

resCons, cabeceras = consultaPostgreSQL(sqlConsulta)

nExcel = xlwt.Workbook(encoding='utf-8')
nHoja = nExcel.add_sheet(cadTiempo)
filaCabecera = 0
filaDatos = 1

#nHoja.write(0,0, tituloExcel )
#Insertamos las cabeceras

for i,c in enumerate(cabeceras,1):
    f.write (str(c))
    nHoja.write(filaCabecera,i-1 , c[0])

f.write("Insertamos las Cabeceras en el Excel" + "\r\n")
nFila = filaDatos
nCol = 0

#Procesamos la consulta al excel
for i,rc in enumerate(resCons):
    for j,c in enumerate(cabeceras):
        if rc[j]:
            nHoja.write(nFila+i, j, rc[j])
    f.write(u"Linea añadida: " + str(i) + '\r\n')



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


try:
    servidor = smtplib.SMTP(servidorSMTP, 587)
    servidor.set_debuglevel(1)
    servidor.ehlo()
    servidor.starttls()
    servidor.ehlo()
    #servidor.esmtp_features['auth'] = 'LOGIN PLAIN'
    servidor.login(usuarioSMTP, passSMTP)
    servidor.sendmail(mail['From'], mail['To'], mail.as_string())
    f.write("Correo enviado ...")
    servidor.quit()
except:
    error = str(mail['To'])
    f.write("Error al enviar email: " + str(error) )
    #logger.info("Error al enviar email")

f.close()