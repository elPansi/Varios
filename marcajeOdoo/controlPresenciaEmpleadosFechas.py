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
emailRecepcion = ['email@recepcion.com' ]
servidorSMTP = 'smtp.envio.com'
emailEnvio = usuarioSMTP
passSMTP = 'passEmail'

mail = MIMEMultipart()
mail['To'] = ", ".join(emailRecepcion)
mail['From'] = emailEnvio
cadAsunto = 'Control Presencia entre fechas: {FECHAINI} y {FECHAFIN}'

tituloExcel = "Presencia"
nombreExcel="Odoo_" + str(tituloExcel)

#FORMATO DE FECHAS -> MM/DD/YYYY

fini = raw_input("""Fecha Inicial (formato dd/mm/aaaa): """)
fechaIniFmt = fini[3:5] + '/' + fini[:2] + '/' + fini[6:10]
#fechaIniFmt = fini[6:10] + '/' + fini[3:5] + '/' + fini[:2]
#print fechaIniFmt
ffin = raw_input("""Fecha Final (formato dd/mm/aaaa): """)
fechaFinFmt = ffin[3:5] + '/' + ffin[:2] + '/' + ffin[6:10]
#print fechaFinFmt

cadAsunto = cadAsunto.replace("{FECHAINI}",fini)
cadAsunto = cadAsunto.replace("{FECHAFIN}",ffin)
mail['Subject'] = cadAsunto


sqlConsulta="""SELECT 
to_char(a.check_in::date, 'DD/MM/YYYY' ) Dia, e.name Empleado, sum(a.worked_hours) Horas
FROM hr_attendance a inner join hr_employee e on e.id = a.employee_id
where a.create_date between '{FECHAINI}' and '{FECHAFIN}'
group by e.name, a.check_in::date
order by a.check_in::date, e.name"""
#Falta comprobar que no esté en las vacaciones
sqlConsulta=sqlConsulta.replace("{FECHAINI}",fechaIniFmt)
sqlConsulta=sqlConsulta.replace("{FECHAFIN}",fechaFinFmt)


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