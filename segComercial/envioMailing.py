# -*- coding: utf-8 -*-

import time
import sys
from datetime import date, timedelta
import logging
import xlwt, xlrd
from libBDCompo import consultaMSSQL
import smtplib, string
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.MIMEBase import MIMEBase
from email import encoders, Utils


reload(sys)
sys.setdefaultencoding("utf-8")

cadTiempo = time.strftime("%Y") + time.strftime("%m") + time.strftime("%d") + time.strftime("%H") + time.strftime("%M") + time.strftime("%S") + "_"
rutaLog = '.\\Log\\'


logger = logging.getLogger("email")
logger.setLevel(logging.DEBUG)
fh = logging.FileHandler(rutaLog + cadTiempo + '_Mailing.log')
fh.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
logger.addHandler(fh)

sqlInsercion = """INSERT INTO gvilla.[dbo].[GV_MAILING]
           ([FECHA]
           ,[EMAIL]
           ,[ERROR])
           VALUES ( GETDATE(), '{EMAIL}', 0)"""

#Preparamos el EMAIL
emailEnvio = 'mailing@envio.com'
emailDestino = [ 'email@recepcio0n.com' ]

servidorSMTP = 'smtp.envio.com'
usuarioSMTP = emailEnvio
passSMTP = 'passEmail'


sqlObtenerEmail = """SELECT DISTINCT LTRIM(RTRIM(C.Email)) Email_CLI
from gcompoplast.dbo.Clientes C INNER JOIN gcompoplast.dbo.ClientesDireccionesEnvio CDE ON C.CodCliente = CDE.CodCliente
INNER JOIN gcompoplast.dbo.ZONAS Z ON Z.CodZona = CDE.CodZona
WHERE --(Z.CodZona IN (1, 4, 5,  6)) and
(C.Email LIKE '%@%.%')
AND (C.Email NOT IN (SELECT EMAIL FROM gvilla.dbo.GV_MAILING WHERE FECHA = CONVERT(DATE, GETDATE(), 101 ) ) )
AND ( C.CodCliente not in ( 1947, 2268, 2085, 2866, 2223, 3083) )
AND (C.EMAIL NOT LIKE '%@terra.%' )
AND (C.EMAIL NOT LIKE '%Ñ%' )
GROUP BY C.CodCliente, C.IDDeCliente, C.NOMBRE, C.NombreComercial, C.Email"""
 
#Cambiar emailRecepcion por emailDestino
emailDestino = consultaMSSQL(sqlObtenerEmail)

error = None

for i,e in enumerate(emailDestino):
    mail = MIMEMultipart('related')
    #mail = MIMEMultipart('message/delivery-status')
    #mail['To'] = ", ".join(emailDestino)
    #print "Enviando mail a: " + str(e[0]) + "\r\n"
	#prueba
    print "Enviando mail a: " + str(e[0]) + "\r\n"
	
    logger.info(str(i+1) + ".- Email a enviar: " + str(e[0]) )
    mail['From'] = emailEnvio
    #-------------------------------------
    mail['To'] = e[0]
	#prueba
    #mail['To'] = e
    #----------------------------------
    mail['Subject'] = 'Asunto del correo'
    mail['Date'] = Utils.formatdate(localtime=True)
    mail['reply-to'] = 'responder@correo.com'
	

    mail['X-Original-To'] = e[0]
	#prueba
    #mail['X-Original-To'] = e
    mail.preamble = ''

    #msgNotifEstadoEntrega = MIMEMultipart("")

    msgAlternative = MIMEMultipart("alternative")
    mail.attach(msgAlternative)

    with open(".\\mailing\\web.html", "r") as fichero:
        datos = fichero.readlines()
        cuerpo = "\n".join(datos[1:])

    msgHtml = MIMEText(cuerpo, 'html')
    msgAlternative.attach(msgHtml)

    img = open(".\\mailing\\imagen.png", "rb")
    msgImage = MIMEImage(img.read())
    img.close()
    msgImage.add_header('Content-ID', '<img>')
    mail.attach(msgImage)
    #mail.attach( MIMEText(cuerpo, 'html' ) )
    logger.info("Preparando envio para: " + mail['To'] + " Desde: " + mail['From'] )
    servidor = smtplib.SMTP(servidorSMTP, 587)
    servidor.ehlo()
    servidor.set_debuglevel(1)
    servidor.starttls()
    servidor.ehlo()
    servidor.login(usuarioSMTP, passSMTP)
    #enviamos el Email
    #logger.info("Contenido del Email: " + mail.as_string() )
    #logger.info("Email: " + str(mail) )
    #servidor.sendmail(emailEnvio, emailDestino, mail.as_string())
    print "Mensaje de: " + mail['From'] + "\r\n"
    print "Mensaje para: " + mail['To'] + "\r\n"
    try:
        servidor.sendmail(mail['From'], mail['To'], mail.as_string() )
        # lo insertamos en la BD
        consultaMSSQL(sqlInsercion.replace("{EMAIL}", str(mail['To']) ))
    except:
        error = str(mail['To'])
    if error:
        logger.info(" ErrorEmail: " + error )
    logger.info ("Correo enviado ...")
    servidor.quit()
    time.sleep(5) #esperamos 5 segundos para el siguiente envio
