# -*- coding: utf-8 -*-

import time
from datetime import date, timedelta
import logging
import xlwt, xlrd
from libBDHoras import consultaFB
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
emailRecepcion = ['email@recepcion.com']
#emailRecepcion = []
servidorSMTP = 'smtp.envio.com'
usuarioSMTP = emailEnvio
passSMTP = 'passEmail'


mail = MIMEMultipart()
mail['To'] = ", ".join(emailRecepcion)
mail['From'] = emailEnvio
mail['Subject'] = '55. Listado de marcajes del Reloj'



sqlReloj = """SELECT u.ID, u.USERNAME, u.FIRSTNAME, u.LASTNAME, d.NAME, u.password, r."WHEN",  
dev.displayname,  r.INOUT
FROM ATTENDANT r inner join users u on r.USERID = u.ID
inner join DEPARTMENTS d on d.id = u.DEPARTMENTID
inner join DEVICES dev on dev.id = r.DEVICEID
where extract(year from r."WHEN") in ( 2017, 2018 )
and u.DEPARTMENTID <> 9 and r.UPDATEINOUT not in (4)
order by r."WHEN" """

rutaLog = ".\\Log\\"
rutaInformes = ".\\Informes\\"
cadTiempo = time.strftime("%Y")+  time.strftime("%m") + time.strftime("%d") +  "_" + time.strftime("%H")  + time.strftime("%M") +  time.strftime("%S")
f = open(rutaLog + "Log_Listado_Marcajes_" + cadTiempo + ".txt", "w")
cadArchivo = rutaInformes + cadTiempo + "_Listado_Marcajes.xls"

f.write ("Consulta: " + str(sqlReloj) + "\r\n")

resCons = consultaFB(sqlReloj)


nExcel = xlwt.Workbook(encoding='latin-1')
nHoja = nExcel.add_sheet(cadTiempo)
date_format = xlwt.XFStyle()
date_format.num_format_str = 'dd/mm/yyyy hh:mm:ss'


#Insertamos las cabeceras
#nHoja.write(0,0, u'Lisato de Marcajes: ')
nHoja.write(1,0, u"ID de usuario")
nHoja.write(1,1, "Nombre de usuario")
nHoja.write(1,2, "Nombre")
nHoja.write(1,3, "Apellido")
nHoja.write(1,4, "Departamento")
nHoja.write(1,5, u"Nº Personal")
nHoja.write(1,6, "Fichaje")
nHoja.write(1,7, "Dispositivo")
nHoja.write(1,8, "Entrada/Salida")

f.write("Insertamos las Cabeceras en el Excel" + "\r\n")

nFila = 2
nCol = 0

for rc in resCons:
    idUsuario = rc[0]
    nombreUsuario = rc[1]
    nombre = rc[2]
    apellido = rc[3]
    departamento = rc[4]
    password = rc[5]
    fichaje = rc[6]
    dispositivo = rc[7]
    entradaSalida = rc[8]
    #id = rc[8]
    f.write("Escribiendo detalle --> IdUsu: " + str(idUsuario) +  " | usuario: " + str(nombreUsuario) + " | depart: "
            + departamento.decode('utf-8', 'ignore') + " | fichaje: " + str(fichaje) + '\r\n')
    nHoja.write(nFila, 0, idUsuario ) #str() )
    nHoja.write(nFila, 1, nombreUsuario) #str(idCli) )
    nHoja.write(nFila, 2, nombre.decode('utf-8', 'ignore')) #str(idCli) )
    nHoja.write(nFila, 3, apellido.decode('utf-8', 'ignore')) #str(idCli) )
    nHoja.write(nFila, 4, departamento.decode('utf-8', 'ignore')  ) #nomCli.decode('utf-8') )
    nHoja.write(nFila, 5, password  ) #nomCli.decode('utf-8') )
    nHoja.write(nFila, 6, fichaje, date_format) #str(impTotal) )
    nHoja.write(nFila, 7, dispositivo )
    #print ("Marcaje: " + str(fichaje) + " || EntradaSalida: " + str(entradaSalida) + "\r\n")
    if entradaSalida == 0:
        nHoja.write(nFila, 8, "Entrada")
    else:
        nHoja.write(nFila, 8, "Salida" )
    nHoja.write(nFila, 9, entradaSalida)
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





