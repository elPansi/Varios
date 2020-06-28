# -*- coding: utf-8 -*-

import time
import sys
from datetime import date, timedelta
import logging
import xlwt, xlrd
from libBDCompo import consultaMSSQL
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.MIMEBase import MIMEBase
from email import encoders, Utils


reload(sys)
sys.setdefaultencoding("utf-8")

#Preparamos el EMAIL
emailEnvio = 'envio@email.com'
emailRecepcion = ['email@recepcion.com']
servidorSMTP = 'smtp.recepcion.com'
usuarioSMTP = emailEnvio
passSMTP = 'passEmail'

mes = raw_input("Numero del mes: ")
ano = 2018

mail = MIMEMultipart()
mail['To'] = ", ".join(emailRecepcion)
mail['From'] = emailEnvio
mail['Subject'] = '61.COMPO. Ventas mensuales: (' + str(mes) + ' / ' + str(ano) + ')'
mail['Date'] = Utils.formatdate(localtime=True)


sqlVentas = """SELECT C.CodCliente 'CodigoCliente', C.IDDeCliente 'IdCliente', C.CIF 'CIF',
C.Nombre 'Nombre Cliente', C.DIRECCION 'Direccion', C.Telefono1 'Telefono',
C.Poblacion 'Poblacion', R.Nombre 'Comercial', prov.Nombre 'Provincia', prov.id IdProvincia,
C.Email, C.NombreContacto, SUM(AD.Cantidad * AD.Precio * (1 - AD.Dto1/100) ) 'Base Imp. Vtas.'
FROM ARTICULOSDETALLES AD INNER JOIN FacturaCab FC ON AD.CodFactura = FC.CodFactura
INNER JOIN CLIENTES C ON C.CodCliente = FC.CodCliente
INNER JOIN provincias PROV ON PROV.codigo = C.CodProvincia
INNER JOIN Representantes R ON R.CodRepresentante = FC.CodComercial
WHERE MONTH(FC.FechaFactura) = {MES} AND YEAR(FC.FECHAFACTURA) = {ANO}
--AND PROV.ID IN (4, 18, 23, 29)
GROUP BY PROV.Nombre, prov.id, C.CodCliente, C.IDDeCliente, C.Nombre, C.Direccion, C.Telefono1,
C.Poblacion, R.Nombre, C.CIF, C.NombreContacto, C.Email"""

consVentas = sqlVentas.replace("{ANO}", str(ano) )
consVentas = consVentas.replace("{MES}", str(mes) )

rutaLog = ".\\Log\\"
rutaInformes = ".\\Informes\\"
cadTiempo = time.strftime("%Y")+  time.strftime("%m") + time.strftime("%d") +  "_" + time.strftime("%H")  + time.strftime("%M") +  time.strftime("%S")
f = open(rutaLog + "Log_Vtas_Compo_" + cadTiempo + ".txt", "w")
cadArchivo = rutaInformes + cadTiempo + "_VtasCompo.xls"

#Recuperamos las ventas.
ventas = consultaMSSQL( consVentas )

nExcel = xlwt.Workbook()
nHoja = nExcel.add_sheet(cadTiempo)


#Insertamos las cabeceras
nHoja.write(0,0, "Cod. Cliente")
nHoja.write(0,1, "Id. Cliente")
nHoja.write(0,2, "CIF")
nHoja.write(0,3, "Nombre Cliente")
nHoja.write(0,4, "Direccion")
nHoja.write(0,5, "Telefono")
nHoja.write(0,6, "Poblacion")
nHoja.write(0,7, "Comercial")
nHoja.write(0,8, "Provincia")
nHoja.write(0,9, "Id. Provincia")
nHoja.write(0,10, "Email")
nHoja.write(0,11, "Nombre Contacto")
nHoja.write(0,12, "Vtas. (Base Imp.)")

f.write("Insertamos las Cabeceras en el Excel" + "\r\n")

nFila = 1
nCol = 0

if not ventas:
    f.write("No existen ventas en el periodo")
    sys.exit(0)

for v in ventas:
    codCliente = v[0]
    idCliente = v[1]
    cif = v[2]
    nombreCliente = v[3]
    direccion = v[4]
    telefono = v[5]
    poblacion = v[6]
    comercial = v[7]
    provincia = v[8]
    idProvincia = v[9]
    email = v[10]
    nombreContacto = v[11]
    vtasBaseImp = v[12]

    f.write ( "IDdeCliente: " + str(idCliente) + " | CIF: " + str(cif) + " | Nombre: " + str(nombreCliente) + " | Base Imponible: " + str(vtasBaseImp) + "\r\n" )
    nHoja.write(nFila, 0, codCliente )
    nHoja.write(nFila, 1, idCliente )
    if cif:
        nHoja.write(nFila, 2, str(cif) )
    if nombreCliente:
        nHoja.write(nFila, 3, nombreCliente.decode('utf-8', 'ignore') )
    if direccion:
        nHoja.write(nFila, 4, direccion.decode('utf-8', 'ignore') )
    if telefono:
        nHoja.write(nFila, 5, telefono.decode('utf-8', 'ignore') )
    if poblacion:
        nHoja.write(nFila, 6, poblacion.decode('utf-8', 'ignore') )
    nHoja.write(nFila, 7, comercial.decode('utf-8', 'ignore') )
    nHoja.write(nFila, 8, provincia.decode('utf-8', 'ignore') )
    nHoja.write(nFila, 9, idProvincia )
    if email:
        nHoja.write(nFila,10, email.decode('utf-8', 'ignore') )
    if nombreContacto:
        nHoja.write(nFila,11, nombreContacto.decode('utf-8', 'ignore') )
    nHoja.write(nFila,12, vtasBaseImp )
    nFila = nFila + 1

nExcel.save(cadArchivo)
f.write("Guardando Excel ")


archExcel = open(cadArchivo,'rb')
#adjunto = MIMEBase('multipart', 'encrypted')
adjunto = MIMEBase('application', "octet-stream")
adjunto.set_payload(archExcel.read())
archExcel.close()
encoders.encode_base64(adjunto)

adjunto.add_header('Content-Disposition', 'attachment', filename = cadTiempo + "_VtasCompo.xls")

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





