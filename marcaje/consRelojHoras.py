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
emailRecepcion = ['a@recepcion.com']
servidorSMTP = 'smtp.envio.com'
usuarioSMTP = emailEnvio
passSMTP = 'passEmail'


mail = MIMEMultipart()
mail['To'] = ", ".join(emailRecepcion)
mail['From'] = emailEnvio
mail['Subject'] = '51. Detalles de marcajes del Reloj'

sqlRelojHoras = """SELECT 
A.USERID, CAST(A."WHEN" AS DATE) DIA, 
MIN(A."WHEN") ENTRADA, MAX(A."WHEN") SALIDA,
DATEDIFF( HOUR, MIN(A."WHEN") , MAX(A."WHEN") ) HORAS
FROM ATTENDANT A
WHERE USERID = {USUARIO} AND EXTRACT(YEAR FROM A."WHEN") = {ANO}
AND EXTRACT(MONTH FROM A."WHEN") = {MES} and extract(day from A."WHEN") = {DIA}
GROUP BY A.USERID, CAST(A."WHEN" AS DATE)
ORDER BY A.USERID, CAST(A."WHEN" AS DATE)"""

sqlSiguienteEntrada = """SELECT FIRST 1
A.USERID, A."WHEN"
FROM ATTENDANT A
WHERE USERID = {USUARIO} AND EXTRACT(YEAR FROM A."WHEN") = {ANO}
AND A.INOUT = 0 AND A."WHEN" > '{ENTRADA}'
AND EXTRACT(MONTH FROM A."WHEN") = {MES} and extract(day from A."WHEN") = {DIA}
ORDER BY A.USERID, CAST(A."WHEN" AS DATE)"""

sqlSiguienteSalida = """SELECT FIRST 1
A.USERID, A."WHEN"
FROM ATTENDANT A
WHERE USERID = {USUARIO} AND EXTRACT(YEAR FROM A."WHEN") = {ANO}
AND A.INOUT = 1 AND A."WHEN" > '{ENTRADA}'
AND EXTRACT(MONTH FROM A."WHEN") = {MES} and extract(day from A."WHEN") = {DIA}
ORDER BY A.USERID, CAST(A."WHEN" AS DATE)"""

sqlTrabajadores = """select id, firstname, lastname from users where id > 1 and departmentid <> 9"""

sqlAnosTrabajadores = """SELECT 
distinct extract(year from A."WHEN")
FROM ATTENDANT A
WHERE A.USERID = {USUARIO}
and extract(year from A."WHEN") IN (2017, 2018)
ORDER BY A."WHEN" """

sqlMesesTrab = """SELECT 
distinct extract(month from A."WHEN")
FROM ATTENDANT A
WHERE A.USERID ={USUARIO}  and extract(year from A."WHEN" ) = {ANO}
ORDER BY A."WHEN" """

sqlDiasTrab = """SELECT 
distinct extract(day from A."WHEN")
FROM ATTENDANT A
WHERE A.USERID = {USUARIO} and extract(year from A."WHEN" ) = {ANO}
AND EXTRACT(MONTH FROM A."WHEN") = {MES}
ORDER BY A."WHEN" """

rutaLog = ".\\Log\\"
rutaInformes = ".\\Informes\\"
cadTiempo = time.strftime("%Y")+  time.strftime("%m") + time.strftime("%d") +  "_" + time.strftime("%H")  + time.strftime("%M") +  time.strftime("%S")
f = open(rutaLog + "Log_Det_Marcajes_Horas_" + cadTiempo + ".txt", "w")
cadArchivo = rutaInformes + cadTiempo + "_Detalle_Marcajes_Horas.xls"

f.write ("Consulta Trabajadores: " + str(sqlTrabajadores) + "\r\n")

trabajadores = consultaFB(sqlTrabajadores)



nExcel = xlwt.Workbook(encoding='latin-1')
nCol = 0
# date_format = xlwt.XFStyle()
# date_format.num_format_str = 'dd/mm/yyyy'


for t in trabajadores:
    f.write("Salida Consulta Trabajadores: " + str(t[0]) + "\r\n"  )
    codTrabajador = str(t[0])
    f.write ("Procesando Trabajador: " + str(codTrabajador) + "\r\n" )
    consRelojHoras = sqlRelojHoras.replace("{USUARIO}", codTrabajador )
    #tablaHoras = consultaFB(consRelojHoras)
    consAnosTrab = sqlAnosTrabajadores.replace("{USUARIO}", codTrabajador )
    f.write(u"Consulta Años: " + "\r\n"+ str(consAnosTrab) + "\r\n" )
    anosTrab = consultaFB(consAnosTrab)
    for a in anosTrab:
        ano = str(a[0])
        consMesesTrab = sqlMesesTrab.replace("{ANO}", ano )
        consMesesTrab = consMesesTrab.replace("{USUARIO}", codTrabajador )
        f.write(u"Consulta Meses: " + "\r\n"+ str(consMesesTrab) + "\r\n" )
        mesesTrab = consultaFB(consMesesTrab)
        for m in mesesTrab:
            mes = str(m[0])
            nomPestana = unicode( codTrabajador + "_" + ano + "_" + mes )
            nHoja = nExcel.add_sheet(nomPestana)
            consDiasTrab = sqlDiasTrab.replace("{USUARIO}", codTrabajador)
            consDiasTrab = consDiasTrab.replace("{ANO}", ano)
            consDiasTrab = consDiasTrab.replace("{MES}", mes)
            f.write(u"Consulta Dias: "+ "\r\n" + str(consDiasTrab) + "\r\n" )
            diasTrab = consultaFB(consDiasTrab)
            for d in diasTrab:
                dia = str(d[0])
                consHorasDia = sqlRelojHoras.replace("{USUARIO}", codTrabajador)
                consHorasDia = consHorasDia.replace("{ANO}", ano)
                consHorasDia = consHorasDia.replace("{MES}", mes)
                consHorasDia = consHorasDia.replace("{DIA}", dia)
                f.write(u"Consulta Horas: "+ "\r\n" + str(consHorasDia) + "\r\n")
                entSal = consultaFB(consHorasDia)
                entrada = str(entSal[0][2])
                salida = str(entSal[0][3])
                horas = entSal[0][4]
                entradaTemp = entrada
                f.write("Entrada: " + str(entrada) + "\r\n")
                f.write("salida: " + str(salida) + "\r\n")
                f.write("entsal: " + str(entSal) + "\r\n" )
                f.write("EntradaTemp: " + str(entradaTemp) + "\r\n")
                #print str(entradaTemp)
                descansos = 0
                f.write ("Cod. Trab: " + codTrabajador + u" || Año: " + ano + u" || Mes: " + mes + u" || Dia: "  +
                         dia + "\r\n")
                activo = True
                f.write("Tipo entradaTemp: " + str(type(entradaTemp)) + " || Tipo salida: " + str(type(salida)) + "\r\n" )
                #print str(entradaTemp) + " --- " + str(salida)
                while(entradaTemp < salida and activo):
                    consSigEnt = sqlSiguienteEntrada.replace("{USUARIO}", codTrabajador)
                    consSigEnt = consSigEnt.replace("{ANO}", ano)
                    consSigEnt = consSigEnt.replace("{MES}", mes)
                    consSigEnt = consSigEnt.replace("{DIA}", dia)
                    consSigEnt = consSigEnt.replace("{ENTRADA}", str(entradaTemp) )
                    f.write(u"---------------> Consulta siguiente entrada: " +"\r\n"  + str(consSigEnt) + "\r\n")
                    salidaTemp = entradaTemp
                    resp = consultaFB(consSigEnt)
                    f.write(u"---------------> Restpuesta ----> " + "\r\n" + str(resp) + "\r\n" +
                            "---> " ) #+ str(resp[1]) )
                    entradaTemp = resp
                    f.write("SalidaTemp: " + str(salidaTemp) + " || EntradaTemp: " + str(entradaTemp) + "\r\n")
                    #ahora buscamos la siguiente salida
                    consSigSal = sqlSiguienteSalida.replace("{USUARIO}", codTrabajador)
                    consSigSal = consSigSal.replace("{ANO}", ano)
                    consSigSal = consSigSal.replace("{MES}", mes)
                    consSigSal = consSigSal.replace("{DIA}", dia)
                    consSigSal = consSigSal.replace("{ENTRADA}", str(salidaTemp) )
                    f.write(u"-----------------> Consulta siguiente salida: "+ "\r\n" + str(consSigSal) + "\r\n")
                    resp = consultaFB(consSigSal)
                    salidaTemp = resp[1]
                    f.write("Tipo entradaTemp: " + str(type(entradaTemp)) + " || Tipo salida: " + str(type(salida)) + " || " +
                            "Tipo salidaTemp: " + str(type(salidaTemp)) )
                    f.write("SalidaTemp: " + str(salidaTemp) + " || EntradaTemp: " + str(entradaTemp) + "\r\n" )
                    if len(salidaTemp)==0:
                        activo = False
                    else:
                        f.write("Descansos: " + str(salidaTemp) + " - " + str(entradaTemp))
                        descansos = descansos + (salidaTemp - entradaTemp)
                horasFinales = horas - descansos
                f.write ("Cod. Trab: " + codTrabajador + u" || Año: " + ano + u" || Mes: " + mes + u" || Dia: "  +
                         dia + u" || Horas: " + str(horasFinales) + u" || Descanso: " + str(descansos) )
                f.write("\r\n------------------------------------------------------ FIN DIA -----------------------------------------------------------\r\n")

				
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





