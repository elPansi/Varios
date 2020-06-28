# -*- coding: utf-8 -*-

import zeep
import requests
import time, sys
import logging
import ast
import xlwt, xlrd
from libBD import consultaMSSQL
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.MIMEBase import MIMEBase
from email import encoders, Utils
from requests import Session
from zeep import Client
from zeep.wsse.signature import Signature
from zeep.exceptions import Fault
from zeep.transports import Transport
from zeep.helpers import serialize_object


reload(sys)
sys.setdefaultencoding('utf8')

#Preparamos el EMAIL
emailEnvio = 'email@envio.com'
emailRecepcion =  [ 'email@envio.com', 'email@recepcion.com']
servidorSMTP = 'smtp.envio.com'
usuarioSMTP = emailEnvio
passSMTP = 'passEmail'

mail = MIMEMultipart()
mail['To'] = ", ".join(emailRecepcion)
mail['From'] = emailEnvio
mail['Subject'] = '07.SII. Registro de Facturas Emitidas'
mail['Date'] = Utils.formatdate(localtime=True)

rutaLog = '.\\Log\\'
rutaInformes = ".\\Informes\\"
rutaCertificados = '.\\certificados\\'
rutaDatos = '.\\datos\\'

#Configuracion de las series de Facturas y series de Facturas rectificativas
serieFacturas = '17'
serieFacturasRectificativas = 'R17'
exentaIVA = 0
sinSalida = 1
mes = raw_input("""Numero mes (año 2018): """)
anio = 2018

#solo para pruebas
#serieFactura = raw_input("Serie Factura: ")
#numeroFactura = raw_input("Numero Factura: ")
# serieFactura = 34
# numeroFactura = 680

cadTiempo = time.strftime("%Y") + time.strftime("%m") + time.strftime("%d") + time.strftime("%H") + time.strftime("%M") + time.strftime("%S") + "_" 
cadArchivo = rutaInformes + cadTiempo + str(mes).zfill(2) + "_factEmitidasPresentadas.xls"

logger = logging.getLogger('zeep')
logger.setLevel(logging.DEBUG)
fh = logging.FileHandler(rutaLog + cadTiempo + str(mes).zfill(2) + '_consulta_FE.log')
fh.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
logger.addHandler(fh)




#mes = 6


logger.info( 'Inicio de la consulta. Ano: ' + str(anio) + " | Mes: " + str(mes).zfill(2) )

##Cadenas ...
nExcel = xlwt.Workbook()
nHoja = nExcel.add_sheet(cadTiempo)

#Insertamos las cabeceras
nHoja.write(0,0, "Facturas Emitidas Presentadas: ")
nHoja.write(1,0, "Ejercicio")
nHoja.write(1,1, "Periodo")
nHoja.write(1,2, "Nif")
nHoja.write(1,3, "NumSerieFE")
nHoja.write(1,4, "FechaExpedicionFE")
nHoja.write(1,5, "TipoFactura")
nHoja.write(1,6, "TipoRectificativa")
nHoja.write(1,7, "NumSerieFE")
nHoja.write(1,8, "FechaExpedicionFE")
nHoja.write(1,9, "ClaveRegEspOTrascencencia")
nHoja.write(1,10, "ClaveRegEspOTrascencenciaAdicional1")
nHoja.write(1,11, "ClaveRegEsplOTrascencenciaAdicional2")
nHoja.write(1,12, "DescripcionOperacion")
nHoja.write(1,13, "InversionDelSujetoPasivo")
nHoja.write(1,14, "TipoImpositivo")
nHoja.write(1,15, "BaseImponible")
nHoja.write(1,16, "CuotaSoportada")
nHoja.write(1,17, "NombreRazon")
nHoja.write(1,18, "FechaOperacion")
nHoja.write(1,19, "TipoNoExenta")
nHoja.write(1,20, "NifPresentador")
nHoja.write(1,21, "TimeStampPresentacion")
nHoja.write(1,22, "CSV")
nHoja.write(1,23, "EstadoCuadre")
nHoja.write(1,24, "TimeStampEstadoCuadre")
nHoja.write(1,25, "TimeStampUltimaModificacion")
nHoja.write(1,26, "EstadoRegistro")
nHoja.write(1,27, "CodigoErrorRegistro")
nHoja.write(1,28, "DescripcionErrorRegistro")
nHoja.write(1,29, "DatosDescuadreContraparte")



#Preparamos la cadena del diccionario
cadCabecera = """{
	"IDVersionSii":1.0,
	"Titular":{
		"NombreRazon":"EMPRESA S.L.",
		"NIF":"CIF-EMPRESA"
		}
	}"""

#FACTURAS
cadFactura="""{
    "PeriodoImpositivo":{
		"Ejercicio":"{ANIO}",
		"Periodo":"{MES}"
		}
	}"""

cadFactFinal = cadFactura.replace("{ANIO}", str(anio) )
cadFactFinal = cadFactFinal.replace("{MES}", str(mes).zfill(2) )


# Convertimos cadenas a Diccionarios
dicCabecera = ast.literal_eval(cadCabecera)
dicFactura = ast.literal_eval(cadFactFinal)


# print dicCabecera
# print '\r\n'
# print dicFactura


#WEBSERVICE AEAT
#wsdl = 'http://www.agenciatributaria.es/static_files/AEAT/Contenidos_Comunes/La_Agencia_Tributaria/Modelos_y_formularios/Suministro_inmediato_informacion/FicherosSuministros/V_07/SuministroFactEmitidas.wsdl'
#wsdl = '/wlpl/SSII-FACT/ws/fe/SiiFactFEV1SOAP/SuministroFactEmitidasPrueba.wsdl'
# Descargar los wsdl de la aeat y modificar la url del servicio por la del servicio de pruebas ya que si no te manda a XXXX
## ------------------------- OJO CREO QUE NO ESTAMOS UTILIZANDO EL SERVICIO DE PRUEBAS --------

wsdl = rutaDatos + 'SuministroFactEmitidas.wsdl'
#wsdl = "./WSDLPruebas/SuministroFactEmitidasPrueba.wsdl"

#SSL CONECTION
session = Session()
session.cert = (rutaCertificados + "clavePublica.pem", rutaCertificados + "clavePrivada.pem")
session.verify = True

transport = Transport(session=session)
#logger.debug()
#FIRMA XML ENVIO
client = Client(wsdl=wsdl,port_name="SuministroFactEmitidas",transport=transport, service_name = 'siiService')
logger.info('Establecida conexion con cliente')
#SELECCION SERVICIO DE PRUEBAS        SuministroLRFacturasEmitidas
service2 = client.bind('siiService', 'SuministroFactEmitidas')
logger.info('Bind al servicio')
#INSERCION DE LA FACTURA Y RESPUESTA EN EL SERVICO DE PRUEBAS.
#print (service2.SuministroLRFacturasEmitidas(cabecera,facturas))
salida = (service2.ConsultaLRFacturasEmitidas(dicCabecera, dicFactura ))
logger.info("--> Salida: --> '" + str(salida) )
logger.info('Consulta finalizada.')

nFila = 2
nCol = 0

#Comprobamos si hay salida
if salida['RegistroRespuestaConsultaLRFacturasEmitidas']:
    dPeriodo = salida['PeriodoImpositivo']
    sinSalida = 0 #hay salida
else:
    logger.info("La peticion no ha devuelto datos")
    sys.exit(0)


if sinSalida == 0: #Procesamos la info.
    ejercicio = dPeriodo['Ejercicio']
    periodo = dPeriodo['Periodo']
    resultado = salida['RegistroRespuestaConsultaLRFacturasEmitidas']
    logger.info('Resultado: ' + str(resultado) )

    for r in resultado:
        entrega = 0
        desIVA = None
        detIVA = None
        idFact = r['IDFactura']
        idEmiFact = idFact['IDEmisorFactura']
        datFactRec = r['DatosFacturaEmitida']
        factRect = datFactRec['FacturasRectificadas']
        tipoDesg = datFactRec['TipoDesglose']
        desgFact = tipoDesg['DesgloseFactura']
        desgTipoOp = tipoDesg['DesgloseTipoOperacion']
        if desgTipoOp:
            entrega = desgTipoOp['Entrega']
            logger.info("Entrega ... : " + str(entrega) )
            if entrega:
                tieneEntrega = 1
                sujeta = entrega['Sujeta']
                if sujeta:
                    exenta = sujeta['Exenta']
                    noExenta = sujeta['NoExenta']
                noSujeta = entrega['NoSujeta']
				
                
        logger.info("tipoDesg: " + str(tipoDesg) )
        logger.info("desgFact: " + str(desgFact) )
		
        if desgFact:
            sujeta = desgFact['Sujeta']
            noSujeta = desgFact['NoSujeta']
            if sujeta:
                exenta = sujeta['Exenta']
                noExenta = sujeta['NoExenta']
            if noExenta:
                tipoNoExenta = noExenta['TipoNoExenta']
                desIVA = noExenta['DesgloseIVA']
        #detIVA = desIVA['DetalleIVA']
        if datFactRec:
            contraparte = datFactRec['Contraparte']
        datosPres = r['DatosPresentacion']
        estadoFact = r['EstadoFactura']
        datosDesContraparte = r['DatosDescuadreContraparte']
        logger.info("Factura: " + str(idFact) + " Detalle: " + str(noExenta) ) #+ " Tipo no exenta: " + str(tipoNoExenta) ) 


        #Vamos escribiendo en Excel
        logger.info("Procesando factura ... " + idFact['NumSerieFacturaEmisor'] )
        nHoja.write(nFila, 0, ejercicio)
        nHoja.write(nFila, 1, periodo)
        nHoja.write(nFila, 2, contraparte['NIF'] )#idEmiFact['NIF'])
        nHoja.write(nFila, 3, idFact['NumSerieFacturaEmisor'])
        nHoja.write(nFila, 4, idFact['FechaExpedicionFacturaEmisor'])
        nHoja.write(nFila, 5, datFactRec['TipoFactura'])
        nHoja.write(nFila, 6, datFactRec['TipoRectificativa'])
        #Si es una factura Rectificativa
        if factRect:
            IDFactRect = factRect['IDFacturaRectificada']
            #Aqui tenemos que tener en cuenta que puede ser mas de una
            nHoja.write(nFila, 7, IDFactRect[0]['NumSerieFacturaEmisor'])
            nHoja.write(nFila, 8, IDFactRect[0]['FechaExpedicionFacturaEmisor'])
        nHoja.write(nFila, 9, datFactRec['ClaveRegimenEspecialOTrascendencia'])
        nHoja.write(nFila, 10, datFactRec['ClaveRegimenEspecialOTrascendenciaAdicional1'])
        nHoja.write(nFila, 11, datFactRec['ClaveRegimenEspecialOTrascendenciaAdicional2'])
        nHoja.write(nFila, 12, datFactRec['DescripcionOperacion'])
        #InversionSujetoPasivo o DesgloseIVA
        logger.info("Exenta: " + str(exenta) )
        #falta controlar mas de un iva
        if exenta:
            nHoja.write(nFila, 13, "Sujeta")
            #detIVA = invSujPas['DetalleIVA']
            if desIVA:
                detIVA = desIVA['DetalleIVA']
            base = exenta['BaseImponible']
            if detIVA: #Detalle de IVA
                for i in range(len(detIVA)):
                    nHoja.write(nFila, 14, 0 ) #detIVA[i]['TipoImpositivo']
                    nHoja.write(nFila, 15, float(base) ) #detIVA[i]['BaseImponible'])
                    nHoja.write(nFila, 16, 0 ) #detIVA[i]['CuotaRepercutida'])
                    if(len(detIVA)-1) > i:
                        nFila = nFila + 1
            else:
                nHoja.write(nFila, 14, 0 )
                nHoja.write(nFila, 15, float(base) )
                nHoja.write(nFila, 16, 0 )
        else:
            nHoja.write(nFila, 13, "No Sujeta")
            detIVA = desIVA['DetalleIVA']
            for i in range(len(detIVA)):
                nHoja.write(nFila, 14, float(detIVA[i]['TipoImpositivo']) )
                nHoja.write(nFila, 15, float(detIVA[i]['BaseImponible']) )
                nHoja.write(nFila, 16, float(detIVA[i]['CuotaRepercutida']) )
                if(len(detIVA)-1) > i:
                    nFila = nFila + 1
        nHoja.write(nFila, 17, contraparte['NombreRazon'])
        nHoja.write(nFila, 18, datFactRec['FechaOperacion'])
        if noExenta:
            nHoja.write(nFila, 19, noExenta['TipoNoExenta'])
        else:
            nHoja.write(nFila, 19, exenta['CausaExencion'])
        nHoja.write(nFila, 20, datosPres['NIFPresentador'])
        nHoja.write(nFila, 21, datosPres['TimestampPresentacion'])
        nHoja.write(nFila, 22, datosPres['CSV'])
        nHoja.write(nFila, 23, estadoFact['EstadoCuadre'])
        nHoja.write(nFila, 24, estadoFact['TimestampEstadoCuadre'])
        nHoja.write(nFila, 25, estadoFact['TimestampUltimaModificacion'])
        nHoja.write(nFila, 26, estadoFact['EstadoRegistro'])
        nHoja.write(nFila, 27, estadoFact['CodigoErrorRegistro'])
        nHoja.write(nFila, 28, estadoFact['DescripcionErrorRegistro'])
        nHoja.write(nFila, 29, datosDesContraparte)
        nFila = nFila + 1

nExcel.save(cadArchivo)

archExcel = open(cadArchivo,'rb')
#adjunto = MIMEBase('multipart', 'encrypted')
adjunto = MIMEBase('application', "octet-stream")
adjunto.set_payload(archExcel.read())
archExcel.close()
encoders.encode_base64(adjunto)

adjunto.add_header('Content-Disposition', 'attachment', filename = cadTiempo + str(mes).zfill(2) + "_factEmitidasPresentadas.xls")

mail.attach(adjunto)
servidor = smtplib.SMTP(servidorSMTP, 587)
servidor.starttls()
servidor.ehlo()
servidor.login(usuarioSMTP, passSMTP)
#enviamos el Email
servidor.sendmail(emailEnvio, emailRecepcion, mail.as_string())
logger.info ("Correo enviado ...")
servidor.quit()


