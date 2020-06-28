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
emailRecepcion = ['email@envio.com']
servidorSMTP = 'smtp.envio.com'
usuarioSMTP = emailEnvio
passSMTP = 'passEmail'


mes = raw_input("""Numero mes(año 2017): """)
anio = 2017

mail = MIMEMultipart()
mail['To'] = ", ".join(emailRecepcion)
mail['From'] = emailEnvio
mail['Subject'] = '09.SII. Consulta Operaciones Intracomunitarias'
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

cadTiempo = time.strftime("%Y") + time.strftime("%m") + time.strftime("%d") + time.strftime("%H") + time.strftime("%M") + time.strftime("%S") + "_"
cadArchivo = rutaInformes + cadTiempo + "_"+  str(mes).zfill(2) +  "_OperacionesIntracomunitarias.xls"

logger = logging.getLogger('zeep')
logger.setLevel(logging.DEBUG)
fh = logging.FileHandler(rutaLog + cadTiempo + '_consulta_OI.log')
fh.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
logger.addHandler(fh)




#mes = 7
#mes = raw_input("Numero mes: ")


logger.info( 'Inicio de la consulta. Ano: ' + str(anio) + " | Mes: " + str(mes).zfill(2) )

##Cadenas ...
nExcel = xlwt.Workbook()
nHoja = nExcel.add_sheet(cadTiempo)

#Insertamos las cabeceras
nHoja.write(0,0, "Facturas Recibidas Presentadas: ")
nHoja.write(1,0, "Ejercicio")
nHoja.write(1,1, "Periodo")
nHoja.write(1,2, "NombreRazonEmisor")
nHoja.write(1,3, "NifEmisor")
nHoja.write(1,4, "ID")
nHoja.write(1,5, "NumSerieFE")
nHoja.write(1,6, "FechaExpedicionFE")
nHoja.write(1,7, "TipoOperacion")
nHoja.write(1,8, "ClaveDeclarado")
nHoja.write(1,9, "EstadoMiembro")
nHoja.write(1,10, "PlazoOperacion")
nHoja.write(1,11, "DescripcionBienes")
nHoja.write(1,12, "DireccionOperador")
nHoja.write(1,13, "FacturasODocumentacion")
nHoja.write(1,14, "NombreRazonContraparte")
nHoja.write(1,15, "Nif")
nHoja.write(1,16, "TimeStampPresentacion")
nHoja.write(1,17, "CSV")
nHoja.write(1,18, "EstadoRegistro")
nHoja.write(1,19, "TimestampUltimaModificacion")
nHoja.write(1,20, "CodigoErrorRegistro")
nHoja.write(1,21, "DescripcionErrorRegistro")
# nHoja.write(1,20, "")
# nHoja.write(1,21, "")
# nHoja.write(1,22, "")
# nHoja.write(1,23, "EstadoCuadre")
# nHoja.write(1,24, "TimeStampEstadoCuadre")
# nHoja.write(1,25, "TimeStampUltimaModificacion")
# nHoja.write(1,26, "EstadoRegistro")
# nHoja.write(1,27, "CodigoErrorRegistro")
# nHoja.write(1,28, "DescripcionErrorRegistro")
# nHoja.write(1,29, "DatosDescuadreContraparte")




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

wsdl = rutaDatos + 'SuministroOpIntracomunitarias.wsdl'
#wsdl = "./WSDLPruebas/SuministroFactEmitidasPrueba.wsdl"

#SSL CONECTION
session = Session()
session.cert = (rutaCertificados + "clavePublica.pem", rutaCertificados + "clavePrivada.pem")
session.verify = True

transport = Transport(session=session)
#logger.debug()
#FIRMA XML ENVIO
#SuministroFactRecibidas
client = Client(wsdl=wsdl,port_name="SuministroOpIntracomunitarias",transport=transport, service_name = 'siiService')
logger.info('Establecida conexion con cliente')
#SELECCION SERVICIO DE PRUEBAS        SuministroLRFacturasEmitidas
#SuministroLRFacturasRecibidas
#SuministroLRFacturasRecibidas
service2 = client.bind('siiService', 'SuministroOpIntracomunitarias')
logger.info('Bind al servicio')
#INSERCION DE LA FACTURA Y RESPUESTA EN EL SERVICO DE PRUEBAS.
salida = (service2.ConsultaLRDetOperacionIntracomunitaria(dicCabecera, dicFactura ))


logger.info("--> Salida: --> '" + str(salida) )
logger.info('Consulta finalizada.')
#print salida


nFila = 2
nCol = 0


#Comprobamos si hay salida
if salida['RegistroRespuestaConsultaLRDetOperIntracomunitarias']:
    dPeriodo = salida['PeriodoImpositivo']
    sinSalida = 0 #hay salida
else:
    logger.info("La peticion no ha devuelto datos")
    sys.exit(0)


if sinSalida == 0: #Procesamos la info.
    ejercicio = dPeriodo['Ejercicio']
    periodo = dPeriodo['Periodo']
    resultado = salida['RegistroRespuestaConsultaLRDetOperIntracomunitarias']
    logger.info('Resultado: ' + str(resultado) )

    for r in resultado:
        idFact = r['IDFactura']
        idEmiFact = idFact['IDEmisorFactura']
        if idEmiFact:
            idOtro = idEmiFact['IDOtro']
        datDetOperIntracomunitarias = r['DatosDetOperIntracomunitarias']
        if datDetOperIntracomunitarias:
            contraparte = datDetOperIntracomunitarias['Contraparte']
            detOperIntra = datDetOperIntracomunitarias['DetOperIntracomunitarias']

        # datFactRec = r['DatosFacturaRecibida']
        # factRect = datFactRec['FacturasRectificadas']
        # desgFact = datFactRec['DesgloseFactura']
        # invSujPas = desgFact['InversionSujetoPasivo']
        # desIVA = desgFact['DesgloseIVA']
        # contraparte = datFactRec['Contraparte']

        datosPres = r['DatosPresentacion']
        estadoFact = r['EstadoFactura']



        #Vamos escribiendo en Excel
        nHoja.write(nFila, 0, ejercicio)
        nHoja.write(nFila, 1, periodo)
        nHoja.write(nFila, 2, idEmiFact['NombreRazon'])
        nHoja.write(nFila, 3, idEmiFact['NIF'])
        if idOtro:
            nHoja.write(nFila, 4, idOtro['ID'])
        nHoja.write(nFila, 5, idFact['NumSerieFacturaEmisor'])
        nHoja.write(nFila, 6, idFact['FechaExpedicionFacturaEmisor'])
        nHoja.write(nFila, 7, detOperIntra['TipoOperacion'])
        nHoja.write(nFila, 8, detOperIntra['ClaveDeclarado'])
        nHoja.write(nFila, 9, detOperIntra['EstadoMiembro'])
        nHoja.write(nFila, 10, detOperIntra['PlazoOperacion'])
        nHoja.write(nFila, 11, detOperIntra['DescripcionBienes'])
        nHoja.write(nFila, 12, detOperIntra['DireccionOperador'])
        nHoja.write(nFila, 13, detOperIntra['FacturasODocumentacion'])

        nHoja.write(nFila, 14, contraparte['NombreRazon'])
        nHoja.write(nFila, 15, datosPres['NIFPresentador'])
        nHoja.write(nFila, 16, datosPres['TimestampPresentacion'])
        nHoja.write(nFila, 17, datosPres['CSV'])

        nHoja.write(nFila, 18, estadoFact['EstadoRegistro'])
        nHoja.write(nFila, 19, estadoFact['TimestampUltimaModificacion'])
        nHoja.write(nFila, 20, estadoFact['CodigoErrorRegistro'])
        nHoja.write(nFila, 21, estadoFact['DescripcionErrorRegistro'])
        logger.info("Escribiendo linea " + str(nFila))
        nFila = nFila + 1

nExcel.save(cadArchivo)

archExcel = open(cadArchivo,'rb')
#adjunto = MIMEBase('multipart', 'encrypted')
adjunto = MIMEBase('application', "octet-stream")
adjunto.set_payload(archExcel.read())
archExcel.close()
encoders.encode_base64(adjunto)

adjunto.add_header('Content-Disposition', 'attachment', filename = cadTiempo + "_" +  str(mes).zfill(2) + "_OperacionesIntracomunitarias.xls")

mail.attach(adjunto)
servidor = smtplib.SMTP(servidorSMTP, 587)
servidor.starttls()
servidor.ehlo()
servidor.login(usuarioSMTP, passSMTP)
#enviamos el Email
servidor.sendmail(emailEnvio, emailRecepcion, mail.as_string())
logger.info ("Correo enviado ...")
servidor.quit()



