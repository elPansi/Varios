# -*- coding: utf-8 -*-

import zeep
import requests
import time, sys, datetime
import logging
import ast
from libBD import consultaMSSQL
from requests import Session
from zeep import Client
from zeep.wsse.signature import Signature
from zeep.exceptions import Fault
from zeep.transports import Transport
from zeep.helpers import serialize_object



def enviaFacturaCompra(serieFactura, numeroFactura):

    reload(sys)
    sys.setdefaultencoding('utf8')

    rutaLog = '.\\Log\\'
    rutaCertificados = '.\\certificados\\'
    rutaDatos = '.\\datos\\'
    cuotaDeducible = 0

    # Configuracion de las series de Facturas y series de Facturas rectificativas
    serieFacturas = '18'
    serieFacturasRectificativas = 'R18'
    exentaIVA = 0

    ##Cadenas ...
    # SQL DETALLE FACTURA
    sqlRecuperaFactura = """SELECT
    'EMPRESA' NOMBRERAZON, 'BXXXXXX' NIF, 'A0' TIPOCOMUNICACION,
    YEAR(FC.FECHAFACTURA) EJERCICIO, MONTH(fc.FECHAFACTURA) PERIODO, PROV.cif IDEMISORFACTURA,
    ( FC.NumDocumento ) NUMSERIEFACTURAEMISOR,
    CONVERT(NVARCHAR(10),FC.FECHAFACTURA,103) FECHAEXPEDICIONFACTURAEMISOR,
    (SELECT CASE
    WHEN FS.Serie = '{SERIEFACTURAS}' THEN 'F1' --SERIE FACTURAS NORMALES
    WHEN FS.SERIE = '{SERIERECTIFICADAS}' THEN 'R1' --SERIE FACTURAS RECTIFICATIVAS
    END ) TIPOFACTURA ,
    --FACTURAS AGRUPADAS
    (SELECT CASE
    WHEN FS.Serie = '{SERIEFACTURAS}' THEN ( FS.SERIE + ' - ' + STR(FC.NumDocPropio) + ' | ' + FC.NumDocumento )
    WHEN FS.Serie = '{SERIERECTIFICADAS}' THEN NULL
    END ) NUMSERIEFACTURAEMISOR,
    (SELECT CASE
    WHEN FS.Serie = '{SERIEFACTURAS}' THEN ( CONVERT(NVARCHAR(10),FC.FechaDocProv, 103) )
    WHEN FS.Serie = '{SERIERECTIFICADAS}' THEN NULL
    END ) FECHAEXPEDICIONFACTURAEMISOR,
    --FACTURAS RECTIFICADAS
    (SELECT CASE
    WHEN FS.Serie = '{SERIEFACTURAS}' THEN  NULL
    WHEN FS.Serie = '{SERIERECTIFICADAS}' THEN ( FS.SERIE + ' - ' + STR(FC.NumDocPropio) + ' | ' + FC.NumDocumento )
    END ) NUMSERIEFACTURAEMISOR,
    (SELECT CASE
    WHEN FS.Serie = '{SERIEFACTURAS}' THEN NULL
    WHEN FS.Serie = '{SERIERECTIFICADAS}' THEN ( CONVERT(NVARCHAR(10),FC.FechaDocProv, 103) )
    END ) FECHAEXPEDICIONFACTURAEMISOR,
    --BASE RECTIFICADA
    (SELECT CASE
    WHEN FS.SERIE = '{SERIERECTIFICADAS}' THEN SUM( FL.Cantidad * FL.Precio*(1-FL.Dto1/100) )
    ELSE 0
    END ) BASERECTIFICADA,
    --CUOTA RECTIFICADA
    (SELECT CASE
    WHEN FS.SERIE = '{SERIERECTIFICADAS}' THEN SUM(FL.Cantidad * FL.Precio * (1-FL.Dto1/100)* FL.iva / 100)
    ELSE 0
    END ) CUOTARECTIFICADA,
    (SELECT CASE
    WHEN ( FC.CodProveedor IN ( 581 ) ) THEN '14'              -- SEGUROS
    WHEN ( PROV.CODPROVINCIA IN ( 35, 38, 51, 52 ) ) THEN '09' -- IGIC/IPSI
    ELSE '01'                                                  -- OPERACIONES REGIMEN COMUN
    END ) CLAVEREGIMENESPECIALOTRASCENDENCIA,
    ('FACTURA DE COMPRA ' + FS.SERIE + ' / ' + FC.NumDocumento) DESCRIPCIONOPERACION,
    1 SITUACIONINMUEBLE, CONVERT(NVARCHAR(40), PROV.NOMBRE) NOMBRERAZONCLIENTE, PROV.CIF NIFCLIENTE,
    '02' TIPOIDENTIFICACIONPAISRESIDENCIA, PROV.CIF NUMEROIDENTIFICACION,
    CONVERT(NVARCHAR(10),FC.FECHAFACTURA,103) FECHACONTABLE,
    FC.OBSERVACIONES OBSERVACIONES
    FROM FACTURACABCOMPRA FC FULL OUTER JOIN ArticulosDetallesCompra FL ON FC.CodFacturaCompra = FL.CodFacturaCompra
    INNER JOIN PROVEEDORES PROV ON PROV.CODPROVEEDOR = FC.CODPROVEEDOR
    FULL OUTER JOIN PROVINCIAS P ON P.codigo = PROV.CodProvincia
    INNER JOIN FACTURASERIE FS ON FS.CODSERIE = FC.SERIE
    WHERE FS.SERIE = '{SERIEFACTURA}' AND FC.NumDocPropio = {NUMEROFACTURA}
    GROUP BY FC.CodFacturaCompra, FS.SERIE, FC.NumDocumento, FC.NumDocPropio, FC.CodProveedor, PROV.CIF, FC.FechaDocProv, PROV.CodProvincia, FC.retencion,
    PROV.Nombre, FC.FechaFactura, FC.OBSERVACIONES, fc.FechaContable
    ORDER BY FS.Serie, FC.NumDocPropio"""

    # CLAVEREGIMENESPECIAL O TRASCENDENCIA
    # SEGUROS                          -- VALOR: 15 -- LOCALIZAR POR EL CODIGO DEL PROVEEDOR
    # OPERACIONES SUJETAS AL IPSI/IGIC -- VALOR: 09 -- LOCALIZAR POR EL CODIGO DE LA PROVINCIA ()
    # RESTO                            -- VALOR: 01 -- RESTO

    # SQL DETALLE DE IMPUESTOS (esta en principio no la voy a usar)
    sqlRecuperaImpuestos = """SELECT
    FC.CODFACTURA, FC.SERIE, FC.NumFactura, FL.IVA,
    SUM(FL.CANTIDAD*FL.Precio*(1-Fl.Dto1/100)) BASEIMPONIBLE,
    SUM(FL.CANTIDAD*FL.PRECIO*(1-FL.DTO1/100)*FL.IVA/100) CUOTAIVA
    FROM FACTURACAB FC FULL OUTER JOIN articulosdetalles FL ON FC.CODFACTURA = FL.CODFACTURA
    WHERE FC.Serie = {SERIEFACTURA} AND FC.NumFactura = {NUMEROFACTURA}
    GROUP BY FC.CODFACTURA, FC.SERIE, FC.NumFactura, FL.IVA"""

    sqlRecuperaImpuesto = """SELECT
    REG_FAC.codfacturagestion CODFACTURACOMPRA, FAC_COM.Serie , FAC_COM.NumDocPropio,
    REG_FAC.IVA, SUM(REG_FAC.BASE) BASE_CONTA,
    SUM(REG_FAC.CUOTA) CUOTA_CONTA
    FROM ( SELECT RF.codfacturagestion, RF.Baseimponible1 BASE, RF.iva1 IVA, RF.cuotaiva1 CUOTA
    FROM cVillaescusa.DBO.RegistroEntradaFacturas RF
    UNION
    SELECT RF.codfacturagestion, RF.Baseimponible2 BASE, RF.iva2 IVA, RF.cuotaiva2 CUOTA
    FROM cVillaescusa.DBO.RegistroEntradaFacturas RF
    UNION
    SELECT RF.codfacturagestion, RF.Baseimponible3 BASE, RF.iva3 IVA, RF.cuotaiva3 CUOTA
    FROM cVillaescusa.DBO.RegistroEntradaFacturas RF
    UNION
    SELECT RF.codfacturagestion, RF.Baseimponible4 BASE, RF.iva4 IVA, RF.cuotaiva4 CUOTA
    FROM cVillaescusa.DBO.RegistroEntradaFacturas RF
    UNION
    SELECT RF.codfacturagestion, RF.Baseimponible5 BASE, RF.iva5 IVA, RF.cuotaiva5 CUOTA
    FROM cVillaescusa.DBO.RegistroEntradaFacturas RF
    UNION
    SELECT RF.codfacturagestion, RF.Baseimponible6 BASE, RF.iva6 IVA, RF.cuotaiva6 CUOTA
    FROM cVillaescusa.DBO.RegistroEntradaFacturas RF ) REG_FAC  INNER JOIN (SELECT
    FS.SERIE, FC.CodFacturaCompra, FC.NumDocPropio, SUM(FL.CANTIDAD*FL.Precio*(1-Fl.Dto1/100)*(1-Fl.Dto2/100)) BASE, FC.FechaFactura,
    SUM(FL.CANTIDAD*FL.PRECIO*(1-FL.DTO1/100)*(1-Fl.Dto2/100)*FL.IVA/100) CUOTAIVA, FL.iva
    FROM gvilla.dbo.FACTURACABCOMPRA FC INNER JOIN gvilla.dbo.articulosdetallesCompra FL ON FC.CODFACTURACOMPRA = FL.CODFACTURACOMPRA
    INNER JOIN gvilla.dbo.FACTURASERIE FS ON FS.CODSERIE = FC.SERIE
    GROUP BY FS.Serie, FC.CodFacturaCompra, FC.NUMDOCPROPIO, FL.IVA, FC.FechaFactura ) FAC_COM ON FAC_COM.CodFacturaCompra = REG_FAC.codfacturagestion AND FAC_COM.iva = REG_FAC.IVA
    WHERE FAC_COM.Serie = '{SERIEFACTURA}' AND FAC_COM.NumDocPropio = {NUMEROFACTURA} AND FAC_COM.iva = {IVA}
    GROUP BY FAC_COM.SERIE, FAC_COM.CODFACTURACOMPRA, FAC_COM.NumDocPropio, REG_FAC.codfacturagestion, REG_FAC.IVA, FAC_COM.BASE, FAC_COM.CUOTAIVA
    ORDER BY FAC_COM.Serie, FAC_COM.NumDocPropio, FAC_COM.BASE"""

    sqlDistintosImpuestos = """SELECT SUBCONSULTA.iva
    FROM  (SELECT CodFacturaCompra, IVA, Cantidad FROM articulosdetallescompra
    WHERE CodFacturacompra = (SELECT CODFACTURACOMPRA FROM FacturaCabCompra WHERE SERIE =
    (SELECT CODSERIE FROM FACTURASERIE WHERE SERIE = '{SERIEFACTURA}' ) AND NumDocPropio = {NUMEROFACTURA} )
    GROUP BY CodfacturaCompra, iva, CANTIDAD) SUBCONSULTA
    WHERE SUBCONSULTA.iva * SUBCONSULTA.Cantidad <> 0
    GROUP BY iva"""

    sqlIvaAlCero = """SELECT SUBCONSULTA.iva, SUBCONSULTA.BASEIMP
    FROM  (SELECT CodFacturaCompra, IVA, Cantidad, SUM(CANTIDAD*PRECIO*(1-DTO1/100)) AS BASEIMP
    FROM articulosdetallescompra
    WHERE CodFacturacompra = (SELECT CODFACTURACOMPRA FROM FacturaCabCompra WHERE SERIE =
    (SELECT CODSERIE FROM FACTURASERIE WHERE SERIE = '{serie}' ) AND NumDocPropio = {numDocPropio} )
    GROUP BY CodfacturaCompra, iva, CANTIDAD, PRECIO, DTO1) SUBCONSULTA
    WHERE SUBCONSULTA.IVA = 0
    GROUP BY SUBCONSULTA.iva, SUBCONSULTA.BASEIMP"""

    sqlExencion = """SELECT TOP(1) SUBSTRING(IDDEARTICULO,1,2) EXENCION
    FROM ARTICULOS WHERE CODARTICULO IN(
    SELECT CODARTICULO FROM ARTICULOSDETALLESCOMPRA
    WHERE CodFacturaCompra = (SELECT CODFACTURACompra FROM FacturaCabCompra WHERE SERIE =
    (SELECT CODSERIE FROM FACTURASERIE WHERE SERIE ='{SERIEFACTURA}') AND NumDocPropio = {NUMEROFACTURA})
    ) AND IDDeArticulo LIKE 'E%'  AND CodFamilia = 34 AND BloqueoVentas = 0"""

    sqlBaseImponible = """SELECT FC.CODFACTURACOMPRA, FS.SERIE, FC.NumDocPropio, FL.IVA,
    SUM(FL.CANTIDAD*FL.Precio*(1-Fl.Dto1/100)) BASEIMPONIBLE,
    SUM(FL.CANTIDAD*FL.PRECIO*(1-FL.DTO1/100)*FL.IVA/100) CUOTAIVA
    FROM FACTURACABCOMPRA FC FULL OUTER JOIN articulosdetallesCompra FL ON FC.CODFACTURACOMPRA = FL.CODFACTURACOMPRA
    INNER JOIN FACTURASERIE FS ON FS.CODSERIE = FC.SERIE
    WHERE FS.Serie = '{SERIEFACTURA}' AND FC.NumDocPropio = {NUMEROFACTURA} AND FL.IVA = 0
    GROUP BY FC.CODFACTURACOMPRA, FS.SERIE, FC.NumDocPropio, FL.IVA"""

    sqlBaseRectificativas = """SELECT FC.CODFACTURACOMPRA, FS.SERIE, FC.NumDocPropio,
    SUM(FL.CANTIDAD*FL.Precio*(1-Fl.Dto1/100)) BASEIMPONIBLE,
    SUM(FL.CANTIDAD*FL.PRECIO*(1-FL.DTO1/100)*FL.IVA/100) CUOTAIVA
    FROM FACTURACABCOMPRA FC FULL OUTER JOIN articulosdetallesCompra FL ON FC.CODFACTURACOMPRA = FL.CODFACTURACOMPRA
    INNER JOIN FACTURASERIE FS ON FS.CODSERIE = FC.SERIE
    WHERE FS.Serie = '{SERIEFACTURA}' AND FC.NumDocPropio = {NUMEROFACTURA}
    GROUP BY FC.CODFACTURACOMPRA, FS.SERIE, FC.NumDocPropio"""

    sqlFacturaARectificar = """SELECT FCC.OBSERVACIONES
    FROM FacturaCabCompra FCC INNER JOIN FacturaSerie FS ON FCC.SERIE= FS.CodSerie
    WHERE FS.SERIE = '{SERIE}' AND FCC.NumDocPropio = {NUMERO}
    AND UPPER(Observaciones) LIKE '%RECTIFICA:%'"""

    sqlGuardaRespuesta = """INSERT INTO [dbo].[GV_SII]
               ([CSV]
               ,[TIPOFAC]
               ,[SERIE]
               ,[NUMERO]
               ,[ESTADOENVIO]
               ,[NUMSERIEFACTEMISOR]
               ,[ESTADOREGISTRO]
               ,[FECHAPRESENTACION]
               ,[CODERRORREGISTRO]
               ,[DESCRIPERRORREGISTRO]
    		   ,[CSV2])
         VALUES
              ('{CSV}', '{TIPOFAC}', '{SERIE}', {NUMERO}, '{ESTADOENVIO}',
               '{NUMSERIEFACTEMISOR}', '{ESTADOREGISTRO}', '{FECHAPRESENTACION}',
               '{CODERRORREGISTRO}', '{DESCRIPERRORREGISTRO}', '{CSV2}')"""

    # Primero comprobamos que la factura este contabilizada
    sqlContabilizada = """SELECT Contabilizado FROM FACTURACABCOMPRA FC INNER JOIN FacturaSerie FS ON FS.CodSerie = FC.SERIE
    WHERE FS.SERIE = '{SERIEFACTURA}' AND FC.NumDocPropio = {NUMEROFACTURA}"""

    # Preparamos la cadena del diccionario
    cadCabecera = """{
    	"IDVersionSii":1.0,
    	"Titular":{
    		"NombreRazon":"EMPRESA",
    		"NIF":"BXXXXX"
    		},
    	"TipoComunicacion":"A0"}""" #A0 - NUEVAS || A1 - MODIFICACION

    # FACTURAS
    cadFactura = """{
        "PeriodoImpositivo":{
    		"Ejercicio":"{anio}",
    		"Periodo":"{mes}"
    		},
    	"IDFactura":{
    		"IDEmisorFactura":{
    		    "NIF":"{IDEMISORFACTURA}",
    		    "IDOtro":{
    			     "IDType":"02",
    				 "ID":"ES"
    				 }
    		    },
    		"NumSerieFacturaEmisor":"{numSerieFactEmisor}",
    		"FechaExpedicionFacturaEmisor":"{fechaExpFactEm}"
    		},
    	"FacturaRecibida":{
    		"TipoFactura":"{tipoFactura}",
    		{CADRECTIFICADAS}
    		"ClaveRegimenEspecialOTrascendencia":"{REGESPECIAL}",
    		"DescripcionOperacion":"{descripOperacion}",
            {CADDESGLOSEFACTURA},
    		"Contraparte":{
    			"NombreRazon":"{nombreCliente}",
    			"NIF":"{nifCliente}",
    			"IDOtro":{
    			     "IDType":"02",
    				 "ID":"ES"
    				 }
            },
    		"FechaRegContable":"{fechaRegContable}",
    		"CuotaDeducible":"{cuotaDeducible}"
    	}
    }"""

    # cadena para facturas rectificativas por sustitucion
    cadRectificadas = """
          "TipoRectificativa":"{TIPORECT}",
            "FacturasRectificadas":{
    		    "IDFacturaRectificada":{
    			    "NumSerieFacturaEmisor":"{numSerieFactEmisor}",
    				"FechaExpedicionFacturaEmisor":"{fechaExpFactEm}"
    				}
    			},
    		"ImporteRectificacion":{
    		    "BaseRectificada":"{BASERECTIFICADA}",
    			"CuotaRectificada":"{CUOTARECTIFICADA}",
    			"CuotaRecargoRectificado": 0
    			},"""

    # cadena para facturas rectificativas por diferencias
    cadRectificadasDif = """      "TipoRectificativa":"{TIPORECT}",
            "FacturasRectificadas":{
    		    "IDFacturaRectificada":{
    			    "NumSerieFacturaEmisor":"{numSerieFactEmisor}",
    				"FechaExpedicionFacturaEmisor":"{fechaExpFactEm}"
    				}
    			},"""

    cadDesgloseFactura = """  "DesgloseFactura":{
                                 "DesgloseIVA":{
                                    "DetalleIVA": [
                                          {DESGLOSESIVAS}
                                     ]
                                 }
                        }"""

    cadDesgloseIVA = """  	{
    							"TipoImpositivo":{porcenIVA},
    							"BaseImponible":{baseImp},
    							"CuotaSoportada":{cuotaIVA}
    						}"""  # concatenar con una coma para distintos IVAS

    cadInvSujPasivo = """ "DesgloseFactura":{
                            "InversionSujetoPasivo":{
    							"DetalleIVA":{
    							"TipoImpositivo":{porcenIVA},
    							"BaseImponible":{baseImp},
    							"CuotaSoportada":{cuotaIVA}
    							}
    						}
                         }"""  # concatenar con una coma para distintos IVAS

    
    cadTiempo = time.strftime("%Y") + time.strftime("%m") + time.strftime("%d") + time.strftime("%H") + time.strftime("%M") + time.strftime("%S") + "_" + str(serieFactura) + "_" + str(numeroFactura) + "_"

    logger = logging.getLogger('zeep')
    logger.setLevel(logging.DEBUG)
    fh = logging.FileHandler(rutaLog + cadTiempo + '_compra_logSii.log')
    fh.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    fh.setFormatter(formatter)
    logger.addHandler(fh)
    logger.info( 'Iniciando el envio de la factura: ' + str(serieFactura) + ' / ' + str(numeroFactura)  )

    consContabilizada = sqlContabilizada.replace("{SERIEFACTURA}", str(serieFacturas) )
    consContabilizada = consContabilizada.replace("{NUMEROFACTURA}", str(numeroFactura) )

    conta = consultaMSSQL(consContabilizada)

    print str(conta)
    #print conta[0][0]
    if conta[0][0] == True:
        logger.info("Factura Contabilizada ...: " + str(serieFacturas) + " / " + str(numeroFactura) )
    else:
        logger.info("Error, factura sin contabilizar. No se enviará la factura: "  + str(serieFacturas) + " / " + str(numeroFactura) )
        sys.exit(1)

    consDistintosImpuestos = sqlDistintosImpuestos.replace("{SERIEFACTURA}", str(serieFactura) )
    consDistintosImpuestos = consDistintosImpuestos.replace("{NUMEROFACTURA}", str(numeroFactura) )

    impuestosFactura = consultaMSSQL(consDistintosImpuestos)

    logger.info("Consulta distintos Impuestos: " + consDistintosImpuestos)


    consRecFact = sqlRecuperaFactura.replace("{SERIEFACTURAS}", str(serieFacturas) )
    consRecFact = consRecFact.replace("{SERIERECTIFICADAS}", str(serieFacturasRectificativas) )
    consRecFact = consRecFact.replace("{SERIEFACTURA}", str(serieFactura) )
    consRecFact = consRecFact.replace("{NUMEROFACTURA}", str(numeroFactura) )



    logger.info( "Impuestos --> " + str(impuestosFactura ) )
    logger.info( consRecFact )
    #logger.info( consRecImp )

    #Recuperamos la factura
    detalleFactura = consultaMSSQL(consRecFact)
    # detalleImpuestos = consultaMSSQL(consRecImp)

    logger.info("Detalle de la factura: " + str(detalleFactura) )


    #Cargamos las variables con la informacion
    ejercicio = detalleFactura[0][3]
    logger.info( "Ejercicio: " + str(ejercicio) )
    mes = detalleFactura[0][4]
    numSerieFactEmisor = detalleFactura[0][6]
    fechaExpFactEmisor = detalleFactura[0][7]
    tipoFactura = detalleFactura[0][8]
    regEspecial = detalleFactura[0][15]
    descripOperacion = detalleFactura[0][16]
    nombreCliente = str(detalleFactura[0][18])
    nifCliente = detalleFactura[0][19]
    fechaContable = detalleFactura[0][22]
    anulacion = detalleFactura[0][23]

    logger.info("Regimen Especial o Trascendencia: " + str(regEspecial) + " | Tipo Factura: " + str(tipoFactura) )
    #En caso de que sea rectificativa construimos la cadena.
    if tipoFactura[0] == 'R':
        logger.info("Factura Rectificativa ......................")
        consBaseRectif = sqlBaseRectificativas.replace("{SERIEFACTURA}", str(serieFactura))
        consBaseRectif = consBaseRectif.replace("{NUMEROFACTURA}", str(numeroFactura) )
        respRectif = consultaMSSQL( consBaseRectif )
        baseRectif = respRectif[0][3]
        cuotaRectif = respRectif[0][4]
        consFactARect = sqlFacturaARectificar.replace("{SERIE}", str(serieFactura) )
        consFactARect = consFactARect.replace("{NUMERO}", str(numeroFactura) )
        facturaARectificar = consultaMSSQL(consFactARect)
        print "Observaciones: " + str(facturaARectificar) + "\r\n"
        #aqui obtenemos el numero de la factura entre corchetes.
        if facturaARectificar:
            cadFact = facturaARectificar[0][0]
            print str(cadFact)
            pFacIni = cadFact.index("#")
            pFacFin = cadFact.index("#", pFacIni + 1 )
            pFechaIni = cadFact.index("#", pFacFin + 1 )
            pFechaFin = cadFact.index("#", pFechaIni + 1 )

            print ("Pos Ini: " + str(pFacIni) + " | Pos Fin: " + str(pFacFin) )
            numFactARectificar = cadFact[pFacIni+1:pFacFin]
            fechaFact = cadFact[pFechaIni+1:pFechaFin]
            #numFactARectificar = cadFact[cadFact.index("#")+1:cadFact.index("#",2)]
            print "Factura: " + str(numFactARectificar) + " | Fecha: " + fechaFact
        else:
            logger.info("Por favor debe introducir la informacion relativa a la factura que se modifica. Campo Observaciones.")

        logger.info("Base Rectificada: " + str(baseRectif) + " Cuota Rectificada: " + str(cuotaRectif) + " FechaFactEmisor: " + str(fechaExpFactEmisor))
        logger.info("Anulacion: " + anulacion.upper() )
        #Caso de sustitucion de factura

        if anulacion.upper().find("ANULACION") >= 0:
            cadFinRectificadas = cadRectificadas.replace("{TIPORECT}", "S")
            cadFinRectificadas = cadFinRectificadas.replace("{numSerieFactEmisor}", str(numFactARectificar))
            cadFinRectificadas = cadFinRectificadas.replace("{fechaExpFactEm}", str(fechaFact))
            cadFinRectificadas = cadFinRectificadas.replace("{BASERECTIFICADA}", '{:12.2f}'.format(baseRectif))
            cadFinRectificadas = cadFinRectificadas.replace("{CUOTARECTIFICADA}", '{:12.2f}'.format(cuotaRectif))
        else: #caso de diferencias
            cadFinRectificadas = cadRectificadasDif.replace("{TIPORECT}", "I")
            cadFinRectificadas = cadFinRectificadas.replace("{numSerieFactEmisor}", str(numFactARectificar))
            cadFinRectificadas = cadFinRectificadas.replace("{fechaExpFactEm}", str(fechaFact))

        logger.info("Cadena Rectificadas: " + str(cadFinRectificadas) )

        cadFactura = cadFactura.replace("{CADRECTIFICADAS}", str(cadFinRectificadas) )
    else:
        cadFactura = cadFactura.replace("{CADRECTIFICADAS}", "")

    logger.info ("Tipo de Factura:" + str(tipoFactura) )



    cadDesgloseImpuestos = ''
    cuotaIVATotal = 0
    cadDesglosesIVAs = ""
    logger.info("Comenzamos a tratar los ivas ... " + str(len(impuestosFactura)) + " | " )
    #Primero debemos cargar los Desgloses de IVA en caso de que exista
    if len(impuestosFactura) > 0: #aqui puede ser que lleve iva al 0
        for i,imp in enumerate(impuestosFactura):
            sqlDetImpuesto = sqlRecuperaImpuesto.replace("{SERIEFACTURA}", str(serieFactura) )
            sqlDetImpuesto = sqlDetImpuesto.replace("{NUMEROFACTURA}", str(numeroFactura) )
            #print imp[0]
            sqlDetImpuesto = sqlDetImpuesto.replace("{IVA}", str(imp[0]) )
            logger.info("Consulta Detalle impuesto: " + sqlDetImpuesto)
            respDetalle = consultaMSSQL(sqlDetImpuesto)
            iva = respDetalle[0][3]
            base = respDetalle[0][4]
            cuota = respDetalle[0][5]
            cadParDesglose = cadDesgloseIVA.replace("{porcenIVA}",'{:3.2f}'.format(iva) )
            cadParDesglose = cadParDesglose.replace("{baseImp}", '{:12.2f}'.format(base) )
            cadParDesglose = cadParDesglose.replace("{cuotaIVA}", '{:12.2f}'.format(cuota) )
            print cuota
            if i == 0:
                cadDesgloseImpuestos = cadParDesglose
                #print "Primer Impuesto: " + cadDesgloseImpuestos
            else:
                cadDesgloseImpuestos = cadDesgloseImpuestos + "," + cadParDesglose
                #print "Cadena final Impuestos" + cadDesgloseImpuestos
            cuotaIVATotal = cuotaIVATotal + cuota
        #aqui debemos comprobar si lleva también iva al 0%
        consIvaAlCero = sqlIvaAlCero.replace( "{serie}", str(serieFactura) )
        consIvaAlCero = consIvaAlCero.replace( "{numDocPropio}", str(numeroFactura) )
        respIvaCero = consultaMSSQL(consIvaAlCero)
        if respIvaCero:
            base = respIvaCero[0][1]
            cadParDesglose = cadDesgloseIVA.replace("{porcenIVA}",'{:3.2f}'.format(0) )
            cadParDesglose = cadParDesglose.replace("{baseImp}",'{:3.2f}'.format(base) )
            cadParDesglose = cadParDesglose.replace("{cuotaIVA}",'{:3.2f}'.format(0) )
            cadDesgloseImpuestos = cadDesgloseImpuestos + "," + cadParDesglose
	

    #Caso de que sea una factura sin IVA
    else:
        logger.info("EXENTA IVA ...")
        exentaIVA = 1
        consExencion = sqlExencion.replace("{SERIEFACTURA}", str(serieFactura) )
        consExencion = consExencion.replace("{NUMEROFACTURA}", str(numeroFactura) )
        # primero deberemos recuperar el tipo de exencion
        logger.info("Consulta Exencion: " + consExencion )
        exencion = consultaMSSQL(consExencion)
        consBaseImponible = sqlBaseImponible.replace("{SERIEFACTURA}", str(serieFactura) )
        consBaseImponible = consBaseImponible.replace("{NUMEROFACTURA}", str(numeroFactura ) )
        base = consultaMSSQL(consBaseImponible)
        logger.info( "Exencion: " + str(exencion) + " | " + "Base: " + str(base[0][4]) )
        #NO USAR cadExentaImpuestos --> usar cadInversionSujetoPasivo
        print "Exencion: " + str(exencion)
        if exencion: #Ha encontrado una E, Inversion de sujeto Pasivo
            print "Inversion de Sujeto Pasivo"
            baseImponible = base[0][4]
            cuotaIva = baseImponible * 0.21
            cuotaDeducible = cuotaDeducible + cuotaIva
            cadFinExenta = cadInvSujPasivo.replace("{porcenIVA}", "21.00")
            cadFinExenta = cadFinExenta.replace ("{baseImp}", '{:12.2f}'.format(base[0][4]) )
            cadFinExenta = cadFinExenta.replace("{cuotaIVA}", '{:12.2f}'.format(cuotaIva) )
        else: #Es una factura Normal --> Desglose IVA
            print "DesgloseIVA"
            cadFinExenta = cadDesgloseIVA.replace("{porcenIVA}", "0.00")
            cadFinExenta = cadFinExenta.replace ("{baseImp}", '{:12.2f}'.format(base[0][4]) )
            cadFinExenta = cadFinExenta.replace("{cuotaIVA}", "0.00")


    if len(impuestosFactura) >= 1:
        cadFinImpuestos = cadDesgloseImpuestos
    else:
        cadFinImpuestos = cadFinExenta

    #Ahora reemplazamos el desglose en los impuestos


    #cadFinImpuestos = cadFactura.replace("{CADDESGLOSEFACTURA}", cadDesgloseImpuestos )

    logger.info("FechaExpedicion: " + str(fechaExpFactEmisor))
    logger.info("FechaContable: " + str(fechaContable))

    fact = cadFactura.replace("{anio}", str(ejercicio) )
    fact = fact.replace("{mes}", str(mes).zfill(2) )
    fact = fact.replace("{numSerieFactEmisor}", str(numSerieFactEmisor) )
    fact = fact.replace("{fechaExpFactEm}", str(fechaExpFactEmisor) )
    fact = fact.replace("{tipoFactura}", str(tipoFactura) )
    fact = fact.replace("{descripOperacion}", str(descripOperacion) )
    fact = fact.replace("{nombreCliente}", str(nombreCliente) )
    fact = fact.replace("{nifCliente}", str(nifCliente) )
    fact = fact.replace("{IDEMISORFACTURA}", str(nifCliente) )
    fact = fact.replace("{REGESPECIAL}", str(regEspecial) )
    fact = fact.replace("{fechaRegContable}", str(fechaContable) )
    fact = fact.replace("{cuotaDeducible}", '{:12.2f}'.format(cuotaDeducible) )


    cadDesFact = cadDesgloseFactura.replace("{DESGLOSESIVAS}", str( cadFinImpuestos) ) #antes cadFinImpuestos

    logger.info("cadDesFact: " + str(cadDesFact) )
    logger.info("Factura Compra --> " + fact )

    #concatenamos la cadena de impuestos
    if exentaIVA == 0: #No Tiene exencion
        fact = fact.replace("{CADDESGLOSEFACTURA}", str(cadDesFact) )
    else:
        #comprobamos primero si es exenta.
        if exencion:
            fact = fact.replace("{CADDESGLOSEFACTURA}", str(cadFinExenta) )
        else:
            fact = fact.replace("{CADDESGLOSEFACTURA}", str(cadDesFact) )

    # fact = fact.replace("{porcenIVA}", '{:3.2f}'.format(porcenIVA) )
    # fact = fact.replace("{baseImp}", '{:12.2f}'.format(baseImp) )
    # fact = fact.replace("{cuotaIVA}", '{:12.2f}'.format(cuotaIVA) )



    logger.info(cadCabecera)
    logger.info(fact)

    #Convertimos cadenas a Diccionarios
    dicCabecera = ast.literal_eval(cadCabecera)
    dicFactura = ast.literal_eval(fact)


    print dicCabecera
    print '\r\n'
    print dicFactura


    #WEBSERVICE AEAT
    #wsdl = 'http://www.agenciatributaria.es/static_files/AEAT/Contenidos_Comunes/La_Agencia_Tributaria/Modelos_y_formularios/Suministro_inmediato_informacion/FicherosSuministros/V_07/SuministroFactEmitidas.wsdl'
    #wsdl = '/wlpl/SSII-FACT/ws/fe/SiiFactFEV1SOAP/SuministroFactEmitidasPrueba.wsdl'
    # Descargar los wsdl de la aeat y modificar la url del servicio por la del servicio de pruebas ya que si no te manda a XXXX
    ## ------------------------- OJO CREO QUE NO ESTAMOS UTILIZANDO EL SERVICIO DE PRUEBAS --------

    wsdl = rutaDatos + 'SuministroFactRecibidas.wsdl'
    #wsdl = "./WSDLPruebas/SuministroFactEmitidasPrueba.wsdl"

    #SSL CONECTION
    session = Session()
    #session.cert='cert.pem'
    session.cert = (rutaCertificados + "clavePublica.pem", rutaCertificados + "clavePrivada.pem")
    session.verify = True

    transport = Transport(session=session)
    #logger.debug()
    #FIRMA XML ENVIO
    #signature = Signature("./certificados/cert.pem","./certificados/clavePrivada.crt")
    client = Client(wsdl=wsdl,port_name="SuministroFactRecibidas",transport=transport, service_name = 'siiService')
    logger.info('Establecida conexion con cliente')
    #SELECCION SERVICIO DE PRUEBAS        SuministroLRFacturasEmitidas, SuministroFactRecibidas
    service2 = client.bind('siiService', 'SuministroFactRecibidas')
    logger.info('Bind al servicio')
    #INSERCION DE LA FACTURA Y RESPUESTA EN EL SERVICO DE PRUEBAS.
    #print (service2.SuministroLRFacturasEmitidas(cabecera,facturas))
    salida = (service2.SuministroLRFacturasRecibidas(dicCabecera,dicFactura))
    logger.info('Factura envidada')
    #print salida

    #INSERCION DE LA FACTURA Y RESPUESTA EN EL SERVICO.
    #print (client.service.SuministroLRFacturasEmitidas(cabecera,facturas))


    #Procesamos la salida
    csv = salida['CSV']
    datosPres = salida['DatosPresentacion']

    logger.info('Datos Presentacion: ' + str(datosPres) )

    if datosPres is not None:
        logger.info("Datos Presentacion: Datos ");
        fechaPresentacion = datosPres['TimestampPresentacion']
    else:
        fa = datetime.datetime.now()
        fechaPresentacion = "%s/%s/%s %s:%s:%s" % (fa.day, fa.month, fa.year, fa.hour, fa.minute, fa.second)

    estadoEnvio = salida['EstadoEnvio']
    respLinea = salida['RespuestaLinea'][0]
    logger.info("Respuesta Linea: " + str(respLinea) )

    idFactura = respLinea['IDFactura']

    if idFactura is not None:
        logger.info("ID Factura: Datos ");
        nSerFacEm = idFactura['NumSerieFacturaEmisor']

    if respLinea is not None:
        logger.info("Respuesta Linea: Datos ");
        estadoRegistro = respLinea['EstadoRegistro']
        codErrorRegistro = respLinea['CodigoErrorRegistro']
        descripErrorReg = respLinea['DescripcionErrorRegistro']
        csv2 = respLinea['CSV']

    #Preparamos la cadena para insertar
    consInsResp = sqlGuardaRespuesta.replace('{CSV}', str(csv) )
    consInsResp = consInsResp.replace('{TIPOFAC}', 'COMPRA')
    consInsResp = consInsResp.replace('{SERIE}', str(serieFactura) )
    consInsResp = consInsResp.replace('{NUMERO}', str(numeroFactura) )
    consInsResp = consInsResp.replace('{ESTADOENVIO}', str(estadoEnvio) )
    consInsResp = consInsResp.replace('{NUMSERIEFACTEMISOR}', str(nSerFacEm) )
    consInsResp = consInsResp.replace('{ESTADOREGISTRO}', str(estadoRegistro) )
    consInsResp = consInsResp.replace('{FECHAPRESENTACION}', str(fechaPresentacion) )
    consInsResp = consInsResp.replace('{CODERRORREGISTRO}', str(codErrorRegistro) )
    consInsResp = consInsResp.replace('{DESCRIPERRORREGISTRO}', str(descripErrorReg) )
    consInsResp = consInsResp.replace('{CSV2}', str(csv2) )

    logger.info("CSV: " + str(csv) + " | SERIE: " + str(serieFactura) + " | Numero: " + str(numeroFactura) +
                " | Estado Envio: " + str(estadoEnvio) + " | Num. Ser. Fac. Em.: " + str(nSerFacEm) +
                " | Estado Registro: " + str(estadoRegistro) + " | Fecha Presentacion: " + str(fechaPresentacion) +
                " | Cod. Error Registro: " + str(codErrorRegistro) + " | Descrip. Error: " + str(descripErrorReg) +
                " | CSV2: " + str(csv2) )

    logger.info("Consulta: " + str(consInsResp) )

    insercionBD = consultaMSSQL(consInsResp)
    logger.info('Respuesta de la insercion: ' + str(insercionBD) )
    logger.removeHandler(fh)


