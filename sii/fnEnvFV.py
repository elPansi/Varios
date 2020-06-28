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





def enviaFacturaVenta(serieFactura, numeroFactura):

    reload(sys)
    sys.setdefaultencoding('utf8')

    rutaLog = '.\\Log\\'
    rutaCertificados = '.\\certificados\\'
    rutaDatos = '.\\datos\\'

    # Configuracion de las series de Facturas y series de Facturas rectificativas
    serieFacturas = '18'
    serieFacturasRectificativas = 'R18'
    exentaIVA = 0

    # SQL DETALLE FACTURA
    sqlRecuperaFactura = """SELECT
    'EMPRESA' NOMBRERAZON, 'BXXXXX' NIF, 'A0' TIPOCOMUNICACION,
    YEAR(FC.FechaFactura) EJERCICIO, MONTH(fc.FechaFactura) PERIODO, 'BXXXXX' IDEMISORFACTURA,
    ( FS.SERIE + ' / ' + STR(FC.NUMFACTURA) ) NUMSERIEFACTURAEMISOR,
    CONVERT(NVARCHAR(10),FC.FechaFactura, 103) FECHAEXPEDICIONFACTURAEMISOR,
    (SELECT CASE
    WHEN FS.Serie = '{SERIEFACTURAS}' THEN 'F1' --SERIE FACTURAS NORMALES
    WHEN FS.SERIE = '{SERIERECTIFICADAS}' THEN 'R1' --SERIE FACTURAS RECTIFICATIVAS
    END ) TIPOFACTURA ,
    --FACTURAS AGRUPADAS
    (SELECT CASE
    WHEN FS.Serie = '{SERIEFACTURAS}' THEN ( STR(FC.NUMFACTURA) + ' - ' +  FS.SERIE)
    WHEN FS.Serie = '{SERIERECTIFICADAS}' THEN NULL
    END ) NUMSERIEFACTURAEMISOR,
    (SELECT CASE
    WHEN FS.Serie = '{SERIEFACTURAS}' THEN ( CONVERT(NVARCHAR(10),FC.FechaFactura, 103) )
    WHEN FS.Serie = '{SERIERECTIFICADAS}' THEN NULL
    END ) FECHAEXPEDICIONFACTURAEMISOR,
    --FACTURAS RECTIFICADAS
    (SELECT CASE
    WHEN FS.Serie = '{SERIEFACTURAS}' THEN  NULL
    WHEN FS.Serie = '{SERIERECTIFICADAS}' THEN ( FS.SERIE + ' - ' + STR(FC.NUMFACTURA) )
    END ) NUMSERIEFACTURAEMISOR,
    (SELECT CASE
    WHEN FS.Serie = '{SERIEFACTURAS}' THEN NULL
    WHEN FS.Serie = '{SERIERECTIFICADAS}' THEN ( CONVERT(NVARCHAR(10),FC.FechaFactura, 103) )
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
    WHEN ( P.CODIGO IN ( 35, 38, 51, 52 ) ) THEN '09' -- IGIC/IPSI
    WHEN ( FC.retencion <> 0 AND SUM(FL.CANTIDAD*FL.Precio*(1-Fl.Dto1/100)*Fl.iva/100) <> 0 ) THEN '13' -- NAVES?? 17/507
    --WHEN ( FC.retencion <> 0 AND SUM(FL.CANTIDAD*FL.Precio*(1-Fl.Dto1/100)*Fl.iva/100) = 0 ) THEN 'XX'
    WHEN ( FL.CODARTICULO IN ( 8790 ) ) THEN '14' --ALQUILERES, AQUI HAY QUE IDENTIFICAR POR EL CODIGO DEL ARTICULO
    ELSE '01'
    END ) CLAVEREGIMENESPECIALOTRASCENDENCIA,
    'FACTURA DE VENTA ' + FS.SERIE + ' / ' + STR(FC.NumFactura)  DESCRIPCIONOPERACION,
    1 SITUACIONINMUEBLE, CONVERT(NVARCHAR(40), C.NOMBRE) NOMBRERAZONCLIENTE, C.CIF NIFCLIENTE,
    '02' TIPOIDENTIFICACIONPAISRESIDENCIA, C.CIF NUMEROIDENTIFICACION,
    FC.OBSERVACIONES OBSERVACIONES
    FROM FACTURACAB FC FULL OUTER JOIN articulosdetalles FL ON FC.CodFactura = FL.CodFactura
    FULL OUTER JOIN CLIENTES C ON C.CodCliente = FC.CodCliente
    FULL OUTER JOIN PROVINCIAS P ON P.codigo = C.CodProvincia
    INNER JOIN FACTURASERIE FS ON FS.CODSERIE = FC.SERIE
    WHERE FS.SERIE = '{SERIEFACTURA}' AND FC.NumFactura = {NUMEROFACTURA}
    GROUP BY FC.CodFactura, FS.SERIE, FC.NumFactura, C.Nombre, C.Direccion, FL.CodArticulo,
    C.Poblacion, C.CP, P.Nombre, C.CIF, FC.FechaFactura, FC.retencion, FC.CodCliente, p.codigo, FC.OBSERVACIONES
    ORDER BY FS.Serie, FC.NumFactura"""

    ##Cadenas ...
    sqlRecuperaImpuesto = """SELECT FC.CODFACTURA, FS.SERIE, FC.NumFactura, FL.IVA,
    SUM(FL.CANTIDAD*FL.Precio*(1-Fl.Dto1/100)*(1-Fl.Dto2/100)) BASEIMPONIBLE,
    SUM(FL.CANTIDAD*FL.PRECIO*(1-FL.DTO1/100)*(1-Fl.Dto2/100)*FL.IVA/100) CUOTAIVA
    FROM FACTURACAB FC FULL OUTER JOIN articulosdetalles FL ON FC.CODFACTURA = FL.CODFACTURA
    INNER JOIN FACTURASERIE FS ON FS.CODSERIE = FC.SERIE
    WHERE FS.Serie = '{SERIEFACTURA}' AND FC.NumFactura = {NUMEROFACTURA} AND FL.IVA = {IVA}
    GROUP BY FC.CODFACTURA, FS.SERIE, FC.NumFactura, FL.IVA"""

    sqlDistintosImpuestos = """SELECT SUBCONSULTA.iva
    FROM  (SELECT CodFactura, IVA, Cantidad FROM articulosdetalles
    WHERE CodFactura = (SELECT CODFACTURA FROM FacturaCab
    WHERE SERIE = ( SELECT CODSERIE FROM FacturaSerie WHERE SERIE = '{SERIEFACTURA}' )
    AND NumFactura = {NUMEROFACTURA} )
    GROUP BY Codfactura, iva, CANTIDAD) SUBCONSULTA
    WHERE SUBCONSULTA.iva * SUBCONSULTA.Cantidad <> 0
    GROUP BY iva"""

    sqlIvaAlCero = """SELECT SUBCONSULTA.iva, SUBCONSULTA.BASEIMP
    FROM  (SELECT CodFactura, IVA, Cantidad, SUM(CANTIDAD*PRECIO*(1-DTO1/100)) AS BASEIMP
    FROM articulosdetalles
    WHERE CodFactura = (SELECT CODFACTURA FROM FacturaCab WHERE SERIE =
    (SELECT CODSERIE FROM FACTURASERIE WHERE SERIE = '{serie}' ) AND NumFactura = {numDocPropio}  )
    GROUP BY Codfactura, iva, CANTIDAD, PRECIO, DTO1) SUBCONSULTA
    WHERE SUBCONSULTA.IVA = 0
    GROUP BY SUBCONSULTA.iva, SUBCONSULTA.BASEIMP
    """

    sqlExencion = """SELECT TOP(1) SUBSTRING(IDDEARTICULO,1,2) EXENCION
    FROM ARTICULOS WHERE CODARTICULO IN(
    SELECT CODARTICULO FROM ARTICULOSDETALLES
    WHERE CodFactura = (SELECT CODFACTURA FROM FacturaCab WHERE ( SERIE = ( SELECT CODSERIE FROM FACTURASERIE WHERE SERIE ='{SERIEFACTURA}') )
    AND NumFactura = {NUMEROFACTURA}  )
    AND IDDeArticulo LIKE 'E%'  AND CodFamilia = 34 AND BloqueoVentas = 0) """

    sqlBaseImponible = """SELECT FC.CODFACTURA, FS.SERIE, FC.NumFactura, FL.IVA,
    SUM(FL.CANTIDAD*FL.Precio*(1-Fl.Dto1/100)) BASEIMPONIBLE,
    SUM(FL.CANTIDAD*FL.PRECIO*(1-FL.DTO1/100)*FL.IVA/100) CUOTAIVA
    FROM FACTURACAB FC FULL OUTER JOIN articulosdetalles FL ON FC.CODFACTURA = FL.CODFACTURA
    INNER JOIN FACTURASERIE FS ON FS.CODSERIE = FC.SERIE
    WHERE FS.Serie = '{SERIEFACTURA}' AND FC.NumFactura = {NUMEROFACTURA} AND FL.IVA = 0
    GROUP BY FC.CODFACTURA, FS.SERIE, FC.NumFactura, FL.IVA"""

    sqlBaseRectificativas = """SELECT FC.CODFACTURA, FS.SERIE, FC.NumFactura,
    SUM(FL.CANTIDAD*FL.Precio*(1-Fl.Dto1/100)) BASEIMPONIBLE,
    SUM(FL.CANTIDAD*FL.PRECIO*(1-FL.DTO1/100)*FL.IVA/100) CUOTAIVA
    FROM FACTURACAB FC FULL OUTER JOIN articulosdetalles FL ON FC.CODFACTURA = FL.CODFACTURA
    INNER JOIN FACTURASERIE FS ON FS.CODSERIE = FC.SERIE
    WHERE FS.Serie = '{SERIEFACTURA}' AND FC.NumFactura = {NUMEROFACTURA}
    GROUP BY FC.CODFACTURA, FS.SERIE, FC.NumFactura"""

    sqlFacturaARectificar = """SELECT FCC.OBSERVACIONES
    FROM FacturaCab FCC INNER JOIN FacturaSerie FS ON FCC.SERIE= FS.CodSerie
    WHERE FS.SERIE = '{SERIE}' AND FCC.NumFactura = {NUMERO}
    AND UPPER(Observaciones) LIKE '%RECTIFICA: %'"""

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

    sqlContabilizada = """SELECT Contabilizado FROM FACTURACAB FC INNER JOIN FacturaSerie FS ON FS.CodSerie = FC.SERIE
    WHERE FS.SERIE = '{SERIEFACTURA}' AND FC.NumFactura = {NUMEROFACTURA} ORDER BY NumFactura DESC"""

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

    cadDesgloseFactura = """
                                 "DesgloseIVA":{
                                    "DetalleIVA": [
                                          {DESGLOSESIVAS}
                                     ]
                                  }"""

    cadDesgloseIVA = """  	{
    							"TipoImpositivo":{porcenIVA},
    							"BaseImponible":{baseImp},
    							"CuotaRepercutida":{cuotaIVA}
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

    cadImpuestos = """  "DesgloseFactura":{
    				"Sujeta":{
    					"NoExenta":{
    						"TipoNoExenta":"S1",
                            {DESGLOSEIVA}
    					}
    				}
    	        },
    	        "DesgloseTipoOperacion":{
    			    "PrestacionServicios":{
    				    "Sujeta":{
    					    "NoExenta":{
    						"TipoNoExenta":"S1",
                            {DESGLOSEIVA}
    						}
                        }
                    }
                }"""

    cadExentaImpuestos = """	"DesgloseFactura":{
    				"Sujeta":{
    					"Exenta":{
    						"CausaExencion":"{EXENCION}",
                            "BaseImponible": "{BASEIMPONIBLE}"
    					}
    				}
    	        },
    	        "DesgloseTipoOperacion":{
    			    "PrestacionServicios":{
    				    "Sujeta":{
    					    "Exenta":{
    						"CausaExencion":"{EXENCION}",
                            "BaseImponible": "{BASEIMPONIBLE}"
    						}
                        }
                    }
    			}"""

    cadFacturaManual = """{
        "PeriodoImpositivo":{
    		"Ejercicio":"2017",
    		"Periodo":"07"
    		},
    	"IDFactura":{
    		"IDEmisorFactura":{
    		    "NIF":"BXXXXX",
    		    },
    		"NumSerieFacturaEmisor":"17 /       1079",
    		"FechaExpedicionFacturaEmisor":"31/07/2017"
    		},
    	"FacturaExpedida":{
    		"TipoFactura":"F1",

    		"ClaveRegimenEspecialOTrascendencia":"01",
    		"DescripcionOperacion":"ESTRUCTURA METALICA",
    		"Contraparte":{
    			"NombreRazon":"PROSOLIA FRANCE SARL",
    			"IDOtro":{
                     'CodigoPais': 'FR',
    			     "IDType":"04",
    				 "ID":"FR90502188717"
    				 }
    			},
            "TipoDesglose":{
    	        "DesgloseTipoOperacion":{
    			    "PrestacionServicios":"",
                    "Entrega":{
    				    "Sujeta":{
    					    "Exenta":{
    						    "CausaExencion":"E5",
                                "BaseImponible":"     4634.19"
    						}
                        }
                    }
                }
                }
        }
    }"""

    cadFacturaManual = """{
        "PeriodoImpositivo":{
    		"Ejercicio":"2017",
    		"Periodo":"08"
    		},
    	"IDFactura":{
    		"IDEmisorFactura":{
    		    "NIF":"BXXXXX",
    		    },
    		"NumSerieFacturaEmisor":"17 /       1261",
    		"FechaExpedicionFacturaEmisor":"25/08/2017"
    		},
    	"FacturaExpedida":{
    		"TipoFactura":"F1",

    		"ClaveRegimenEspecialOTrascendencia":"01",
    		"DescripcionOperacion":"FACTURA DE VENTA 17 /       1261",
    		"Contraparte":{
    			"NombreRazon":"SARL HOTEL CLUB LAROSE BLEVE",
    			"IDOtro":{
                                     "CodigoPais":"DZ",
    			         "IDType":"04",
    				 "ID":"000816098137052"
    				 }
    			},
            "TipoDesglose":{
                     "DesgloseFactura":{
    				"Sujeta":{
    					"NoExenta":{
    						"TipoNoExenta":"S1",

                                 "DesgloseIVA":{
                                    "DetalleIVA": [
                                            	{
    							"TipoImpositivo":21.00,
    							"BaseImponible":       83.33,
    							"CuotaRepercutida":       17.50
    						}
                                     ]
                                  }
    					}
    				}
    	        },
    	        "DesgloseTipoOperacion":{
    			    "PrestacionServicios":{
    				    "Sujeta":{
    					    "NoExenta":{
    						"TipoNoExenta":"S1",

                                 "DesgloseIVA":{
                                    "DetalleIVA": [
                                            	{
    							"TipoImpositivo":21.00,
    							"BaseImponible":       83.33,
    							"CuotaRepercutida":       17.50
    						}
                                     ]
                                  }
    						}
                        }
                    }
                }
                }
        }
    }"""

    #Preparamos la cadena del diccionario
    cadCabecera = """{
            "IDVersionSii":1.0,
            "Titular":{
                "NombreRazon":"EMPRESA",
                "NIF":"BXXXXX"
                },
            "TipoComunicacion":"A0"}"""

    #FACTURAS
    cadFactura="""{
            "PeriodoImpositivo":{
                "Ejercicio":"{anio}",
                "Periodo":"{mes}"
                },
            "IDFactura":{
                "IDEmisorFactura":{
                    "NIF":"BXXXXX",
                    },
                "NumSerieFacturaEmisor":"{numSerieFactEmisor}",
                "FechaExpedicionFacturaEmisor":"{fechaExpFactEm}"
                },
            "FacturaExpedida":{
                "TipoFactura":"{tipoFactura}",
                {CADRECTIFICADAS}
                "ClaveRegimenEspecialOTrascendencia":"{REGESPECIAL}",
                "DescripcionOperacion":"{descripOperacion}",
                "Contraparte":{
                    "NombreRazon":"{nombreCliente}",
                    "NIF":"{nifCliente}",
                    "IDOtro":{
                         "IDType":"02",
                         "ID":"ES"
                         }
                    },
                "TipoDesglose":{
                       {CADDESGLOSEFACTURA}
                    }
            }
        }"""

    cadTiempo = time.strftime("%Y") + time.strftime("%m") + time.strftime("%d") + time.strftime("%H") + time.strftime(
        "%M") + time.strftime("%S") + "_" + str(serieFactura) + "_" + str(numeroFactura) + "_"

    logger = logging.getLogger('zeep')
    logger.setLevel(logging.DEBUG)
    fh = logging.FileHandler(rutaLog + cadTiempo + '_vta_logSii.log')
    fh.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    logger.info( 'Iniciando el envio de la factura: ' + str(serieFactura) + ' / ' + str(numeroFactura) )

    #Primero comprobamos que la factura este contabilizada
    consContabilizada = sqlContabilizada.replace("{SERIEFACTURA}", str(serieFacturas) )
    consContabilizada = consContabilizada.replace("{NUMEROFACTURA}", str(numeroFactura) )

    conta = consultaMSSQL(consContabilizada)
    #print conta[0][0]
    if conta[0][0] == True:
        logger.info("Factura Contabilizada ...: " + str(serieFacturas) + " / " + str(numeroFactura) )
    else:
        logger.info("Error, factura sin contabilizar. No se enviará la factura: "  + str(serieFacturas) + " / " + str(numeroFactura) )
        sys.exit(1)

    consDistintosImpuestos = sqlDistintosImpuestos.replace("{SERIEFACTURA}", str(serieFactura) )
    consDistintosImpuestos = consDistintosImpuestos.replace("{NUMEROFACTURA}", str(numeroFactura) )
    logger.info("Consulta Impuestos: " + consDistintosImpuestos)

    impuestosFactura = consultaMSSQL(consDistintosImpuestos)


    consRecFact = sqlRecuperaFactura.replace("{SERIEFACTURAS}", str(serieFacturas) )
    consRecFact = consRecFact.replace("{SERIERECTIFICADAS}", str(serieFacturasRectificativas) )
    consRecFact = consRecFact.replace("{SERIEFACTURA}", str(serieFactura) )
    consRecFact = consRecFact.replace("{NUMEROFACTURA}", str(numeroFactura) )



    logger.info( "Impuestos --> " + str(impuestosFactura ) )
    logger.info( consRecFact )


    #Recuperamos la factura
    detalleFactura = consultaMSSQL(consRecFact)


    logger.info(detalleFactura)
    #logger.info(detalleImpuestos)

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
    anulacion = detalleFactura[0][22]


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
        logger.info("Base Rectificada: " + str(baseRectif) + " Cuota Rectificada: " + str(cuotaRectif) + " FechaFactEmisor: " + str(fechaExpFactEmisor))
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
    cadFinExenta = ""

    logger.info("Comenzamos a tratar los ivas ... " + str(len(impuestosFactura)) + " | " )
    #Primero debemos cargar los Desgloses de IVA en caso de que exista
    if len(impuestosFactura) > 0:
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
            cadParDesglose = cadDesgloseIVA.replace("{porcenIVA}","{:3.2f}".format(iva) )
            cadParDesglose = cadParDesglose.replace("{baseImp}", "{:12.2f}".format(base) )
            cadParDesglose = cadParDesglose.replace("{cuotaIVA}", "{:12.2f}".format(cuota) )
            logger.info( "cadParDesglose: " + cadParDesglose )
            if i == 0:
                cadDesgloseImpuestos = cadParDesglose
                #print "Primer Impuesto: " + cadDesgloseImpuestos
            else:
                cadDesgloseImpuestos = cadDesgloseImpuestos + "," + cadParDesglose
                #print "Cadena final Impuestos" + cadDesgloseImpuestos
            cuotaIVATotal = cuotaIVATotal + cuota
        #aqui debemos comprobar si existe el iva al 0%
        consIvaAlCero = sqlIvaAlCero.replace("{serie}", str(serieFactura) )
        consIvaAlCero = consIvaAlCero.replace("{numDocPropio}", str(numeroFactura) )
        respIvaCero = consultaMSSQL(consIvaAlCero)
        if respIvaCero:
            base = respIvaCero[0][1]
            cadParDesglose = cadDesgloseIVA.replace("{porcenIVA}","{:3.2f}".format(0) )
            cadParDesglose = cadParDesglose.replace("{baseImp}", "{:12.2f}".format(base) )
            cadParDesglose = cadParDesglose.replace("{cuotaIVA}", "{:12.2f}".format(0) )
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
        print "Exencion: " + str(exencion)
        if exencion: #Ha encontrado una E, Inversion de sujeto Pasivo -> Sujeta/Exenta
            print "Inversion de Sujeto Pasivo"
            cadFinExenta = cadExentaImpuestos.replace("{EXENCION}", str(exencion[0][0]) )
            cadFinExenta = cadFinExenta.replace ("{BASEIMPONIBLE}", "{:12.2f}".format(base[0][4]) )
        else: #Es una factura Normal --> Desglose IVA
            print "DesgloseIVA"
            cadFinExenta = cadDesgloseIVA.replace("{porcenIVA}", "0.00")
            cadFinExenta = cadFinExenta.replace ("{baseImp}", "{:12.2f}".format(base[0][4]) )
            cadFinExenta = cadFinExenta.replace("{cuotaIVA}", "0.00")

		
		
    # cadFinExenta = cadExentaImpuestos.replace( "{EXENCION}", str(exencion[0][0]) )
	
    # cadFinExenta = cadFinExenta.replace("{BASEIMPONIBLE}", "{:12.2f}".format(base[0][4]) )


    if len(impuestosFactura) >= 1:
        cadFinImpuestos = cadDesgloseImpuestos
    else:
        cadFinImpuestos = cadFinExenta


    logger.info( "cadFinExenta: --> :" + str(cadFinExenta) )
    logger.info( "Cad Fin Impuestos: --> : " + str(cadFinImpuestos) )
    #logger.info( "Cad Fin Exenta: --> " + str(cadFinExenta) )

    #Ahora reemplazamos el desglose en los impuestos


    # cadFinImpuestos = cadImpuestos.replace("{DESGLOSEIVA}", str(cadDesgloseImpuestos) )
    logger.info("FechaExpedicion: " + str(fechaExpFactEmisor) )

    fact = cadFactura.replace("{anio}", str(ejercicio) )
    fact = fact.replace("{mes}", str(mes).zfill(2) )
    fact = fact.replace("{numSerieFactEmisor}", str(numSerieFactEmisor) )
    fact = fact.replace("{fechaExpFactEm}", str(fechaExpFactEmisor) )
    fact = fact.replace("{tipoFactura}", str(tipoFactura) )
    fact = fact.replace("{descripOperacion}", str(descripOperacion) )
    fact = fact.replace("{nombreCliente}", str(nombreCliente) )
    fact = fact.replace("{nifCliente}", str(nifCliente) )
    fact = fact.replace("{REGESPECIAL}", str(regEspecial) )

    logger.info("CadFinImpuestos: " + cadFinImpuestos)
    cadDesFact = cadDesgloseFactura.replace("{DESGLOSESIVAS}", str(cadFinImpuestos) )

    logger.info( "CadDesFact: --> " + str(cadDesFact) )
    #concatenamos la cadena de impuestos


    if exentaIVA == 0: #No tiene exencion Factura Normal
        logger.info("Factura Normal - cadDesFact: " + str(cadDesFact) )
        logger.info("cadFinExenta: " + str(cadFinExenta) )
        logger.info("cadFinImpuestos: " + str(cadFinImpuestos) )
        cadDesgloseFactura = cadImpuestos.replace("{DESGLOSEIVA}", str(cadDesFact) )
        fact = fact.replace("{CADDESGLOSEFACTURA}", str(cadDesgloseFactura ) )#antes cadDesFact
    else:
        #caso de sujeta bien
        logger.info("Sujeta o sin IVA")
        logger.info("Exenta IVA: " + str(exentaIVA) )
        logger.info("Exencion : " + str(exencion) )
        if exencion:
            logger.info("cadFinExenta: " + str(cadFinExenta) )
            logger.info("cadFinImpuestos: " + str(cadFinImpuestos) )
            fact = fact.replace("{CADDESGLOSEFACTURA}", str(cadFinExenta) ) #antes cadFinExenta
        else:
            logger.info( "CadDesFact: (2) --> " + str(cadDesFact) )
            cadDesgloseFactura = cadImpuestos.replace("{DESGLOSEIVA}", str(cadDesFact) )
            logger.info("CadDesgloseFactura: " + cadDesgloseFactura )
            fact = fact.replace("{CADDESGLOSEFACTURA}", str(cadDesgloseFactura) ) #antesw cadDesgloseFactura



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


    wsdl = rutaDatos + 'SuministroFactEmitidas.wsdl'
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
    client = Client(wsdl=wsdl,port_name="SuministroFactEmitidas",transport=transport, service_name = 'siiService')
    logger.info('Establecida conexion con cliente')
    #SELECCION SERVICIO DE PRUEBAS        SuministroLRFacturasEmitidas
    service2 = client.bind('siiService', 'SuministroFactEmitidas')
    logger.info('Bind al servicio')
    #INSERCION DE LA FACTURA Y RESPUESTA EN EL SERVICO DE PRUEBAS.
    #print (service2.SuministroLRFacturasEmitidas(cabecera,facturas))
    logger.info('Enviando factura ...')
    try:
        salida = (service2.SuministroLRFacturasEmitidas(dicCabecera,dicFactura))
    except:
        e = sys.exc_info()[0]
        logger.info("Error al enviar factura: " + str(e) )
        salida = None
        datosPres = None

    logger.info('Factura envidada')
    logger.info('Salida: ' + str(salida) )

    #Procesamos la salida
    if salida:
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
    consInsResp = consInsResp.replace('{TIPOFAC}', 'VENTA')
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
    print "Terminado..."
    logger.removeHandler(fh)


