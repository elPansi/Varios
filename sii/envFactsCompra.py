from fnEnvFC2 import enviaFacturaCompra
import time

serieFactura = raw_input("Serie Factura: ")
facturaIni = int(raw_input("Numero Factura Inicial: "))
facturaFin = int(raw_input("Numero Factura Final: "))

for i in range(facturaIni, facturaFin + 1):
    print "Enviando Factura: " + str(serieFactura) + " / " + str(i)
    enviaFacturaCompra(serieFactura, i)
    time.sleep(2)  # esperamos un par de segundos para la siguiente factura.
