import requests
from libBDWebempresa import consultaMysql
import time, codecs
import ast
import sys

rutaLog = ".\\Log\\"
rutaInformes = ".\\Informes\\"
cadTiempo = time.strftime("%d")+ "_" +  time.strftime("%m") + "_" + time.strftime("%Y") +  "_" + time.strftime("%H") +  "_" + time.strftime("%M") +  "_" + time.strftime("%S")
f = open(rutaLog + "Log_Act_Pos" + cadTiempo + ".txt", "w")


sqlRegistros = u"""SELECT FECHA FROM VISITAS
WHERE LATITUD = 0 AND LONGITUD = 0
AND ( MCC <> 0 AND MNC <> 0 AND LAC <> 0 AND CID <> 0 )
ORDER BY IDTEMP DESC"""


sqlDetalles = u"""SELECT FECHA, MCC, MNC, LAC, CID
FROM VISITAS WHERE FECHA = '{FECHA}'"""

sqlActualizaCoordenadas = u"""UPDATE VISITAS
SET LATITUD = {LATITUD}, LONGITUD = {LONGITUD}
WHERE FECHA = '{FECHA}'"""

url = "https://eu1.unwiredlabs.com/v2/process.php"
payload = "{\"token\": \"xxxxxxxx\",\"radio\": \"gsm\",\"mcc\": {MMC},\"mnc\": {MNC},\"cells\": [{\"lac\": {LAC},\"cid\": {CELLID}}],\"address\": 1}"

f.write("Consulta Recuperar Registros --> " + str(sqlRegistros) + "\r\n")
registros = consultaMysql(sqlRegistros)


#actualizamos cada uno de los registros
for r in registros:
    fecha = r[0]
    detalles = consultaMysql(sqlDetalles.replace("{FECHA}", fecha))
    f.write("Detalles: " + str(detalles) + "\r\n" )
    mcc = detalles[0][1]
    f.write("MCC: " + str(mcc) + "\r\n" )
    mnc = detalles[0][2]
    f.write("MNC: " + str(mnc) + "\r\n" )
    lac = detalles[0][3]
    f.write("LAC: " + str(lac) + "\r\n" )
    cid = detalles[0][4]
    f.write("Cell ID: " + str(cid) + "\r\n" )

    cadPayload = payload.replace("{MMC}", mcc)
    cadPayload = cadPayload.replace("{MNC}", mnc)
    cadPayload = cadPayload.replace("{LAC}", lac)
    cadPayload = cadPayload.replace("{CELLID}", cid)

    response = requests.request("POST", url, data=cadPayload)
    f.write("Respuesta: " + str(response.text) + "\r\n")

    respuesta = ast.literal_eval(response.text)

    balance = respuesta['balance']
    f.write("Iteracion: " + str(balance) + "\r\n" )

    if balance == 0:
        f.write("Alcanzado numero maximo de iteraciones por hoy ..." + "\r\n" )
        sys.exit(1)
    f.write("Fecha: " + str(fecha) + "\r\n")
    latitud = respuesta['lat']
    f.write("Latitud: " + str(latitud)+ "\r\n")
    longitud = respuesta['lon']
    f.write("Longitud: " + str(longitud)+ "\r\n")

    consActualizaCoordenadas = sqlActualizaCoordenadas.replace("{FECHA}", str(fecha))
    consActualizaCoordenadas = consActualizaCoordenadas.replace("{LATITUD}", str(latitud))
    consActualizaCoordenadas = consActualizaCoordenadas.replace("{LONGITUD}", str(longitud))

    f.write("Consulta Actualizacion: " + str(consActualizaCoordenadas) + "\r\n")

    salida = consultaMysql(consActualizaCoordenadas)
    print ("Fecha: " + str(fecha))







