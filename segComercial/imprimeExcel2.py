# -*- coding: cp1252 -*- 
#import sys
#reload(sys)
#sys.setdefaultencoding('utf8')

from libBDWebempresa import consultaMysql
#from libBD import consultaMSSQL
import time, codecs
import xlrd, xlwt
#import locale

# def ls(ruta = Path.cwd()):
#     return [arch.name for arch in Path(ruta).iterdir() if arch.is_file()]

#archivos = ls(rutaDatos)
# print str(archivos)


rutaLog = ".\\Log\\"
rutaInformes = ".\\Informes\\"
rutaDatos = ".\\excel\\"

cadTiempo = time.strftime("%d")+ "_" +  time.strftime("%m") + "_" + time.strftime("%Y") +  "_" + time.strftime("%H") +  "_" + time.strftime("%M") +  "_" + time.strftime("%S")
f = open(rutaLog + "Log_Excel_" + cadTiempo + ".txt", "w")
m = open(rutaInformes + "Macro_Excel_" + cadTiempo + ".txt", "w")

#print rutaInformes + cadTiempo

docExcel = xlrd.open_workbook(rutaDatos + 'excelEntrada2.xlsx')
hojaExcel = docExcel.sheet_by_index(0)

docSalida = xlwt.Workbook()



#leemos los usuarios del excel
fila = 1
filaMax = hojaExcel.nrows
f.write( "Fila Max: " + str(filaMax) + "\r\n")

listaUsuarios = []
while fila < filaMax:
    #cogemos los apellidos
    codUsuario = hojaExcel.cell_value(fila, 0 )
    mes = hojaExcel.cell_value(fila, 9 )
    ano = hojaExcel.cell_value(fila, 10 )
    elementosPestana = [codUsuario, mes, ano]
    if elementosPestana not in listaUsuarios:
        listaUsuarios.append(elementosPestana)
    fila = fila + 1

print str(listaUsuarios)
contadorRegistros = 0

#para cada usuario vamos a ver los años y los meses que tiene
for u in listaUsuarios:
    fila = 1
    nFila = 1
    usuarioActual = u[0]
    mesActual = u[1]
    anoActual = u[2]
    for a,m in listaUsuarios[2],listaUsuarios[1]: #2 anos
        #for m in listaUsuarios[1]: #1 meses
        if m == mesActual:
            #Creamos la pestana del mes
            f.write("Fila excel Destino: " + str(nFila))
                f.write("Usuario: " + str(usuarioActual) + " - Mes actual: " + str(mesActual) + " - Ano actual: " + str(
                    anoActual) + "\r\n")
                cadenaPestana = str(int(usuarioActual)) + "_" + str(int(mesActual)) + "_" + str(int(anoActual))
                nHoja = docSalida.add_sheet(cadenaPestana)
                if (int(hojaExcel.cell_value(fila, 0)) == int(usuarioActual) \
                        and int(hojaExcel.cell_value(fila, 10)) == int(anoActual) \
                        and int(hojaExcel.cell_value(fila, 9)) == int(mesActual)):
                        #
                        # while fila < filaMax:
                        #     temp = 1
                        #     while temp=1:
                        #         nFila = 1

                        #         f.write(str(hojaExcel.cell_value(fila, 0)) + "  --  " + str(hojaExcel.cell_value(fila, 9)) + "  --  " + str(hojaExcel.cell_value(fila, 10)) + "\r\n" )
                        #
                    contadorRegistros = contadorRegistros + 1
                    #insertamos el usuario
                    nHoja.write(nFila, 1, int(usuarioActual) )
                    #insertamos el nombre
                    nHoja.write(nFila, 2, hojaExcel.cell_value(fila, 1))
                    # insertamos el Apellido
                    nHoja.write(nFila, 3, hojaExcel.cell_value(fila, 2))
                    # Departamento
                    nHoja.write(nFila, 4, hojaExcel.cell_value(fila, 3))
                    #numPersonal
                    nHoja.write(nFila, 5, hojaExcel.cell_value(fila, 4))
                    # Fichaje
                    nHoja.write(nFila, 6, hojaExcel.cell_value(fila, 5))
                    #Dispositivo
                    nHoja.write(nFila, 7, hojaExcel.cell_value(fila, 6))
                    #entSalida
                    nHoja.write(nFila, 8, hojaExcel.cell_value(fila, 7))
                    #verificacion
                    nHoja.write(nFila, 9, hojaExcel.cell_value(fila, 8))
                nFila = nFila + 1
                fila = fila + 1
            nFila = 0
    f.write( "Registros: " + str(contadorRegistros) + "\r\n" )
    contadorRegistros = 0


# # guardamos el excel
docSalida.save(rutaDatos + cadTiempo + "_ExcelFinal" + ".xls")
# f.write("--> Guardamos el Excel\r\n")
m.close()
f.close()
	


