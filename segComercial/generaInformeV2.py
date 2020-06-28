# -*- coding: cp1252 -*- 
#import sys
#reload(sys)
#sys.setdefaultencoding('utf8')

from libBDWebempresa import consultaMysql
#from libBD import consultaMSSQL
import time, codecs
import xlrd, xlwt
#import locale

rutaLog = ".\\Log\\"
rutaInformes = ".\\Informes\\"
cadTiempo = time.strftime("%d")+ "_" +  time.strftime("%m") + "_" + time.strftime("%Y") +  "_" + time.strftime("%H") +  "_" + time.strftime("%M") +  "_" + time.strftime("%S")
f = open(rutaLog + "Log_" + cadTiempo + ".txt", "w")
m = open(rutaInformes + "Macro_" + cadTiempo + ".txt", "w")

print rutaInformes + cadTiempo
#locale.setlocale(locale.LC_ALL, "es_ES.utf-8")

consDistUsuarios = u"""SELECT DISTINCT(USUARIO)
FROM VISITAS WHERE IDTEMP >= ( SELECT date_add(NOW(), INTERVAL -7 DAY)  )
ORDER BY FECHA"""

consDetInforme = """SELECT V.IDTEMP FECHA, V.CLIENTE, R.Texto RESULTADO, V.CONTACTO CONTACTO, V.EUROS PEDIDO, 
V.COMENTARIOS COMENTARIOS, V.COBROS,
( SELECT COUNT(*) FROM VISITAS WHERE IDTEMP >= ( SELECT date_add(NOW(), INTERVAL -7 DAY)  )
AND USUARIO = '{USUARIO}' ) LINEAS
FROM VISITAS V INNER JOIN RESULTADOS R ON V.Resultado = R.Resultado
WHERE IDTEMP >= ( SELECT date_add(NOW(), INTERVAL -7 DAY)  )
AND USUARIO = '{USUARIO}'
ORDER BY FECHA;
"""

macro = """    Sheets("{USUARIO}").Select
    Columns("E:E").Select
    Selection.TextToColumns Destination:=Range("E1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("G:G").Select
    Selection.TextToColumns Destination:=Range("G1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit
    Columns("F:F").Select
    Selection.ColumnWidth = 50
	Rows("1:1").RowHeight = 30
    Range("A1:G1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
	Range("A3:G{MAXPOS}").Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$3:$G${MAXPOS}"), , xlYes).Name = _
        "Tabla1"
    Range("A3:G{MAXPOS}").Select
    ActiveSheet.ListObjects("Tabla1").TableStyle = "TableStyleLight4"
    Range("E{POSSUM}").Activate
    ActiveCell.FormulaR1C1 = "=SUM(R[-{LINEAS}]C:R[-1]C)"
    Range("E{POSSUM}").Select
    Selection.NumberFormat = "#,##0.00 $"
    Range("G{POSSUM}").Activate
    ActiveCell.FormulaR1C1 = "=SUM(R[-{LINEAS}]C:R[-1]C)"
    Range("G{POSSUM}").Select
    Selection.NumberFormat = "#,##0.00 $"
    Range("C{POSSUM}").Select
    ActiveCell.FormulaR1C1 = "EUROS PEDIDOS:"
    Range("C{POSSUM}").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    'Euros Cobrados
    Range("F{POSSUM}").Select
    ActiveCell.FormulaR1C1 = "EUROS COBRADOS:"
    Range("F{POSSUM}").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    'Formato al nombre del comercial
    Range("B1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Name = "Calibri Light"
        .Size = 22
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    ActiveWorkbook.Save
	"""


f.write("Consulta --> " + str(consDistUsuarios) ) 
usuarios = consultaMysql(consDistUsuarios)
nExcel = xlwt.Workbook()

for u in usuarios:
    #Para cada usuario generaremos una Hoja
    nFila = 0
    nCol = 0
    usuario = u[0]
    nHoja = nExcel.add_sheet(usuario)
    #Insertamos el nombre del comercial como titulo
    nHoja.write(nFila, 1, "Seguimiento comercial - " + usuario )
    #Dejamos un hueco
    nFila = nFila + 2
    #Insertamos las cabeceras
    nHoja.write(nFila, nCol, "FECHA")
    nCol = nCol + 1
    nHoja.write(nFila, nCol, "CLIENTE")
    nCol = nCol + 1
    nHoja.write(nFila, nCol, "RESULTADO")
    nCol = nCol + 1	
    nHoja.write(nFila, nCol, "CONTACTO")
    nCol = nCol + 1
    nHoja.write(nFila, nCol, "PEDIDO")
    nCol = nCol + 1
    nHoja.write(nFila, nCol, "COMENTARIOS")
    nCol = nCol + 1
    nHoja.write(nFila, nCol, "COBROS")
    nCol = 0
    nFila = nFila + 1;
    #Fin de cabeceras
	
    sqlDetalle = consDetInforme.replace("{USUARIO}", usuario)
    f.write("--> Cons. SQL: " + str(sqlDetalle) + "\r\n")  
    detInforme = consultaMysql(sqlDetalle)
    f.write("--> Generando Excel: " + usuario + cadTiempo + "\r\n" )
    for d in detInforme:
        #FECHA
        #f.write("--> Anadimos Fecha \r\n");
        nHoja.write(nFila, nCol, str(d[0]).encode('utf8'))
        nCol = nCol + 1;
        #CLIENTE
        nHoja.write(nFila, nCol, unicode(d[1], errors = 'replace') + "\r\n" )
        #print str(d[1])
        #if( str(d[1]) == 'None' ):
            #nHoja.write(nFila, nCol, "")
        #else:
            #nHoja.write(nFila, nCol, str(d[1]).encode('cp1252'))
            #nHoja.write(nFila, nCol, unicode(str(d[1]), errors = 'replace' ))
            #nHoja.write(nFila, nCol, repr(d[1]).encode('cp1252'))
        nCol = nCol + 1;
        #RESULTADO
        nHoja.write(nFila, nCol, unicode(d[2], errors = 'replace') + "\r\n" )
        nCol = nCol + 1;
        #CONTACTO
        nHoja.write(nFila, nCol, unicode(d[3], errors = 'replace') + "\r\n" )
        nCol = nCol + 1;
        #PEDIDO
        nHoja.write(nFila, nCol, str(d[4].replace(".", ",")).encode('utf8'))
        nCol = nCol + 1;
        #COMENTARIOS
        nHoja.write(nFila, nCol, unicode(d[5], errors = 'replace') + "\r\n" )
        nCol = nCol + 1;
        #COBROS
        nHoja.write(nFila, nCol, str(d[6].replace(".", ",")).encode('utf8'))
        nCol = 0
        nFila = nFila + 1;
        maxPos = d[7]
    #Creamos la Macro
    macro2 = macro.replace("{USUARIO}", usuario)
    macro2 = macro2.replace("{MAXPOS}", str(maxPos+3) )
    macro2 = macro2.replace("{POSSUM}", str(maxPos+6) )
    macro2 = macro2.replace("{LINEAS}", str(maxPos+2) )
    m.write(macro2 + "\r\n")
    #print usuario, str(len(usuario))

# guardamos el excel
nExcel.save(rutaInformes + cadTiempo + ".xls")
f.write("--> Guardamos el Excel\r\n")
m.close()
f.close()
	


