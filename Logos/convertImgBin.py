import numpy as np
import cv2

nomFich = "stop_16_16"

fichEscritura = open(nomFich + ".txt", "w")

img = cv2.imread(nomFich +'.png')
imgGris = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

(thresh, imgBn) = cv2.threshold(imgGris, 127, 255, cv2.THRESH_BINARY)

# #Dimensiones imagen
# imgBn.shape
# (24, 24)
# imgBn.ndim
# 2
max_y,max_x = imgBn.shape

cadenaTexto=''

for y in range (0, max_y):
    for x in range (0, max_x):
        if(imgBn[y][x]==255):
            cadenaTexto+= '0'
        if(imgBn[y][x]==0):
            cadenaTexto+='1'
    cadenaTexto+='\n'
    fichEscritura.write(cadenaTexto)
    cadenaTexto = ''

fichEscritura.close()


