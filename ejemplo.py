##########################################################################
# Leer ficheros excel desde python y mostrar la informacion que contiene #
##########################################################################


# Fuente https://www.youtube.com/watch?v=B3N8YvCEfx4
# previamente hay que instalar el modulo xlrd
# con: pip3 install xlrd==1.2.0 
# la version 1.2.0, lee los .xlsx
# fuente: https://stackoverflow.com/questions/65254535/xlrd-biffh-xlrderror-excel-xlsx-file-not-supported

import xlrd

# ruta donde este el fichero .xlsx
filePath = "./documento prueba.xlsx"

#filePath="/home/minino/Documentos/Gambas2/99 python/extraer informacion de excel/documento prueba.xlsx"

openFile = xlrd.open_workbook(filePath)

# Poner nombre exacto de la pestaña a trabajar de la excel
sheet = openFile.sheet_by_name("Sheet1")

print("Numero de filas",sheet.nrows)
print("Numero de Columnas",sheet.ncols)

for i in range(sheet.nrows):
	for j in range(sheet.ncols):
		print("[",i,",",j,"]=",sheet.cell_value(i,j), end="; ")
	print("\n")