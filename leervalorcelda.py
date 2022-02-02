''' 
Comando para mostrar el valor de una celda, dado el:
- nombre de la pestaña
- letra de la columna
- numero de fila
- y nombre del fichero .xlsx
'''




import sys
import xlrd

# Argumentos:
# 0: nombre del comando
# 1: nombre de la pestaña
# 2: columnas [A-Z]
# 3: fila 
# 4: nombre fichero .xlsx

if len(sys.argv)<5:
	print("Error: No hay argumentos suficientes...\n")
	print("Los argumentos deben de ser:\n")
	print("1) nombre de la pestaña\n")
	print("2) columna [A-Z]\n")
	print("3) fila [numero]\n")
	print("4) nombre del fichero .xlsx\n")	
else:
	filePath = sys.argv[4]
	nombrePestana=sys.argv[1]
	# columna: la letra la paso a numero de columna
	letraColumna=sys.argv[2].upper()
	numeroColumna=ord(letraColumna)-65
	# fila
	fila=int(sys.argv[3])-1
	
	# Abrir fichero y coger pestaña
	openFile = xlrd.open_workbook(filePath)
	sheet = openFile.sheet_by_name(nombrePestana)
	
	# prueba para mostrar los valores que lee
	# print(nombrePestana,letraColumna,numeroColumna,fila,filePath)
	
	#Muestra en consola el valor de la celda leida
	print(sheet.cell_value(fila,numeroColumna))


''' 
Para leer, necesita tener instalado xlrd, con el siguiente comando se instala:
pip3 install xlrd==1.2.0 
Y para escribir, necesita el xlswriter:
pip3 install xlsxwriter
'''
