''' 
Comando para crear un fichero .xlsx con un valor de una celda [para crear reportes de datos]
- nombre de la pestaña
- letra de la columna
- numero de fila
- texto para poner en la celda
- y nombre del nuevo fichero .xlsx
'''

import sys
from openpyxl import Workbook


# Argumentos:
# 0: nombre del comando
# 1: nombre de la pestaña
# 2: columnas [A-Z]
# 3: fila 
# 4: texto o numero a insertar 'reporte de un dato!!! por ser un ejemplo :)'
# 4: nombre fichero .xlsx
#
# ejemplo uso:
# $ python3 CrearReporteDatosEnExcel.py Test B 1 holita salida2.xlsx


if len(sys.argv)<6:
	print("Error: No hay argumentos suficientes...\n")
	print("Los argumentos deben de ser:\n")
	print("1) nombre de la pestaña\n")
	print("2) columna [A-Z]\n")
	print("3) fila [numero]\n")
	print("4) valor\n")
	print("5) nombre del fichero .xlsx\n")	
else:

	nombrePestana=sys.argv[1]
	# columna: la letra la paso a numero de columna
	letraColumna=sys.argv[2].upper()
	# fila
	fila=int(sys.argv[3])
	valor=sys.argv[4]
	filePath = sys.argv[5]
	print(nombrePestana,letraColumna,fila,filePath)
	
	# Abrir fichero 
	# y selecciona la pestaña y lo pone como activa
	wb= Workbook()
	wb['Sheet'].title=nombrePestana
	#escribo los dato, en este caso al ser un ejemplo, solo escribo un dato.
	#zona a ampliar para meter mas datos...(leidos de un fichero externo, por ejemplo)
	wb.active[letraColumna+str(fila)].value=valor
	#guardo informacion creando un fichero nuevo con el nombre dado
	wb.save(filePath)

'''
Para leer, necesita tener instalado xlrd, con el siguiente comando se instala:
pip3 install xlrd==1.2.0 
Y para crear .xlsx necesita el openpyxl:
pip3 install openpyxl
'''
