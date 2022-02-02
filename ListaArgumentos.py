# Muestra argumentos introducidos en la linea de comandos:
# Para ejecutarlo: python3 ListaArgumentos.py hola mundo
#  mostrara la siguiente salida:
#       Argumentos count: 3
#       Argumentos      0; ListaArgumentos.py
#       Argumentos      1; hola
#       Argumentos      2; mundo
#fuente: https://realpython.com/python-command-line-arguments/

import sys

print(f"Argumentos count: {len(sys.argv)}")
for i, arg in enumerate(sys.argv):
    print(f"Argumentos {i:>6}; {arg}")
