#Importar panda y numpy

import pandas
import numpy
import openpyxl

#Abrir archivo1

m = openpyxl.load_workbook('Archivo1.xlsx',data_only = True)

#Activar la hoja del archivo1 en la que se encuentran los datos

hoja1 = m.active

#Acceder a las celdas que continen la información a utilizar

celdas = hoja1['A1':'A3']

#Recorrer y obtener el contenido de las celdas para añadirlo a una lista

matrizm = []

for fila in celdas:
  a = [celda.value for celda in fila] 
  matrizm.append(a)

#print(matrizm)

#Convertir la lista obtenida en una serie

serie1 = pandas.Series(matrizm)

#Añadir los nombres de las celdas en la que se encuentran los datos

serie12 = pandas.Series(matrizm,index=["A1","A2","A3"])
#print(str(serie12))

#Abrir archivo2

n = openpyxl.load_workbook('Archivo2.xlsx',data_only = True)

#Activar la hoja del archivo1 en la que se encuentran los datos

hoja2 = n.active

#Acceder a las celdas que continen la información a utilizar

celdas2 = hoja2['A1':'A3']

#Recorrer y obtener el contenido de las celdas para añadirlo a una lista

matrizn = []

for fila in celdas2:
  b = [celda.value for celda in fila] 
  matrizn.append(b)

#print(matrizn)

#Convertir la lista obtenida en una serie

serie2 = pandas.Series(matrizn)

#print(str(serie2))

#Añadir los nombres de las celdas en la que se encuentran los datos

serie22 = pandas.Series(matrizn,index=["A1","A2","A3"])

#Crear el archivo 3 en donde se van a guardar las respuestas del archivo 1

e = open('Archivo3.xlsx','w+')

#Crear el archivo 4 en donde se van a guardar las respuestas del archivo 2

f = open('Archivo4.xlsx','w+')

##Trabajo en el archivo 3

#Escribir el contenido del archivo 1 en el archivo 3

e.write("El contenido del archivo 1 es:\n")
e.write(str(matrizm))

#Escribir como una serie el contenido del archivo 1 en el archivo 3 

e.write("\n\nLa serie con el contenido del archivo 1 es:\n")
e.write(str(serie1))

#Escribir la serie con indices del contenido del archivo 1 en el archivo 3 

e.write("\n\nLa serie con el contenido del archivo 1 con sus respectivos nombres de celdas es:\n")
e.write(str(serie12))

##Trabajo en el archivo 4

#Escribir el contenido del archivo 2 en el archivo 4

f.write("El contenido del archivo 2 es:\n")
f.write(str(matrizn))

#Escribir como una serie el contenido del archivo 1 en el archivo 3 

f.write("\n\nLa serie con el contenido del archivo 2 es:\n")
f.write(str(serie2))

#Escribir la serie con indices del contenido del archivo 1 en el archivo 3 

f.write("\n\nLa serie con el contenido del archivo 2 con sus respectivos nombres de celdas es:\n")
f.write(str(serie22))

#Crear una nueva serie

a = pandas.Series([4,7,1],index=["A4","A5","A6"]) 

#Escribir la serie a en el archivo 4

f.write("\n\nSe creó una serie nueva llamada a, la cual es la siguiente:\n")
f.write(str(a))

#Sumar la serie del archivo 2 con la serie a

d = pandas.concat([serie22,a])
print(d)

#Escribir la suma anterior en el archivo 4

f.write("\n\nSi sumamos la serie del archivo 2 con la nueva serie a obtenemos la siguiente serie:\n")
f.write(str(d))

#Cerrar y guardar los archivos 3 y 4

e.close()
f.close()