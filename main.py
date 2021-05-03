# Sofía Múnera Medina
# Python
# POO con archivos .xslx
# Fuentes:


# EJEMPLO 1

# Importamos la libreria pandas para poder tratar archivos de Excel

import pandas as pd

# ExcelWriter lo usaremos para crear un nuevo archivo .xlsx

from pandas import ExcelWriter

# Importamos openpyxl para leer un archivo .xlsx

import openpyxl


# Creamos la clase pais

class Pais:

  # Documentacion 

  ''' Lee la informacion de un archivo .xls, la modifica y crea otro donde almacena los cambios '''

  # Mediante __init__ crearemos los atributos de instancia 

  def __init__(self):

    # Documentacion 

    ''' Metodo constructor de la clase Pais, establece 3 atributos de instancia con valores predefinidos '''

    # Creamos los atributos de instancia, sus valores serán cadenas de texto que representan 3 paises 

    self.pais0='España'
    self.pais1='Estados Unidos'
    self.pais2='China'
  
  # Creamos un metodo llamado de la clase pais llamado modificar 

  def modificar(self):

    # Documentacion 

    ''' Lee un archivo .xlsx, almacena la informacion que contiene y luego crea un nuevo archivo con esta informacion y otra adicional '''

    # Usando la libreria openpyxl y la funcion load_workbook vamos a leer el archivo que usaremos de base llamado Libro1.xlsx

    libro1=openpyxl.load_workbook('Libro1.xlsx')

    # Ahora vamos a usar la Hoja1 del libro que acabamos de leer 

    hoja1=libro1['Hoja1']

    # Mediante Hoja1[celdas] vamos a leer los datos que usaremos en el archivo nuevo. De A2:A4 hay una lista de 3 capitales y de B2:B4 estan los idiomas que hablan en cada una de estas ciudades

    capitales=hoja1['A2':'A4']
    idiomas=hoja1['B2':'B4']


    # Ahora comenzaremos a crear un nuevo libro

    # Creamos una lista con los atributos de instancia

    paises=[self.pais0,self.pais1,self.pais2]

    # Mediante DataFrame crearemos una tabla donde las cabeceras de cada columna son Pais,Capital e Idioma y a cada uno de estos encabezados le vamos a asignar unos datos. En el caso de Capital e Idioma los datos serán los que leimos del archivo original (Libro1) y para Pais se le signa la lista que contiene los valores de los atributos de instancia 

    copia1=pd.DataFrame({'Pais':paises,'Capital':capitales,'Idioma':idiomas})

    # Crearemos una nueva lista con informacion de continentes 

    continentes=['Europa','América','Asia']

    # Insertaremos una columna con encabezado Continente y valores de la lista continentes. Esta irá en la posicion/locacion 2, lo que significará que la columna Idioma se desplazará una a la derecha 

    # el orden de columnas será Pais, Capital, Continente y por ultimo Idioma

    copia1.insert(2,'Continente',continentes,allow_duplicates=False)

    # mediante Excel writer crearemos un archivo llamado Copia1.xlsx

    archivo=ExcelWriter('Copia1.xlsx')

    # Ahora con ayuda de .to_excel el Dataframe que creamos (incluyendo la informacion insertada posteriormente) lo guardaremos en este nuevo archivo (Copia1), la hoja se llamara Hoja Copia y podemos index=False para que no se incluya encabezado de numeros 

    copia1.to_excel(archivo,'Hoja Copia',index=False)

    # Guardamos lo ingresado al archivo copia 1

    archivo.save()

    # Cerramos el archivo 

    archivo.close()

# Llamamos la clase Pais, que inmediantamente creará los 3 atributos de instancia 

a=Pais()

# Mediante el metodo modificar vamos a leer la información de Libro1 y crearemos un nuevo archivo llamado Copia 1 

a.modificar()

# FINALIZA EJEMPLO 1


#