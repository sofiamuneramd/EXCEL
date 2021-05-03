# Sofía Múnera Medina
# Python
# POO con archivos .xslx
# Fuentes:


import pandas as pd
from pandas import ExcelWriter


class Excel:

  def __init__(self):

    self.pais0='España'
    self.pais1='Estados Unidos'
    self.pais2='China'
  
  def leer(self):

    libro1=pd.ExcelFile('Libro1.xls' )

    hoja1=libro1.parse('Hoja1')

    paises=[self.pais0,self.pais1,self.pais2]

    capitales=hoja1['CAPITAL'].values

    idiomas=hoja1['IDIOMA'].values

    copia1=pd.DataFrame({'Pais':paises,'Capital':capitales,'Idioma':idiomas})

    continentes=['Europa','América','Asia']

    copia1.insert(2,'Continente',continentes,allow_duplicates=False)

    archivo=ExcelWriter('copia1.xls')
    copia1.to_excel(archivo,'Hoja Copia',index=False)
    archivo.save()
    archivo.close()


a=Excel()
a.leer()