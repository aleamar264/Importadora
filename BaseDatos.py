# #-Falta hacer un for o un while para recorrer todas la filas y exportar todos los datos
# a excel.
# #-Falta partir los datos de cada fila ya que en cada fila se encuentran mas de un producto.

## Se importan las librerias necesarias para leer y escribir en Excel
import pandas as pd
import numpy as np
## Se carga la base de Datos y se separan para su manejo
Base=pd.read_excel('MuestraDatos.xlsx')
Base = Base.values
Declaracion =Base[0::1][:, 0:1]
Declaracion = np.array(Declaracion)
Otro = Base[0::1][:, 1:2]
aux = [0,0,0,0,0,0,0,0,0,0]
# print(Prueba)
# #Se crea un Diccionario con las palabras mas relevantes de la baseDatos
Diccionario = ['MARCA','REF','DIM','D.O.','CANT','HECHO EN', 'FORMATO' , 'PEDIDO','PRODUCTO:', 'PRODUCTO=']
##Cada una de las posiciones del Diccionario estan guardadas en la var Pos
##Los valores con -1 no se encuentran dentro del texto
Pos_ini = [[0] * len(Diccionario) for i in range(len(Otro))]
Pos_Fin = [[0] * len(Diccionario) for i in range(len(Otro))]
Final = [[0] * len(Diccionario) for i in range(len(Otro))]
## Si find no encuentra alugna de las palabras en el Diccionario el resultado es -1
## para efectos de control se le asigna un numero alto a los valores de -1

for a in range (1,len(Base)+1):
    Prueba = str(Otro[a-1])
    for i in range (1,len(Diccionario)+1):
        Pos_ini[a-1][i-1]=Prueba.find(Diccionario[i-1])
        if Pos_ini [a-1][i-1]==-1:
            Pos_ini[a-1][i-1] = 20000
    for j in range (1,len(Diccionario)+1):
        for k in range (1,len(Diccionario)+1):
            aux[k-1]=Pos_ini[a-1][j-1]-Pos_ini[a-1][k-1]
            # if aux[k]<=0:
            #     aux[k] = 20000
        aux[j-1]=20000
        Pos_Fin[a-1][j-1]=min(abs(np.array(aux)))
    for i in range (1,len(Diccionario)+1):
        if Pos_ini[a-1][i-1]== 20000:
            Pos_ini[a-1][i-1] = -1
    Max = max(Pos_ini[a-1])
    Pos_Max = Pos_ini[a-1].index(Max)
    Pos_Fin[a-1][Pos_Max] = -1
    for h in range (1,len(Diccionario)+1):
        Final[a-1][h-1] = Prueba[Pos_ini[a-1][h-1]+len(Diccionario[h-1]):Pos_Fin[a-1][h-1]+Pos_ini[a-1][h-1]]
        if Pos_Fin[a-1][h-1] == -1:
            Final[a-1][h-1] = Prueba[Pos_ini[a-1][h-1]+len(Diccionario[h-1]):Pos_Fin[a-1][h-1]]

## Se importa la libreria i0 para poder exportar datos tipos string en pandas
import io
Final= np.array(Final)
output = io.BytesIO()
writer = pd.ExcelWriter(output, engine='xlsxwriter')
## Se Pasan los datos encontrados a una hoja de Excel para su visualizacion
writer = pd.ExcelWriter('baseDatos.xlsx', engine='xlsxwriter')
for i in range(1,len(Diccionario)+1):
    dfM = pd.DataFrame({Diccionario[i-1]:Final[:,i-1]})
    dfM.to_excel(writer, sheet_name='Sheet1',startcol=i , index=False)
df = pd.DataFrame({'Declaracion': Declaracion[:,0]})
df.to_excel(writer, sheet_name='Sheet1',startcol=0 , index=False)
workbook  = writer.book
worksheet = writer.sheets['Sheet1']
format1 = workbook.add_format({'num_format': '#'})
worksheet.set_column('A:A', 16, format1)
writer.save()
