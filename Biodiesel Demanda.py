import pandas as pd
import os
import numpy as np




base = pd.read_excel('2021-12_retiradas-biodiesel-distribuidoras.xlsx', header=3)

del base['Unnamed: 0']
base.drop(0, inplace=True)

table = base
table.drop('RaizCNPJ', inplace=True, axis=1)
table['Acumulado 2021'] = table.sum(axis=1, numeric_only=True)
table['1ºTri'] = table['Janeiro']+table['Fevereiro']+table['Março']
table['2ºTri'] = table['Abril']+table['Maio']+table['Junho']
table['3ºTri'] = table['Julho']+table['Agosto']+table['Setembro']
table['4ºTri'] = table['Outubro']+table['Novembro']+table['Dezembro']
table['%'] = (table['Acumulado 2021'] / table['Acumulado 2021'].sum())
table = table[['Distribuidora', '1ºTri', '2ºTri', '3ºTri','4ºTri','Acumulado 2021', '%']]
table.sort_values(by='Acumulado 2021', inplace=True, ascending=False)



writer = pd.ExcelWriter('demanda.xlsx', engine='xlsxwriter')
base.to_excel(writer, sheet_name='base', index=False, header=True, startrow=0)
table.to_excel(writer, index=False, sheet_name='table', header=True, startrow=0)
writer.save()
os.startfile('demanda.xlsx')