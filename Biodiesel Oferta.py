import pandas as pd
import os
import numpy as np
import plotly.express as px



base = pd.read_excel('2021-12_entregas-biodiesel-usinas-produtoras.xlsx', header=3)

del base['Unnamed: 0']
base.drop(0, inplace=True)
base['Empresa'] =  base['Usina Produtora de Biodiesel'].str.split(' -').str[0]

table = base[base.columns[4:]]
table = table.groupby('Empresa', as_index=False).sum()
table['Acumulado 2021'] = table.sum(axis=1, numeric_only=True)
table['1ºTri'] = table['Janeiro']+table['Fevereiro']+table['Março']
table['2ºTri'] = table['Abril']+table['Maio']+table['Junho']
table['3ºTri'] = table['Julho']+table['Agosto']+table['Setembro']
table['4ºTri'] = table['Outubro']+table['Novembro']+table['Dezembro']
table['Market Share 2021'] = (table['Acumulado 2021'] / table['Acumulado 2021'].sum())
table = table[['Empresa', '1ºTri', '2ºTri', '3ºTri','4ºTri','Acumulado 2021', 'Market Share 2021']]
table.sort_values(by='Acumulado 2021', inplace=True, ascending=False)


# fig = px.treemap(data_frame=table,
#                  path=['Empresa'],
#                  values='Acumulado 2021',
#                  title='Market Share 2021')
# fig.update_layout(
#                    title=dict(text="Market Share 2021",
#                               x=0.5),
#                    title_font= dict(family = 'Arial', size = 35),
#                    font = dict(size = 25, family = 'Verdana'),
#                    hovermode = False,
#                    width  = 1400, height = 1400)
#
# fig.show()


writer = pd.ExcelWriter('biodiesel.xlsx', engine='xlsxwriter')
base.to_excel(writer, sheet_name='base', index=False, header=True, startrow=0)
table.to_excel(writer, index=False, sheet_name='table', header=True, startrow=0)
writer.save()
os.startfile('biodiesel.xlsx')