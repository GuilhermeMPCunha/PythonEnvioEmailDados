import pandas as pd
import win32com.client as win32

# importar a base de dados

TabelaVendas = pd.read_excel('Vendas.xlsx')

# visualizar a base de dados

#pd.set_option('display.max_columns',None)
#print(TabelaVendas)

# faturamento por loja

faturamento = TabelaVendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

print(faturamento)

# quantidade de produtos por loja

print('-' * 50)

FatQtd = TabelaVendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

print(FatQtd)


# ticket medio por produto

print('-' * 50)

TckMed = (faturamento['Valor Final'] / FatQtd['Quantidade']).to_frame()

print(TckMed)

# enviar um email com o relatório

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'guilhermempcunha10@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = '''

Prezados, 

Segue o Relatório de Vendas por cada loja

Faturamento:
{}

Quantidade Vendida:
{}

Ticket Medio dos produtos em cada Loja:
{}

Qualquer dúvida estou à disposição

att
João da Silva

'''

mail.Send()

print('-' * 50)

print('Email Enviado')