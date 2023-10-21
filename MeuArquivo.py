# DATA BASE DE VENDAS EM SHOPPINGS LOCAIS
#OBJETIVOS PRIMA:
# 1-Obter o faturamento por loja,
# 2-saber a quantidade de produtos vendidos por loja
# 3-saber o ticket medio por produdo em cada loja.
# (cacular: dividir o faturamento pela quantidade media vendida para obte
# o ticket)
#OBJETIVO SEC: ENVIAR UM EMAIL COM O RELATORIO

import pandas as pd
import win32com.client as win32
# leitura de tabela
tabela_vendas = pd.read_excel('Vendas.xlsx')
pd.set_option('display.max_columns',None) #verificação de leitura correnta do arquivo em questão.
print(tabela_vendas)
# faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('=' * 50)
# ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

#OBJETIVO SEC:

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'hinxoxo@gmail.com' # email do destinatario
mail.Subject = 'Relatorio de Teste' #assunto do email
mail.HTMLBudy =f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>Ari Santos</p>
'''

mail.Send()

print('Email Enviado')