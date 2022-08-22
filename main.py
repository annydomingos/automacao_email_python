# para ler db em excel:
import pandas as pd
import win32com.client as win32

#Importando database
tabela_vendas = pd.read_excel('Vendas.xlsx')


# Visualizando database completa
pd.set_option('display.max_columns', None)


# Calcular faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-' * 50)

# Calcular Quantidade de produtos em cada loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print('-' * 50)

# Calcular Ticket médio por produto em cada loja (faturamento/quantidade produto)
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0 : 'Ticket Médio'})
print(ticket_medio)
print('-' * 50)

#Enviar email com relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'annyds1@icloud.com'
mail.Subject = 'Relatório de vendas por loja'
mail.HTMLBody = f'''
<h2>Prezados,</h2>

<p>Segue o relatório de vendas por loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters= {'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade de itens vendidos:</p>
{quantidade.to_html(formatters= {'Quantidade':'R${:,.2f}'.format})}

<p>Ticket médio:</p>
{ticket_medio.to_html(formatters={'Ticket Médio' :'R${:,.2f}'.format})}

<p>Quaisquer dúvidas estou à disposição.</p>
<p>Atenciosamente,</p>
<p>Anny Domingos</p>
'''

mail.Send()
