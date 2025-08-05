import pandas as pd
import win32com.client as win32

# importar base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar e se precisar tratar bd
# mostrar todas as colunas sem ocultar
pd.set_option('display.max_columns', None)
print(tabela_vendas)
print('-' * 50)
# faturamento por loja

faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-' * 50)
# quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print('-' * 50)
# ticket medio por quantidade de produto vendido por loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# enviar email com relatorio

outlook = win32.Dispatch('Outlook.Application')
mail = outlook.CreateItem(0)
mail.To = 'emailbacana@teste.com'
mail.Subject = 'Projeto Pandas Python - Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}


<p>Quantidade Vendida:</p>
{quantidade.to_html}


<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>


<p>Att.,</p>
<p>Nome</p>

'''

mail.Send()