#Lógica da Análise

#pandas- biblioteca para trabalhar com tabelas
import pandas as pd

#Importar a base de dados
tabela_vendas = pd.read_excel("Vendas.xlsx")

#Visualizar a base
pd.set_option('display.max.columns',None)
print (tabela_vendas)
print ('-' * 50)

#Faturarmento por Lojas
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print (faturamento)
print ('-' * 50)

#Qtde de produtos vendidos por loja
qtde_produto = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print (qtde_produto)
print ('-' * 50)

#Ticket médio por produto em cada loja (faturamento/qtde de produtos)
ticket_medio = (faturamento['Valor Final']/qtde_produto['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print (ticket_medio)


#Enviar por e-mail o relatório
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail=outlook.CreateItem(0)
mail.To= 'carolinaaloisio.99@gmail.com'
mail.Subject = 'Teste Relatório de Vendas'
mail.HTMLBody = f'''
<p> Prezados, </p>

<p> Segue o relatório de vendas por loja: </p>

<p> Faturamento: </p>
{faturamento.to_html(formatters={'Valor Final':'R${:,.2f}'.format})}

<p> Quantidade de produto vendida: </p>
{qtde_produto.to_html()}

<p> Ticket Médio por Produto e Loja: </p>
{ticket_medio.to_html(formatters={'Ticket Médio':'R${:,.2f}'.format})}

<p> Att, </p>
Carol.

'''
mail.Display()
print('E-mail enviado com sucesso')
