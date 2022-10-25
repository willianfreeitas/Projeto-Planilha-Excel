import pandas as pd
import win32com.client as win32


# Importar a base de dados
from pip._vendor.colorama import win32

tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados
pd.set_option('display.max_columns', None)

# Faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# Quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

# Ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})

"""
# Enviar e-mails via Gmail
import smtplib
import email.message

def enviar_email():
    corpo_email = “””
    <p>Parágrafo1</p>
    <p>Parágrafo2</p>
    “””

    msg = email.message.Message()
    msg[‘Subject’] = “Assunto”
    msg[‘From’] = ‘remetente’
    msg[‘To’] = ‘destinatario’
    password = ‘senha’
    msg.add_header(‘Content-Type’, ‘text/html’)
    msg.set_payload(corpo_email )

    s = smtplib.SMTP(‘smtp.gmail.com: 587’)
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg[‘From’], password)
    s.sendmail(msg[‘From’], [msg[‘To’]], msg.as_string().encode(‘utf-8’))
    print(‘Email enviado’)
"""

# Enviar e-mails via Outlook ( A função Dispatch funciona apenas em windows, por isso o erro. )
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'willfreitaspaula@hotmail.com'
mail.Subject = 'Relatório de vendas por loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f'.format})}

<p>Qualquer dúvida estou a disposição.</p>

<p>Atenciosamente,</p>
<p>Willian Freitas.</p>
'''

mail.Send()


