import pandas as pd
import pathlib
import win32com.client as win32
import smtplib
import email.message
from email.mime.text import MIMEText
from email.message import EmailMessage

emails = pd.read_excel(r'Bases de Dados\Emails.xlsx')
lojas =  pd.read_csv(r'Bases de Dados\Lojas.csv', encoding='latin1', sep=';')
vendas = pd.read_excel(r'Bases de Dados\Vendas.xlsx')

vendas = vendas.merge(lojas, on='ID Loja')

dicionario_lojas = {}
for loja in lojas['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja'] == loja, :] 
    
    
dia_indicador = vendas['Data'].max()

caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')
arquivo_pasta_backup = caminho_backup.iterdir()

lista_nomes_backup = [arquivo.name for arquivo in arquivo_pasta_backup]
    
for loja in dicionario_lojas:
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()    
    
    #Salvar dentro da pasta
    nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    local_arquivo = caminho_backup / loja / nome_arquivo
    
    dicionario_lojas[loja].to_excel(local_arquivo) 
    
meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qntdprodutos_dia = 4
meta_qntdprodutos_ano = 120
meta_ticketmedito_dia = 500
meta_ticketmedito_ano = 500

for loja in dicionario_lojas:
    
    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']== dia_indicador,:]

    #faturamento 
    faturamento_ano = vendas_loja['Valor Final'].sum()
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()

    #diversidade de produtos
    qntd_produtos_ano = len(vendas_loja['Produto'].unique())
    qntd_produtos_dia = len(vendas_loja_dia['Produto'].unique())

    #ticket médio
    valor_venda = vendas_loja.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_ano = valor_venda['Valor Final'].mean()

    # ticket medio dia
    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()
    
    if faturamento_dia >= meta_faturamento_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'
    if faturamento_ano >= meta_faturamento_ano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'
    

    if qntd_produtos_dia >= meta_qntdprodutos_dia:
        cor_qntd_dia = 'green'
    else:
        cor_qntd_dia = 'red'
    if qntd_produtos_ano >= meta_qntdprodutos_ano:
        cor_qntd_ano = 'green'
    else:
        cor_qntd_ano = 'red'
    
    if ticket_medio_dia >= meta_ticketmedito_dia: 
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia = 'red'
    if ticket_medio_ano >= meta_ticketmedito_ano:
        cor_ticket_ano = 'green'
    else:
        cor_ticket_ano = 'red'


    nome = emails.loc[emails['Loja'] == loja, 'Gerente'].values[0]
    para_pessoa = emails.loc[emails['Loja'] == loja, 'E-mail'].values[0]
    assunto = f'One page dia {dia_indicador.day}/{dia_indicador.month} Loja {loja}'
    corpo_email = f'''

    <h2>Bom dia, {nome}</h2>

    <p>o resultado de ontem ({dia_indicador.day},{dia_indicador.month}) da loja {loja} foi:</p>

    <p>segue a planilha com todos os dados para mais detalhes</p>

    <p>Qualquer duvida estou a disposição </p>
    <p>att., Hugo</p>


    <table>
    <tr>
        <th>indicador</th>
        <th>valor dia</th>
        <th>meta dia</th>
        <th>Cenario dia</th>
    </tr>
    <tr>
        <td> Faturamento </td>
        <td style="text-align: center">R$ {faturamento_dia:.2f}</td>
        <td style="text-align: center">R$ {meta_faturamento_dia:.2f}</td>
        <td style="text-align: center"><font color= "{cor_fat_dia}">◙</font></td>
    </tr>
    <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qntd_produtos_dia}</td>
        <td style="text-align: center">{meta_qntdprodutos_dia}</td>
        <td style="text-align: center"><font color=" {cor_qntd_dia}">◙</font></td>
    </tr>
    <tr>
        <td>Ticket médio</td>
        <td style="text-align: center">R$ {ticket_medio_dia:.2f}</td>
        <td style="text-align: center">R$ {meta_ticketmedito_dia:.2f}</td>
        <td style="text-align: center"><font color= "{cor_ticket_dia}">◙</font></td>
    </tr>
    </table>

    <br>

    <table>
    <tr>
        <th>indicador</th>
        <th>valor dia</th>
        <th>meta dia</th>
        <th>Cenario dia</th>
    </tr>
    <tr>
        <td> Faturamento </td>
        <td style="text-align: center">R$ {faturamento_ano:.2f}</td>
        <td style="text-align: center">R$ {meta_faturamento_ano:.2f}</td>
        <td style="text-align: center"><font color= "{cor_fat_ano}">◙</font></td>
    </tr>
    <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qntd_produtos_ano}</td>
        <td style="text-align: center">{meta_qntdprodutos_ano}</td>
        <td style="text-align: center"><font color=" {cor_qntd_ano}">◙</font></td>
    </tr>
    <tr>
        <td>Ticket médio</td>
        <td style="text-align: center">R$ {ticket_medio_ano:.2f}</td>
        <td style="text-align: center">R$ {meta_ticketmedito_ano:.2f}</td>
        <td style="text-align: center"><font color= "{cor_ticket_ano}">◙</font></td>
    </tr>
    </table>

    '''
    
    def enviar_email():
        corpo_email 
        
    msg = email.message.Message()
    msg['Subject'] = assunto
    msg['From'] = 'email_remetente'
    msg['To'] = 'email_destinatario'
    Password = 'senha app (do gmail)'
    msg.add_header ('content-type','text/html')
    msg.set_payload(corpo_email)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()

    s.login(msg['From'], Password)
    s.sendmail(msg['From'], [msg['To'],], msg.as_string().encode('UTF - 8'))

    print(f'E-mail enviado da loja {loja} com sucesso !')
    
faturamento_lojas = vendas.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_lojas_ano = faturamento_lojas.sort_values(by='Valor Final', ascending=False)


nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_Ranking_Anul.xlsx'
faturamento_lojas_ano.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo)) 


vendas_dia = vendas.loc[vendas['Data']==dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)


nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_Ranking_Dia.xlsx'
faturamento_lojas_dia.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo)) 

assunto = f'Resumo das Lojas Dia: {dia_indicador.day} / {dia_indicador.month}'
  


smtp_server = 'smtp.gmail.com'
smtp_port = 587
email_remetente = 'email_do_remetente'
senha = 'senha app (do gmail)'

email_destinatario = 'email_do_destinatario'

msg = EmailMessage()
msg['From'] = email_remetente
msg['To'] = email_destinatario
msg['Subject'] = assunto


arquivo_caminho1 = pathlib.PureWindowsPath(f'caminho_do_arquivo 1 \\{dia_indicador.month}_{dia_indicador.day}_Ranking_Dia.xlsx')
arquivo_caminho2 = pathlib.PureWindowsPath(f'caminho_do_aquivo 2 \\{dia_indicador.month}_{dia_indicador.day}_Ranking_Ano.xlsx')

arquivo_caminho1 = str(arquivo_caminho1)
arquivo_caminho2 = str(arquivo_caminho2)

arquivo_nome = f'{dia_indicador.month}_{dia_indicador.day}_Ranking_Dia.xlsx'

caminhos = [arquivo_caminho1 , arquivo_caminho2]

for caminho in caminhos:
    with open(caminho, 'rb') as arquivo:
        msg.add_attachment(arquivo.read(), maintype='application', subtype='octet-stream', filename=caminho)

try:
    with smtplib.SMTP(smtp_server, smtp_port) as servidor:
        servidor.starttls()
        servidor.login(email_remetente, senha)
        servidor.sendmail(email_remetente, email_destinatario, msg.as_string())
    print("E-mail enviador com sucesso!")
except Exception as e:
    print(f'Erro ao enviar e-mail: {e}')