import win32com.client as win32com
import pandas as pd 
from datetime import datetime 
import psycopg2 

conexao = psycopg2.connect(
    host='',
    database="",
    user="",
    password="",
    port =""
)

if conexao:
    print("Conexão ao BD V2 foi bem-sucedida.")

query = '''
    select
    c."name",
    p.id,
    p.purchase_date,
    t.nsu,
    i.bar_code_captured,
    i.amount::numeric / 100 as Valor
from
    purchase p
join purchase_transactions pt on
    pt.purchase_id = p.id
join "transaction" t on
    pt.transactions_id = t.id
join purchase_invoices pi2 on
    pi2.purchase_id = p.id
join invoice i on
    pi2.invoices_id = i.id
join company c on
    p.company_id = c.id
where
    p.purchase_channel = 'POS'
    and p.purchase_date::date between '2023-11-01' and '2023-11-02'
    and p.company_id in ('22', '23', '26')
    and p.status in ('1', '2')
    and i.return_file_uuid is null
order by
    p.id desc;'''

print("Query ENEL realizada.")

query1 = '''
select
    c."name",
    p.id,
    p.purchase_date,
    t.nsu,
    i.bar_code_captured,
    i.amount::numeric / 100 as Valor
from
    purchase p
join purchase_transactions pt on
    pt.purchase_id = p.id
join "transaction" t on
    pt.transactions_id = t.id
join purchase_invoices pi2 on
    pi2.purchase_id = p.id
join invoice i on
    pi2.invoices_id = i.id
join company c on
    p.company_id = c.id
where
p.purchase_date::date between '2023-11-01' and '2023-11-02'
    and p.company_id = '103'
    and p.status in ('1', '2')
order by
    p.id desc;
    '''

print("Query TJRO realizada.")

dfenel = pd.read_sql(query, conexao)
dftjro = pd.read_sql(query1, conexao)
conexao.close()

# Criar um objeto ExcelWriter usando o Pandas
nome_arquivo = 'ENEL_POS.xlsx'
dfenel.to_excel(nome_arquivo, sheet_name='Planilha enel', index=False)
nome_arquivo1 = 'TJRO.xlsx'
dftjro.to_excel(nome_arquivo1, sheet_name='Planilha tjro', index=False)

de = "@gmail.com"
para = "@gmail.com;@gmail.com"
senha = ""
data_atual = datetime.now().strftime('%d-%m-%Y')

outlook = win32.Dispatch('Outlook.Application')
mail = outlook.CreateItem(0)  

mail.Subject = f"ENEL POS Ficha - {data_atual}"  # Assunto personalizado com a data

# Corpo do e-mail 
mensagem = """
Prezados, bom dia.

Segue em anexo transações ENEL POS conforme acordado.

Qualquer dúvida estamos à disposição.

Atenciosamente,


"""
mail.Body = mensagem

print("E-mail sendo digitado.")

# Anexar a planilha ao e-mail
attachment = r"C:\\CAMINHO\\DO\\ARQUIVO\\ENEL_POS.xlsx"
mail.Attachments.Add(attachment)

# Destinatário do e-mail
mail.To = para

# Enviar o e-mail
mail.Send()

print("E-mail enviado com sucesso.")

mail = outlook.CreateItem(0)  # 0 significa que estamos criando um novo e-mail

mail.Subject = f"TJRO Ficha - {data_atual}" 

# Verifica se o DataFrame dftjro está vazio
if dftjro.empty:
    mensagem = """
    Prezados, bom dia.

    Não tivemos transações TJRO no período solicitado.

    Qualquer dúvida estamos à disposição.

    Atenciosamente,

    """
else:
    # Caso contrário, mantém a mensagem original
    mensagem = """
    Prezados, bom dia.

    Segue em anexo transações TJRO conforme acordado.

    Qualquer dúvida estamos à disposição.

    Atenciosamente,

    """
    
mail.Body = mensagem

print("E-mail sendo digitado.")

# Anexar a planilha ao e-mail
attachment = r"C:\\CAMINHO\\DO\\ARQUIVO\\TJRO.xlsx"
mail.Attachments.Add(attachment)

mail.To = para
mail.Send()

print("E-mail enviado com sucesso.")
