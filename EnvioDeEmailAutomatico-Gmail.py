import numpy as np
import pandas as pd
import yagmail #pip install yagmail
from beautifultable import BeautifulTable
from pathlib import Path
from meusdados import MeusDados


table = BeautifulTable(maxwidth=140)
table.set_style(BeautifulTable.STYLE_BOX)


meu_email = MeusDados()
minha_senha = MeusDados(2) #A senha que será colocada aqui tem q ser criada através do App Password da google: https://myaccount.google.com/apppasswords (OBS: a conta deve conter a verificação de duas etapas para prosseguir)
usuario = yagmail.SMTP(user=meu_email,password=minha_senha)
planilha_df = pd.read_excel(r'C:\Users\Fulano\Pasta\AlgumArquivo.xlsx')

print(meu_email)


contato_df = planilha_df[['Aluno','Email Aluno','Celular']]
contato_df = contato_df.dropna(how='any')
contato_df = contato_df.drop_duplicates()
contato_df['Aluno'] = [each.title() for each in contato_df['Aluno']]
contato_df['Email Aluno'] = [each.strip() for each in contato_df['Email Aluno']]
contato_df = contato_df.sort_values(by='Aluno')


table.columns.header = ['','Nome','Email','Contato','Situação']


for e, email in enumerate(contato_df['Email Aluno']):
    aluno = contato_df.iloc[e, 0]
    numero = contato_df.iloc[e, 2]

    assunto_email ='Algum Assunto'

    html = f"""
    <h1>Prezado(a), {aluno}</h1>
    Já está sabendo da novidade que está rolando neste mês de agosto!?
    Não perca tempo!
    Atente-se à imagem e clique no link para saber mais!
    
    Atenciosamente, 
    Empresa X

    <p><b><a href='https://wa.me/message/00000000000000'>Para mais informações, clique aqui!</a></b></p>

    """
    img = yagmail.inline("AlgumaImagem.jpeg")
    
    try:
      usuario.send(
              to=email,
              subject=assunto_email,
              contents=[html, img]) 
      
      table.rows.append([f'{e+1}º', aluno, email, numero,'Email Entregue'])
  
    except:
      table.rows.append([f'{e+1}º', aluno, email, numero,'Email Não Entregue'])


print(table)
    
