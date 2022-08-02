import win32com.client as win32
import pandas as pd
from datetime import date

grupo = str(input('\nQual grupo de empresas executar?\n\ncorrente\nteste\n\nGrupo: '))

if grupo == 'corrente':
    xet = pd.read_excel('Planilha.xlsx')
elif grupo == 'teste':
    xet = pd.read_excel('Teste.xlsx')
else:
    print('\nGrupo inexistente!')

planilha = xet

outlook = win32.Dispatch('Outlook.Application')

emissor = outlook.session.Accounts['felipeofragoso@gmail.com']

dados = planilha[['Empresa', 'E-mail']].values.tolist()

hoje = date.today()

data = '{}/{}'.format(hoje.month, hoje.year)

for dado in dados:
    mensagem = outlook.CreateItem(0)
    mensagem.To = dado[1]
    mensagem.Subject = 'Comunicado para '+dado[0]+' ('+data+')'
    mensagem.CC = 'felipeofragoso@gmail.com'
    mensagem.HTMLBody = '''
<div style="
    width: 400px;
    height: 80px;

    display: flex;
    align-items: center;
    justify-content: center;

    border-radius: 50px;
    background-color: #eee;
    box-shadow: 10px 5px 50px rgba(0, 0, 0, 0.514);
">
    <h1 style="
    margin-top: 10px;

    text-align: center;
    font-weight: 100;
    font-family: Arial, Helvetica, sans-serif;
    ">E-mail Automatizado ✌😎</h1>
</div>

<h3>Dado da planilha coluna 1: </h3>{}
<h3>Dado da planilha coluna 2: </h3>{}
    '''.format(dado[0], dado[1])

    mensagem.Save()
    mensagem.Send()
