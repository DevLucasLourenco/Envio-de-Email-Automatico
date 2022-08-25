email_dados = ('FulanoDeTal@gmail.com')
senha_dados = ('SenhaDoFulano') #A senha que será colocada aqui tem q ser criada através do App Password da google: https://myaccount.google.com/apppasswords (OBS: a conta deve conter a verificação de duas etapas para prosseguir)
def MeusDados(padrao=1):
    '''
    selecionar 1 para email, 2 para senha.
    '''
    if padrao==1:
        return email_dados
    elif padrao==2:
        return senha_dados

if __name__ == '__main__':
    print(MeusDados(2))

    