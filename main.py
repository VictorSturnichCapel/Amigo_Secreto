# Esse projeto é destinado a elaboração do amigo secreto sem erro de um tirar o outro e acabar com o ciclo de entrega dos presentes, assim como a não utilização de plataformas

# Requer pacotes de instalação:
# pip install pandas
# pip install openpyxl

# importação bibliotecas
import pandas as pd
import random
import smtplib

# configurar email para desparo
# Usuário e senha para utilizar o email para enviar o aviso
email_envio = 'seu_email'
email_senha = "sua_senha"

# Enviar email para as pessoas
def envia_email(remetente_nome, remetente_email, destinatario_nome, destinatario_email, genero, presente):
    assunto = f"O seu amigo secreto foi sorteado!"
    if genero == 'masculino':
        textoEmail=f'Subject:'+assunto+f'\n\nOii, {remetente_nome}.\n\nO seu amigo secreto é o {destinatario_nome}!\n\nO presente que ele gostaria é o: '+presente+'\n\nAtenciosamente,\nBot Victor Capel'
    else:
        textoEmail=f'Subject:'+assunto+f'\n\nOii, {remetente_nome}.\n\nA sua amiga secreta é a {destinatario_nome}!\n\nO presente que ela gostaria é o: '+presente+'\n\nAtenciosamente,\nBot Victor Capel'

    textoEmail = textoEmail.encode('utf-8')
    server = smtplib.SMTP(host='smtp.gmail.com', port=587)
    server.ehlo()
    server.starttls()
    server.login(email_envio, email_senha)
    server.sendmail(email_envio, remetente_email, textoEmail)

# Gerar lista de participantes randomica
def montar_participantes(df):
    # lista dos participantes
    lista_participantes = list(df['Email'].unique())

    # total de participantes
    total_participante = len(lista_participantes)

    # contador para identificar os participantes
    i = 0

    # nova lista rândomica para sorteio
    lista_sorteada_participantes = []

    # gerar a lista de participantes onde cada participante irá presentear o participante a sua direita
    while i < total_participante:

        # gerar escolhido da lista rândomicamente
        escolhido = random.choice(lista_participantes)

        # adicionar na lista sorteada
        lista_sorteada_participantes.append(escolhido)

        # remover da lista dos participantes, para não se repetir
        lista_participantes.remove(escolhido)

        # aumentar contador
        i += 1
    
    return lista_sorteada_participantes

# Passar por cada participantes para enviar o email 
def preparando_envio(lista_sorteada_participantes,df):
    try:
        # passar por cada participantes para enviar o email 
        for num, i in enumerate(lista_sorteada_participantes):
            if num+1 < len(lista_sorteada_participantes):
                envia_email(df.loc[df['Email']==i,'Nome'].iloc[0],i,df.loc[df['Email']==lista_sorteada_participantes[num+1],'Nome'].iloc[0],lista_sorteada_participantes[num+1],df.loc[df['Email']==lista_sorteada_participantes[num+1],'gênero'].iloc[0],df.loc[df['Email']==lista_sorteada_participantes[num+1],'presente'].iloc[0])
            else:
                envia_email(df.loc[df['Email']==i,'Nome'].iloc[0],i,df.loc[df['Email']==lista_sorteada_participantes[0],'Nome'].iloc[0],lista_sorteada_participantes[0],df.loc[df['Email']==lista_sorteada_participantes[0],'gênero'].iloc[0],df.loc[df['Email']==lista_sorteada_participantes[0],'presente'].iloc[0])
    except Exception as e:
        print(f'Erro: {e}')

if __name__ == "__main__":

    try:
        # importar dados dos participantes
        df = pd.read_excel('participantes.xlsx')

        # gerar lista de participantes randomica
        lista_sorteada_participantes = montar_participantes(df)

        # Se a lista de pessoas adicionados for maior que 1, continua o processo
        if len(lista_sorteada_participantes) > 1:
            # passar por cada participantes para enviar o email 
            preparando_envio(lista_sorteada_participantes,df)
            print('Sorteio finalizado')
        else:
            print('Adicione mais de uma pessoa')
        
    except Exception as e:
        print(f'Erro: {e}')