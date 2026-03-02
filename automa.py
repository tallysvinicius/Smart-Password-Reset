import win32com.client
import subprocess
import json
import ollama
from datetime import datetime


# ========================================
# CONECTAR AO OUTLOOK E LISTAR EMAILS
# ========================================

def selecionar_email():

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)

    emails = inbox.Items

    # força atualização
    emails = emails.Restrict("[MessageClass] = 'IPM.Note'")

    emails.Sort("[ReceivedTime]", True)

    lista = []

    print("\nEmails disponíveis:\n")

    contador = 0

    for email in emails:

        try:

            assunto = email.Subject
            remetente = email.SenderName
            data = email.ReceivedTime.strftime("%d/%m/%Y %H:%M")

            print(f"{contador} - {assunto} | {remetente} | {data}")

            lista.append(email)

            contador += 1

            if contador >= 22:
                break

        except:
            pass

    escolha = int(input("\nEscolha o número do email: "))

    return lista[escolha]


# ========================================
# IA COM OLLAMA PARA EXTRAIR NOME
# ========================================

def extrair_nome_com_ia(email):

    corpo = email.Body

    prompt = f"""
Você é um assistente de TI.

Analise o email abaixo e extraia SOMENTE o nome completo da pessoa que precisa ter a senha redefinida.

Regras:

- Retorne APENAS o nome completo
- Não explique nada
- Não escreva frases
- Não escreva pontuação extra
- Se não encontrar, retorne: NAO_ENCONTRADO

Email:
{corpo}
"""

    try:

        resposta = ollama.chat(
            model="llama3",
            messages=[{"role": "user", "content": prompt}]
        )

        nome = resposta["message"]["content"].strip()

        if nome == "NAO_ENCONTRADO" or len(nome) < 3:
            return email.SenderName

        return nome

    except Exception as e:

        print("Erro IA:", e)

        return email.SenderName


# ========================================
# BUSCAR USUARIO NO ACTIVE DIRECTORY
# ========================================

def buscar_usuario_por_nome(nome):

    comando = f"""
    Get-ADUser -Filter "Name -like '*{nome}*'" |
    Select Name,SamAccountName |
    ConvertTo-Json
    """

    resultado = subprocess.run(

        ["powershell", "-Command", comando],
        capture_output=True,
        text=True,
        encoding="latin-1"
    )

    if not resultado.stdout.strip():
        return []

    dados = json.loads(resultado.stdout)

    if isinstance(dados, dict):
        return [dados]

    return dados


# ========================================
# ESCOLHER USUARIO
# ========================================

def escolher_usuario(usuarios):

    if not usuarios:

        print("\nNenhum usuário encontrado no AD")

        return None

    print("\nUsuários encontrados:\n")

    for i, user in enumerate(usuarios):

        print(f"{i} - {user['Name']} ({user['SamAccountName']})")

    escolha = int(input("\nEscolha o número correto: "))

    return usuarios[escolha]["SamAccountName"]


# ========================================
# RESETAR SENHA
# ========================================

def resetar_senha(login):

    nova_senha = login[:3].lower() + "012@"

    comando = f"""
    Set-ADAccountPassword -Identity "{login}" -Reset -NewPassword (ConvertTo-SecureString "{nova_senha}" -AsPlainText -Force)
    Unlock-ADAccount -Identity "{login}"
    """

    subprocess.run(

        ["powershell", "-Command", comando],
        capture_output=True,
        text=True
    )

    print("\nSenha redefinida com sucesso!")
    print("Login:", login)
    print("Nova senha:", nova_senha)

    return nova_senha


# ========================================
# RESPONDER EMAIL
# ========================================

def responder_email(email, login, senha):

    resposta = f"""
Olá,

A senha foi redefinida com sucesso.

Login: {login}
Senha: {senha}

Solicitamos que altere a senha após o primeiro acesso.

Atenciosamente,
Suporte TI
"""

    reply = email.Reply()

    reply.Body = resposta

    reply.Send()

    print("\nEmail respondido com sucesso.")


# ========================================
# FLUXO PRINCIPAL
# ========================================

def main():

    email = selecionar_email()

    print("\nAnalisando email com IA...")

    nome = extrair_nome_com_ia(email)

    print("Nome identificado:", nome)

    usuarios = buscar_usuario_por_nome(nome)

    login = escolher_usuario(usuarios)

    if login:

        senha = resetar_senha(login)

        responder_email(email, login, senha)


# ========================================

main()