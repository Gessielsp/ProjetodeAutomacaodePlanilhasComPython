import win32com.client as win32
from pathlib import Path
import re
import os

# primeira fase do projeto, entando no outlook e pegando todos os email por assunto e baxaindo seus anexos

import win32com.client as win32
from pathlib import Path
import re
from datetime import datetime

import win32com.client as win32
from pathlib import Path
import re
import hashlib
from datetime import datetime

# Função para sanitizar o nome do diretório e arquivos
def sanitize_filename(filename):
    return re.sub(r'[<>:"/\\|?*]', '', filename)

# Função para calcular o hash SHA-256 de um arquivo
def calculate_file_hash(file_path):
    hash_sha256 = hashlib.sha256()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_sha256.update(chunk)
    return hash_sha256.hexdigest()

# Define o local específico onde a pasta principal será criada
local_especifico = Path("C:/Users/gessiel.passos/Documents/Anexos_Outlook")  # Altere este caminho para onde deseja salvar os arquivos

# Cria a pasta principal
destino = local_especifico / 'Emails_Anexos'
destino.mkdir(parents=True, exist_ok=True)
print(f"Pasta principal criada em: {destino}")

# Define os assuntos fixos
assuntos_fixos = {
    "001": "ENC: Contingenciamento Emendas de Bancada RP 7",
    "002": "Assunto Opção 002",
    "003": "Assunto Opção 003",
    "004": "Assunto Opção 004"
}

# Exibe as opções para o usuário
print("Escolha um dos assuntos abaixo:")
for chave, assunto in assuntos_fixos.items():
    print(f"{chave}: {assunto}")

# Solicita ao usuário o número da opção desejada
opcao = input("Digite o número da opção desejada: ")

# Verifica se a opção é válida
if opcao not in assuntos_fixos:
    print("Opção inválida. Encerrando o script.")
    exit()

# Define o assunto desejado com base na opção do usuário
assunto_desejado = assuntos_fixos[opcao]

# Inicializa o Outlook
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Acessando a Caixa de Entrada
inbox = outlook.GetDefaultFolder(6)  # 6 é o código da Caixa de Entrada padrão

# Arquivo para registrar e-mails processados e hashes de anexos
arquivo_processados = destino / 'emails_processados.txt'
hashes_processados = destino / 'hashes_processados.txt'
emails_processados = set()
hashes_armazenados = set()

# Carrega o registro de e-mails processados e hashes armazenados, se existir
if arquivo_processados.exists():
    with open(arquivo_processados, 'r') as f:
        emails_processados = set(f.read().splitlines())

if hashes_processados.exists():
    with open(hashes_processados, 'r') as f:
        hashes_armazenados = set(f.read().splitlines())

# Iterando sobre as mensagens na Caixa de Entrada
for message in inbox.Items:
    try:
        subject = message.Subject
        entry_id = message.EntryID
        received_time = message.ReceivedTime.strftime("%Y-%m-%d_%H-%M-%S")  # Formata a data e hora

        # Verifica se o e-mail já foi processado ou se o assunto não é o desejado
        if entry_id in emails_processados or assunto_desejado not in subject:
            continue

        attachments = message.Attachments

        # Criando uma subpasta para cada assunto, dentro da pasta principal
        pasta_de_destino = destino / sanitize_filename(subject)
        pasta_de_destino.mkdir(parents=True, exist_ok=True)
        print(f"Pasta criada para o assunto '{subject}': {pasta_de_destino}")

        # Salvando o e-mail inteiro no formato .msg
        email_filename = sanitize_filename(subject) + f"_{received_time}.msg"
        try:
            message.SaveAs(str(pasta_de_destino / email_filename), 3)  # 3 é o código para formato .msg
            print(f"E-mail salvo como {email_filename} em {pasta_de_destino}")
        except Exception as e:
            print(f"Erro ao salvar o e-mail '{subject}': {e}")

        # Salvando os anexos e verificando se são novos ou duplicados
        for attachment in attachments:
            attachment_filename = sanitize_filename(attachment.FileName)
            attachment_path = pasta_de_destino / attachment_filename

            # Salva o anexo
            try:
                attachment.SaveAsFile(attachment_path)
                print(f'Anexo {attachment_filename} salvo em {pasta_de_destino}')
            except Exception as e:
                print(f"Erro ao salvar o anexo '{attachment_filename}': {e}")
                continue

            # Calcula o hash do anexo
            attachment_hash = calculate_file_hash(attachment_path)
            if attachment_hash in hashes_armazenados:
                print(f"Anexo {attachment_filename} é duplicado. O arquivo já foi processado anteriormente.")
            else:
                hashes_armazenados.add(attachment_hash)

        # Adiciona o e-mail ao registro de processados
        with open(arquivo_processados, 'a') as f:
            f.write(entry_id + '\n')

        # Atualiza o arquivo com hashes processados
        with open(hashes_processados, 'w') as f:
            f.write('\n'.join(hashes_armazenados))

    except Exception as e:
        print(f"Erro ao processar o e-mail: {e}")

print("Processamento concluído!")

# segunda fase pro projeto, ler varios aqruivos em excel parar obter o resultado