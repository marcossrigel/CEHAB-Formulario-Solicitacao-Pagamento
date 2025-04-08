from oauth2client.service_account import ServiceAccountCredentials
from cryptography.fernet import Fernet
import pyautogui
import webbrowser
import time
import os
import gspread

def conectar_google_sheets(caminho_credencial_temp):
    scopes = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive',
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(caminho_credencial_temp, scopes)
    return gspread.authorize(creds)

def enviar_mensagens(planilha):
    dados = planilha.get_all_records()
    cabecalho = planilha.row_values(1)
    coluna_status_index = cabecalho.index('Status') + 1

    for i, linha in enumerate(dados, start=2):
        if linha.get('Status') == 'Liberado':
            if linha.get('Origem da demanda / Setor') == 'DOE - Diretoria de Obras Estratégicas':
                nome = 'Conceição'
                telefone = '+5581991492389'
            if linha.get('Origem da demanda / Setor') == 'DOB - Diretoria de Obras':
                nome = 'Ana Paula'
                telefone = '+5581998772704'
            if linha.get('Origem da demanda / Setor') == 'DPH - Diretoria de Programas Habitacionais':
                if (linha.get('Empresa') == 'Maia Melo Engenharia') & (linha.get('Nº Contrato') == '114/2022'):
                    nome = 'Mariana'
                    telefone = '+5581995026464'
                else:
                     nome = 'Bianca'
                     telefone = '+5581995384148'

            mensagem = (
                        f'Olá, {nome}! Você já pode solicitar a disponibilidade financeira para pagamento do(s) contrato(s) abaixo: %0A %0A'
                        f'Objeto do Contrato: {linha.get("Objeto do contrato")}%0A%0A'
                        f'Local da obra ou serviço: {linha.get("Local da obra ou serviço")}%0A'
                        f'Empresa: {linha.get("Empresa")}%0A'
                        f'Número do Contrato: {linha.get("Nº Contrato")}%0A'
                        f'BM nº: {linha.get("BM nº ")}%0A'
                        f'Valor: {linha.get("Valor")}%0A'
                        f'Fonte de Recursos do Pagamento: {linha.get("Fonte de Recursos do Pagamento")}%0A'
                        f'Número do SEI: {linha.get("Nº SEI")}'
                    )

            webbrowser.open(f'https://web.whatsapp.com/send?phone={telefone}&text={mensagem}')

            time.sleep(10)
            pyautogui.press('enter')
            pyautogui.hotkey('ctrl', 'w')
            planilha.update_cell(i, coluna_status_index, 'Enviado')
            time.sleep(5)

          
def descriptografar_credencial(caminho_cripto, caminho_temp):
    fernet = Fernet('Cv_tn1OWIWHUVQ3YW6jy2SUAdZWE-Br2ws05_jfs1o0=')
    with open(caminho_cripto, 'rb') as arquivo_criptografado:
        criptografado = arquivo_criptografado.read()
    descriptografado = fernet.decrypt(criptografado)
    with open(caminho_temp, 'wb') as arquivo_temp:
        arquivo_temp.write(descriptografado)

    return caminho_temp


def main():
    caminho_cripto = 'formulariosolicitacaopagamento-f683a63c3e41.json'
    caminho_temp = 'credenciais_temp.json'

    try:
        credencial_temp = descriptografar_credencial(caminho_cripto, caminho_temp)
        client = conectar_google_sheets(credencial_temp)

        planilha_completa = client.open(
            title="FORMULÁRIO DE SOLICITAÇÃO DE PAGAMENTO (respostas)", 
            folder_id="1lZL70UCxfGk7uscdzWyv6JQSsgSVMX0b"
        )
        planilha = planilha_completa.get_worksheet(0)
        enviar_mensagens(planilha)

    finally:
        if os.path.exists(caminho_temp):
            os.remove(caminho_temp)


if __name__ == '__main__':
    main()