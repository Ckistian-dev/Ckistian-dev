from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
import time
from datetime import datetime
import os
import glob
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Caminho para o ChromeDriver
servico = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=servico)

# Função para aguardar e interagir com elementos
def esperar_e_clicar(seletor, tipo=By.XPATH, tempo=30):
    elemento = WebDriverWait(driver, tempo).until(EC.element_to_be_clickable((tipo, seletor)))
    try:
        elemento.click()
    except StaleElementReferenceException:
        elemento = WebDriverWait(driver, tempo).until(EC.element_to_be_clickable((tipo, seletor)))
        elemento.click()

# Função para aguardar e enviar texto para elementos
def esperar_e_enviar_texto(seletor, texto, tipo=By.XPATH, tempo=30):
    elemento = WebDriverWait(driver, tempo).until(EC.visibility_of_element_located((tipo, seletor)))
    elemento.clear()
    elemento.send_keys(texto)
    return elemento

# Funções para autenticação e upload no Google Drive
def authenticate_gdrive():
    SCOPES = ['https://www.googleapis.com/auth/drive.file']
    creds = service_account.Credentials.from_service_account_file('credentials.json', scopes=SCOPES)
    service = build('drive', 'v3', credentials=creds)
    return service

def get_latest_downloads():
    downloads_folder = os.path.join(os.path.expanduser('~'), 'Downloads')
    xls_files = glob.glob(os.path.join(downloads_folder, '*.xls'))
    csv_files = glob.glob(os.path.join(downloads_folder, '*.csv'))

    xls_files.sort(key=os.path.getmtime, reverse=True)
    csv_files.sort(key=os.path.getmtime, reverse=True)

    latest_files = []
    if xls_files:
        latest_files.append(xls_files[0])  # Adiciona o último arquivo .xls
    if csv_files:
        latest_files.append(csv_files[0])  # Adiciona o último arquivo .csv

    return latest_files

def upload_files(service, files, folder_id):
    for file in files:
        file_metadata = {
            'name': os.path.basename(file),
            'parents': [folder_id]
        }
        media = MediaFileUpload(file, mimetype='application/octet-stream')
        try:
            service.files().create(body=file_metadata, media_body=media, fields='id').execute()
            print(f'Uploaded: {file}')
        except Exception as e:
            print(f'An error occurred: {e}')

# Definir data inicial
data_atual = datetime.now()

mes_subtraido = (data_atual.month - 4) % 12
if mes_subtraido == 0:
    mes_subtraido = 12  # Corrige para Dezembro, caso o valor seja 0
    ano_correspondente = data_atual.year - 1
else:
    ano_correspondente = data_atual.year if (data_atual.month - 4) > 0 else data_atual.year - 1

data_inicial = f"01/{mes_subtraido:02d}/{ano_correspondente}"  

try:

# Captar o relatório do Tiny ERP
    driver.get('https://erp.tiny.com.br/')

# Login Tiny
    esperar_e_enviar_texto('//*[@id="kc-content-wrapper"]/react-login/div/div/div[1]/div[1]/div[1]/form/div[1]/div/input', 'adm01@talatto.com.br')
    esperar_e_enviar_texto('//*[@id="kc-content-wrapper"]/react-login/div/div/div[1]/div[1]/div[1]/form/div[2]/div/input', 'Talattomudar123@')
    esperar_e_clicar('//*[@id="kc-content-wrapper"]/react-login/div/div/div[1]/div[1]/div[1]/form/div[3]/button')
    esperar_e_clicar('//*[@id="bs-modal-ui-popup"]/div/div/div/div[3]/button')

# Aguardar o login processar
    time.sleep(10)

# Acessar relatórios
    driver.get('https://erp.tiny.com.br/relatorios_personalizados#/view/376')

# Ajustar Período
    esperar_e_clicar('//*[@id="root-relatorios-personalizados"]/div/div[1]/div[4]/ul/li[1]/a')
    esperar_e_clicar('//*[@id="root-relatorios-personalizados"]/div/div[1]/div[4]/ul/li[1]/div/div[2]/div/div[5]/button')
    esperar_e_enviar_texto('//*[@class="form-control hasDatepicker"]', data_inicial)
    esperar_e_clicar('//*[@id="root-relatorios-personalizados"]/div/div[1]/div[4]/ul/li[1]/div/div[4]/button[1]')

# Fazer o Download do Relatório
    esperar_e_clicar('//button[text()="download"]')
    esperar_e_clicar('//button[text()=" processar outro arquivo"]')
    esperar_e_clicar('//button[text()=" continuar"]')
    esperar_e_clicar('//button[text()=" Baixar"]', tempo=999)

# Navegar para o Lira
    driver.get('https://lirasistemas.com.br/nfe/')
    
# Login
    esperar_e_enviar_texto('//*[@id="user"]', 'cristian_talatto')
    esperar_e_enviar_texto('//*[@id="senha"]', 'mudar123')
    esperar_e_clicar('//*[@id="login"]/form/fieldset/div[1]/span[2]/input')

# Tempo para fazer o login
    time.sleep(10)

# Acessar o relatório
    driver.get('https://lirasistemas.com.br/nfe/relatorios/pedidos.jsp?menuRelatorio=active%20open')

# Alterar o período
    esperar_e_enviar_texto("//*[@id='data_inicio']", data_inicial)
    data_final = datetime.now().strftime("%d%m%Y")
    esperar_e_enviar_texto("//*[@id='data_final']", data_final)

# Baixar o Relatório
    time.sleep(2)
    esperar_e_clicar('//*[@id="btnGearCSV"]')

# Fazer o upload dos dois arquivos
    folder_id = '1rycL9CwMmWKTuWjgLHFm3vopsgYwwufw'  # Substitua pelo ID da sua pasta no Google Drive
    service = authenticate_gdrive()
    latest_files = get_latest_downloads()
    upload_files(service, latest_files, folder_id)

# Tempo para Baixar o Relatório
    time.sleep(5)

# Acessar a planilha para atualizar os dados
    driver.get('https://docs.google.com/spreadsheets/d/1vITmZvs-JskK3xVyIaO4HtfS5rZZ-CsGL055U8Qq3sM/edit?usp=sharing')

# Manter a planilha aberta
    time.sleep(999999999)

finally:
    driver.quit()
