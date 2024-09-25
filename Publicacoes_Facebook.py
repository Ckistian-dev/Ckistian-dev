from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pickle
import os
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import sys  # Adicionado para verificar se o código está rodando como .exe

# Função para configurar o driver do Chrome
def configurar_driver():
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Função para aguardar e enviar texto para elementos
def esperar_e_enviar_texto(driver, seletor, texto, tipo=By.XPATH, tempo=1):
    elemento = WebDriverWait(driver, tempo).until(EC.visibility_of_element_located((tipo, seletor)))
    elemento.clear()
    elemento.send_keys(texto)

# Função para aguardar e capturar texto
def esperar_e_capturar_texto(driver, seletor, tipo=By.XPATH, tempo=1):
    try:
        elemento = WebDriverWait(driver, tempo).until(EC.visibility_of_element_located((tipo, seletor)))
        return elemento.text
    except Exception:
        return None

# Função para salvar e carregar cookies
def gerenciar_cookies(driver, caminho_arquivo):
    if os.path.exists(caminho_arquivo):
        with open(caminho_arquivo, 'rb') as f:
            for cookie in pickle.load(f):
                driver.add_cookie(cookie)
        return True
    return False

# Função para capturar quantidade de visualizações
def capturar_visualizacoes(driver):
    xpaths = [
        '/html/body/div[1]/div/div[1]/div/div[3]/div/div/div[1]/div[1]/div[2]/div/div/div[2]/div/div[3]/div/div/div/div/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div[1]/div/div[1]/div/div/div/div/div[1]/div/div[2]/span',
        '/html/body/div[1]/div/div/div[1]/div/div[3]/div/div/div[1]/div[1]/div[2]/div/div/div[2]/div/div[3]/div/div/div/div/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div[1]/div/div[1]/div/div/div/div/div[1]/div/div[2]/span',
        '/html/body/div[1]/div/div[1]/div/div[3]/div/div/div[1]/div[1]/div[2]/div/div/div[2]/div/div[4]/div/div/div/div/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div[1]/div/div[1]/div/div/div/div/div[1]/div/div[2]/span'
    ]
    
    for xpath in xpaths:
        visualizacao = esperar_e_capturar_texto(driver, xpath)
        if visualizacao:
            return visualizacao
    return None

# Função para autenticar e acessar a planilha do Google Sheets
def acessar_planilha():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
    
    # Resolver o caminho do 'sheetcredentials.json'
    if getattr(sys, 'frozen', False):  # Verifica se o código está rodando como .exe
        creds_path = os.path.join(sys._MEIPASS, 'sheetcredentials.json')
    else:
        creds_path = os.path.join(os.path.dirname(__file__), 'sheetcredentials.json')
    
    creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
    return gspread.authorize(creds).open("Publicações Facebook").sheet1

# Função para adicionar dados na planilha
def adicionar_dados(planilha, email, visualizacao):
    planilha.append_row([datetime.now().strftime("%d/%m/%Y"), email, visualizacao])

# Função para fazer login e capturar visualizações
def login_e_capturar_visualizacoes(email):
    driver = configurar_driver()
    caminho_arquivo_cookies = f"cookies_{email}.pkl"

    try:
        driver.get('https://www.facebook.com/login')
        if not gerenciar_cookies(driver, caminho_arquivo_cookies):
            driver.execute_script(f"alert('Digite a senha para a conta: {email}');")
            time.sleep(3)
            driver.switch_to.alert.accept()
            esperar_e_enviar_texto(driver, '//*[@id="email"]', email)

            # Aguardar o login
            WebDriverWait(driver, 9999999).until(lambda d: d.current_url == "https://www.facebook.com/")
            # Salvar os cookies após o login
            with open(caminho_arquivo_cookies, 'wb') as f:
                pickle.dump(driver.get_cookies(), f)
        else:
            driver.refresh()

        driver.get('https://www.facebook.com/marketplace/you/dashboard')
        visualizacao = capturar_visualizacoes(driver)
        if visualizacao:
            print(f"Valor capturado para {email}: {visualizacao}")
            planilha = acessar_planilha()
            adicionar_dados(planilha, email, visualizacao)
        else:
            print(f"Visualização não encontrada para {email}.")

    finally:
        driver.quit()

# Lê o arquivo .txt e processa cada par de e-mail
with open('Contas.txt', 'r') as arquivo:
    for linha in arquivo:
        login_e_capturar_visualizacoes(linha.strip())

# Abre a planilha dos dados
driver = configurar_driver()
driver.get('https://docs.google.com/spreadsheets/d/1EPddtH-eGVEkowYD6KvU-7ZwpEbqInaaXMVkcPgKtq8/edit?usp=sharing')
time.sleep(999999)
