import os
import sys
import time
import json
import smtplib
from datetime import date
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

#Caminho do arquivo
diretorio_atual, __DUMMY = os.path.split(os.path.abspath(__file__))
raiz = os.path.abspath(os.path.join(diretorio_atual, '..'))
pasta_downloads = os.path.join(raiz, 'Users', 'barba', 'Downloads')
caminho_pdf = os.path.join(pasta_downloads, 'Uniesp.pdf')

data_hoje = date.today()

#As informações foram trocadas por segurança
matricula = "aaaaaaaaaa"
senha = "bbbbbbbbbb"
email_remetente = 'meuemail@gmail.com' 
senha_remetente = 'cccccccccc'
email_destinatario = 'outroemail@gmail.com'

def baixa_boleto_sistema_uniesp(matricula, senha):

    #Configurações para poder salvar o PDF corretamente
    chrome_options = webdriver.ChromeOptions()
    settings = {
        "recentDestinations": [{
            "id": "Save as PDF",
            "origin": "local",
            "account": "",
        }],
        "selectedDestinationId": "Save as PDF",
        "version": 2
    }
    prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings)}
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument('--kiosk-printing')
    caminho_executavel_chromedriver = r'C:\Users\barba\Downloads\chromedriver.exe'

    driver = webdriver.Chrome(executable_path=caminho_executavel_chromedriver, options=chrome_options)
    driver.get('https://sistemas.iesp.edu.br/iesponline/Login')
    driver.maximize_window()

    caixa_de_texto_matricula = driver.find_element(by=By.XPATH, value='//*[@id="login"]')
    caixa_de_texto_matricula.send_keys(matricula)
    caixa_de_texto_senha = driver.find_element(by=By.XPATH, value='//*[@id="senha"]')
    caixa_de_texto_senha.send_keys(senha)
    time.sleep(1)

    enter = driver.find_element(by=By.XPATH, value='//*[@id="login-box-body"]/form[1]/div[3]/div[2]/button')
    enter.click()
    time.sleep(1)
    botao_financeiro = driver.find_element(by=By.XPATH, value='//*[@id="wrapper"]/aside/section/ul/li[2]/a')
    botao_financeiro.click()
    time.sleep(1)
    ficha_financeira = driver.find_element(by=By.XPATH, value='//*[@id="wrapper"]/aside/section/ul/li[2]/ul/li[1]/a/span')
    ficha_financeira.click()
    time.sleep(1)
    botao_de_impressao = driver.find_element(by=By.XPATH, value='//*[@id="collapse0"]/div/div/div/table/tbody/tr[{}]/td[6]/a'.format(data_hoje.month))
    botao_de_impressao.click()
    time.sleep(3)

    driver.quit()

def envia_email(email_remetente, senha_remetente, email_destinatario):

    mensagem = MIMEMultipart()
    mensagem['From'] = email_remetente
    mensagem['To'] = email_destinatario
    mensagem['Subject'] = 'O boleto novo chegou!'
    mensagem.attach(MIMEText('Conteúdo do e-mail', 'plain'))

    with open(caminho_pdf, 'rb') as arquivo_pdf:
        anexo = MIMEApplication(arquivo_pdf.read(), _subtype='pdf')
    anexo.add_header('Content-Disposition', 'attachment',filename='BoletoUniesp{}.pdf'.format(data_hoje.month))

    mensagem.attach(anexo)

    sessao = smtplib.SMTP('smtp.gmail.com', 587)
    sessao.starttls()
    sessao.login(email_remetente, senha_remetente)
    sessao.send_message(mensagem)
    sessao.quit()

if __name__ == '__main__':

    baixa_boleto_sistema_uniesp(matricula, senha)
    envia_email(email_remetente, senha_remetente, email_destinatario)

#Excluindo o arquivo para não acumular
if os.path.isfile(caminho_pdf):
    os.remove(caminho_pdf)

