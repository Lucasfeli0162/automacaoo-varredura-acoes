from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
from time import sleep
import os
import pyautogui
from pyperclip import copy
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

pyautogui.PAUSE = 1.8


class Main:
    def __init__(self):
        resposta = ''
        self.empresas = []
        while True:
            self.empresas.append(input('Digite o nome da empresa que quer ver a ação: '))
            resposta = input('Gostaria de adicionar mais empresas a lista? ')
            if resposta not in 'Ss':
                break
        self.email_from = self.verify_email(input('Digite o seu e-mail(lembre-se que só aceitamos outlook): '))
        self.password = input('Digite a senha do seu e-mail: ')
        self.email_to = self.verify_email(input('Digite o e-mail que você gostaria de enviar a planilha: '))
        self.subject = 'Planilha de ações'
        self.msg = f'''Prezado, Lucas.
                    Em anexo está a planilha com as ações que você pediu
                    Ao todo você pediu {len(self.empresas)} ações
                    Att. Lucas Feitosa'''
        self.caminho = os.path.abspath(os.getcwd())
        self.codigos = []
        self.precos = []
        self.navegador = webdriver.Chrome()
        self.navegador.maximize_window()
        self.pegar_codigos_e_precos()
        self.criar_tabela()
        self.enviar_email()

    def verify_email(self, email):
        import re
        while True:
            if not re.match(r"[^@]+@[^@]+\.[^@]+", email):
                email = input('ERRO! e-mail inválido, tente novamente: ')
            else:
                break
        return email

    def pegar_codigos_e_precos(self):
        for empresa in self.empresas:
            self.navegador.get('https://www.google.com')
            pesquisa = self.navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input')
            pesquisa.send_keys(f'Código ação {empresa}')
            pesquisa.submit()
            try:
                codigo = self.navegador.find_element(By.CLASS_NAME, 'WuDkNe').get_attribute('innerHTML')
                valor = self.navegador.find_element(By.XPATH, '//*[@id="knowledge-finance-wholepage__entity-summary"]/div/g-card-section/div/g-card-section/div[2]/div[1]/span[1]/span/span[1]').get_attribute('innerHTML')
            except:
                self.navegador.get('https://www.google.com')
                pesquisa = self.navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input')
                pesquisa.send_keys(f'investidor10 {empresa}')
                pesquisa.submit()
                self.navegador.find_element(By.XPATH, '//*[@id="rso"]/div[1]/div/div[1]/div/div/div/div[1]/a').click()
                codigo = self.navegador.find_element(By.XPATH, '//*[@id="header_action"]/div[1]/div[2]/h1').get_attribute('innerHTML')
                valor = self.navegador.find_element(By.XPATH, '//*[@id="cards-ticker"]/div[1]/div[2]/div/span').get_attribute('innerHTML')
            self.codigos.append(codigo)
            self.precos.append(valor)

    def criar_tabela(self):
        df = pd.DataFrame()
        df['empresa'] = self.empresas
        df['codigo'] = self.codigos
        df['valores'] = self.precos
        df.to_excel(excel_writer='valores-acoes.xlsx', index=False)

    def enviar_email(self):
        self.navegador.get('https://login.live.com/login.srf?wa=wsignin1.0&rpsnv=13&ct=1634597740&rver=7.0.6737.0&wp=MBI_SSL&wreply=https%3a%2f%2foutlook.live.com%2fowa%2f%3fnlp%3d1%26RpsCsrfState%3d1ac800e8-2569-5c37-e0f4-b752160b7ff9&id=292841&aadredir=1&whr=outlook.com&CBCXT=out&lw=1&fl=dob%2cflname%2cwld&cobrandid=90015')
        sleep(3)
        email = self.navegador.find_element(By.CSS_SELECTOR, '#i0116')
        email.send_keys(self.email_from)
        self.navegador.find_element(By.CSS_SELECTOR, '#idSIButton9').click()
        sleep(3)
        senha = self.navegador.find_element(By.CSS_SELECTOR, '#i0118')
        senha.send_keys(self.password)
        senha.submit()
        sleep(5)
        botao = self.navegador.find_element(By.CSS_SELECTOR, '#idSIButton9')
        botao.click()
        sleep(5)
        tela = self.navegador.find_element(By.TAG_NAME, 'html')
        tela.send_keys('n')
        sleep(3)
        action = ActionChains(self.navegador)
        copy(self.email_to)
        action.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
        action.send_keys(Keys.TAB)
        copy(self.subject)
        action.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
        action.send_keys(Keys.TAB)
        copy(self.msg)
        action.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
        self.caminho = str(self.caminho)
        self.caminho.replace('\\', '/')
        self.caminho = self.caminho + '/valores-acoes.xlsx'
        print(self.caminho)
        file = self.navegador.find_element(By.CLASS_NAME, '_3h0ZVoFuunvPZlu_hLoIBu')
        file.send_keys(self.caminho)
        sleep(5)
        action.key_down(Keys.CONTROL).send_keys(Keys.ENTER).key_up(Keys.CONTROL).perform()


main = Main()
