# Imports
import os

from app.loggers import AppLogger
from webdriver_manager.chrome import ChromeDriverManager

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

class WebDriver():
    def __init__(self, webdriver_path: str = None):
        """
            Construtor. O self._driver procura o webdriver mais atualizado para acompanhar as atualizações do Google Chrome. Caso haja algum erro, busca-se um webdriver já baixado.
        """
        options = webdriver.ChromeOptions()
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        try:
            self._driver = webdriver.Chrome(executable_path=ChromeDriverManager().install(), options=options)
            self.logger.info(f'Download do WebDriver realizado com sucesso! Versão: {sorted(os.listdir(webdriver_path))[-1]}')
        except:
            webdriver_version = sorted(os.listdir(webdriver_path))[-1]
            self._driver = webdriver.Chrome(executable_path=os.path.join(webdriver_path, webdriver_version, 'chromedriver.exe'), options=options)
    
    def get_driver(self):
        """
            Retorna o driver do navegador.
        """        
        return self._driver

    def open(self, url:str):
        """
            Abre o navegador na url desejada.
            :param url: URL que se deseja acessar.
        """
        self._driver.get(url)


    def login(self, user:str, pwd:str):
        """
            Realiza o login a partir das credenciais.
            IMPORTANTE: Necessário alterar ID do campo caso necessidade.
             
            :param user: usuário para login no AGForms.
            :param pwd: senha para login no AGForms.
        """
        login = WebDriverWait(self._driver, 60).until(EC.presence_of_element_located((By.ID, 'Login')), message='Não foi possivel entrar com o username.')
        login.send_keys(user)
        senha = WebDriverWait(self._driver, 60).until(EC.presence_of_element_located((By.ID, 'Senha')), message='Não foi possivel entrar com a senha.')
        senha.send_keys(pwd + Keys.RETURN)

    def insert_info(self, xpath: str, info: str, enter: bool = False):
        """
            Entra com as informações em um campo de texto a partir do xpath do campo.
            :param xpath: xpath do campo.
            :param info: informação que se deseja entrar no campo.
            :param enter: flag que sinaliza se é necessário pressionar Enter depois de se entrar com o dado.
        """
        element = WebDriverWait(self._driver, 30).until(EC.element_to_be_clickable((By.XPATH, xpath)), message=f'Erro ao preencher o campo com a informação "{info}".')
        element.send_keys(info + Keys.ENTER) if enter else element.send_keys(info)