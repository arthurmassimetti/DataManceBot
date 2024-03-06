import time

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, NoAlertPresentException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

opcoes = webdriver.ChromeOptions()
opcoes.add_argument("--headless=new")
driver = webdriver.Chrome()

def fazerLoginSoc():
    senha = ('Pro@@2023')


    driver.get("https://sistema.soc.com.br/WebSoc/")
    driver.maximize_window()  # maximizando a janela
    login = driver.find_element(By.XPATH, '//*[@id="usu"]').send_keys("proarthurms")  # Colocando o usuario no campo user
    senha = driver.find_element(By.XPATH, '//*[@id="senha"]').send_keys(senha)  # Colocando a senha no campo pass
    time.sleep(0.5)  # Fazendo a declaração de todos os botões do ID
    botao1 = driver.find_element(By.XPATH, '//*[@id="bt_1"]')
    botao2 = driver.find_element(By.XPATH, '//*[@id="bt_2"]')
    botao3 = driver.find_element(By.XPATH, '//*[@id="bt_3"]')
    botao4 = driver.find_element(By.XPATH, '//*[@id="bt_4"]')
    botao5 = driver.find_element(By.XPATH, '//*[@id="bt_5"]')
    botao6 = driver.find_element(By.XPATH, '//*[@id="bt_6"]')
    botao7 = driver.find_element(By.XPATH, '//*[@id="bt_7"]')
    botao8 = driver.find_element(By.XPATH, '//*[@id="bt_8"]')
    botao9 = driver.find_element(By.XPATH, '//*[@id="bt_9"]')
    botao0 = driver.find_element(By.XPATH, '//*[@id="bt_0"]')

    if botao1.get_attribute("value") == "3":
        botao1.click()
    else:
        pass

    if botao2.get_attribute("value") == "3":
        botao2.click()
    else:
        pass

    if botao3.get_attribute("value") == "3":
        botao3.click()
    else:
        pass

    if botao4.get_attribute("value") == "3":
        botao4.click()
    else:
        pass

    if botao5.get_attribute("value") == "3":
        botao5.click()
    else:
        pass

    if botao6.get_attribute("value") == "3":
        botao6.click()
    else:
        pass

    if botao7.get_attribute("value") == "3":
        botao7.click()
    else:
        pass

    if botao8.get_attribute("value") == "3":
        botao8.click()
    else:
        pass

    if botao9.get_attribute("value") == "3":
        botao9.click()
    else:
        pass

    if botao0.get_attribute("value") == "3":
        botao0.click()
    else:
        pass

    # PROXIMO DIGITO

    if botao1.get_attribute("value") == "7":
        botao1.click()
    else:
        pass

    if botao2.get_attribute("value") == "7":
        botao2.click()
    else:
        pass

    if botao3.get_attribute("value") == "7":
        botao3.click()
    else:
        pass

    if botao4.get_attribute("value") == "7":
        botao4.click()
    else:
        pass

    if botao5.get_attribute("value") == "7":
        botao5.click()
    else:
        pass
    if botao6.get_attribute("value") == "7":
        botao6.click()
    else:
        pass

    if botao7.get_attribute("value") == "7":
        botao7.click()
    else:
        pass

    if botao8.get_attribute("value") == "7":
        botao8.click()
    else:
        pass

    if botao9.get_attribute("value") == "7":
        botao9.click()
    else:
        pass

    if botao0.get_attribute("value") == "7":
        botao0.click()
    else:
        pass

    # PROXIMO DIGITO

    if botao1.get_attribute("value") == "3":
        botao1.click()
    else:
        pass

    if botao2.get_attribute("value") == "3":
        botao2.click()
    else:
        pass

    if botao3.get_attribute("value") == "3":
        botao3.click()
    else:
        pass

    if botao4.get_attribute("value") == "3":
        botao4.click()
    else:
        pass

    if botao5.get_attribute("value") == "3":
        botao5.click()
    else:
        pass

    if botao6.get_attribute("value") == "3":
        botao6.click()
    else:
        pass

    if botao7.get_attribute("value") == "3":
        botao7.click()
    else:
        pass

    if botao8.get_attribute("value") == "3":
        botao8.click()
    else:
        pass

    if botao9.get_attribute("value") == "3":
        botao9.click()
    else:
        pass

    if botao0.get_attribute("value") == "3":
        botao0.click()
    else:
        pass

    # PROXIMO DIGITO

    if botao1.get_attribute("value") == "8":
        botao1.click()
    else:
        pass

    if botao2.get_attribute("value") == "8":
        botao2.click()
    else:
        pass

    if botao3.get_attribute("value") == "8":
        botao3.click()
    else:
        pass

    if botao4.get_attribute("value") == "8":
        botao4.click()
    else:
        pass

    if botao5.get_attribute("value") == "8":
        botao5.click()
    else:
        pass

    if botao6.get_attribute("value") == "8":
        botao6.click()
    else:
        pass

    if botao7.get_attribute("value") == "8":
        botao7.click()
    else:
        pass

    if botao8.get_attribute("value") == "8":
        botao8.click()
    else:
        pass

    if botao9.get_attribute("value") == "8":
        botao9.click()
    else:
        pass

    if botao0.get_attribute("value") == "8":
        botao0.click()
    else:
        pass

    entrar = driver.find_element(By.XPATH, '//*[@id="bt_entrar"]').click()  # CLICANDO EM ENTRAR


def InteragirCaixaDeBusca():

    iframe = driver.find_element(By.XPATH,'//*[@id="socframe"]')
    driver.switch_to.frame(iframe)
    print('inserindo codigo da empresa FortServ')
    caixadebusca = driver.find_element(By.XPATH, '/html/body/form[1]/div[4]/div[2]/p/input')
    caixadebusca.send_keys('1515878')
    print('pesquisando na lupa')
    lupa = driver.find_element(By.XPATH, '/html/body/form[1]/div[4]/div[2]/p/a[1]/img')
    lupa.click()
    print('esperando carregar')
    wait = WebDriverWait(driver, 30)
    time.sleep(2.5)
    print('clicando na empresa')
    clicarempresa = driver.find_element(By.XPATH,'/html/body/form[1]/div[4]/div[2]/div[2]/table/tbody/tr/td[2]/a')
    clicarempresa.click()
    time.sleep(0.3)
    driver.switch_to.default_content()

def CaixaDeBuscaAgil():

    print('procurando a barra de pesquisa')
    busca = driver.find_element('xpath', '//*[@id="cod_programa"]')
    busca.click()
    print('inserindo codigo de pesquisa de exames (309)')
    busca.send_keys('309')
    busca.send_keys(Keys.ENTER)

#DENTRO DO 309 PARA BUSCA DE EXAMES

def SelecionarTipoExames():

    print('buscando ')
    todosExames = driver.find_element(By.XPATH, '/html/body/div[4]/div/main/form/section/section[2]/fieldset/div[3]/div/div/ul/li/div/label/span')
    todosExames.click()
    time.sleep(0.3)
    todosExames.click()



fazerLoginSoc()
time.sleep(1)
InteragirCaixaDeBusca()
time.sleep(1)
CaixaDeBuscaAgil()
time.sleep(555555)
SelecionarTipoExames()
time.sleep(44444)