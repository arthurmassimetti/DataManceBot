import pyautogui as pygui
from selenium import webdriver
import time
import undetected_chromedriver as uc
from undetected_chromedriver import *
from selenium.webdriver.support.ui import WebDriverWait
import time
import pyautogui
from openpyxl import load_workbook

from openpyxl import load_workbook
import pyautogui
import time


def buscaExcel():
    workbook = load_workbook(r'C:\Users\ADMINISTRATOR\Desktop\excelRobo\robo.xlsx')

    worksheet = workbook.active

    # Inicializar listas para armazenar os dados extraídos
    cpf_list = []
    nome_list = []
    tipo_exame_list = []
    data_exame_list = []

    # Iterar sobre as linhas e extrair os dados
    for row in worksheet.iter_rows(min_row=1, max_row=1, values_only=True):  # Comece da segunda linha, assumindo que a primeira linha são cabeçalhos
        cpf_list.append(row[0])
        nome_list.append(row[1])
        tipo_exame_list.append(row[2])
        data_exame_list.append(row[3])

    # Imprimir e pausar após cada linha
    for i in range(len(cpf_list)):
        print("\n" * 2)
        print("CPF:", cpf_list[i])
        print("Nome:", nome_list[i])
        print("Tipo de Exame:", tipo_exame_list[i])
        print("Data do Exame:", data_exame_list[i])
        print("\n" * 2)

        # Aguarde 5 segundos antes de imprimir a próxima linha
        time.sleep(0.3)

driver = uc.Chrome() #define o navegador (undetected chrome)

def fazer_login_um():

    print('iniciando automação!')
    driver.get("https://rh.dtm.com.br/")  # define o url para pesquisa
    driver.maximize_window()  # maximiza a tela

    usuario = 'FORT-46'  # acesso de usuario

    print('inserindo usuario')
    login = driver.find_element('xpath', '//*[@id="Editbox1"]')
    login.click()
    login.send_keys(usuario)
    time.sleep(0.3)

    senha = '3Q7kViZyqt'  # senha de usuario

    print('inserindo senha')
    login_senha = driver.find_element('xpath', '//*[@id="Editbox2"]')
    login_senha.click()
    login_senha.send_keys(senha)
    time.sleep(0.3)

    print('confirmando login')
    clicar_login = driver.find_element('xpath', '//*[@id="buttonLogOn"]')
    clicar_login.click()
    time.sleep(15)

def fazer_login_um_sem_page():

    usuario = 'FORT-46'  # acesso de usuario

    login = driver.find_element('xpath', '//*[@id="Editbox1"]')
    login.click()
    login.send_keys(usuario)
    time.sleep(0.3)

    senha = '3Q7kViZyqt'  # senha de usuario

    login_senha = driver.find_element('xpath', '//*[@id="Editbox2"]')
    login_senha.click()
    login_senha.send_keys(senha)
    time.sleep(0.3)

    clicar_login = driver.find_element('xpath', '//*[@id="buttonLogOn"]')
    clicar_login.click()
    time.sleep(15)

def fazer_login_dois():
    pyautogui.moveTo(734, 419, duration=2)
    pyautogui.click()
    time.sleep(0.3)

    print("apagando caso tenha algo escrito")
    pyautogui.typewrite(['backspace'] * 8)
    time.sleep(0.3)
    pyautogui.typewrite("PROSAUDE")
    time.sleep(0.3)  # pausa aqui

    pyautogui.moveTo(729, 440, duration=2)
    pyautogui.click()
    time.sleep(0.3)

    print("apagando caso tenha algo escrito")
    pyautogui.typewrite(['backspace'] * 10)
    time.sleep(0.3)
    pyautogui.typewrite("Pro@@2023")
    time.sleep(0.3)
    pyautogui.moveTo(779, 531, duration=2)

    pyautogui.click()
    time.sleep(1)

def exames_medicos():

    print('fechando anuncios')
    time.sleep(4)
    pyautogui.typewrite(['enter'] * 6)
    time.sleep(1)

    print('indo até MO')
    pyautogui.moveTo(222, 150, duration=2)
    pyautogui.click()
    time.sleep(0.3)

    print('clicando em exames medicos')
    pyautogui.moveTo(1006, 198, duration=2)
    pyautogui.click()
    time.sleep(0.3)

    print('especifica exames medicos')
    pyautogui.moveTo(311, 239, duration=2)
    pyautogui.click()
    time.sleep(0.3)

    print('abre exames medicos')
    pyautogui.moveTo(301, 283, duration=2)
    pyautogui.click()
    time.sleep(0.3)

def pesquisa_nome(nome_list, i):

    nome_list= []

    pyautogui.moveTo(339, 238, duration=2)
    pyautogui.click()
    time.sleep(0.3)

    print('abrindo aba para nome')
    pyautogui.moveTo(805, 363, duration=2)
    pyautogui.click()
    print('deixando vazio o espaço para inserção de nome')
    pyautogui.typewrite(['backspace'] * 8)
    time.sleep(0.3)

    pyautogui.typewrite(nome_list[i], interval=0.5)  # Aqui vai a procedure de nome
    pyautogui.typewrite(['enter'] * 2)
    time.sleep(0.3)

def pesquisa_cpf():

    cpf = cpf_list
    pyautogui.moveTo(639,240, duration=2)
    pyautogui.click()
    pyautogui.click()
    print('inserindo CPF de funcionario')
    pyautogui.typewrite("38426423850",  interval=0.2)
    pyautogui.typewrite(['enter'] * 1)
    time.sleep(0.3)

def sair():

    print('saindo do site')
    pyautogui.typewrite(['esc'] * 7)
    pyautogui.moveTo(1372, 113, duration=2)
    pyautogui.click()
    time.sleep(0.2)
    pyautogui.press('tab')
    time.sleep(0.2)
    pyautogui.press('enter')

x, y = 719, 349

def verificar_cor(x, y):

    cor = pyautogui.pixel(x, y)
    if cor == (132, 130, 255):
        print("login não realizado, realizando login para concluir procedimentos")
        fazer_login_dois()

    else:
        print("Login já realizado, prosseguindo com os procedimentos")
        pyautogui.typewrite(['esc'] * 3)

def verificarCorSecundaria():
    x, y = 133, 600

    print('verificando cor de fundo')
    for i in range(50):
        lugardacor = pyautogui.pixel(x, y)
        if lugardacor == (231,239,247):
            time.sleep(1.5)
            break
        else:
            continue

def incluir():

    print('abrindo incluir')
    pyautogui.moveTo(305, 711, duration=2)
    pyautogui.click()
    time.sleep(2)

def nomeMedico():

    print('indo até nome do medico')
    pyautogui.moveTo(685, 347, duration=2)
    pyautogui.click()
    time.sleep(1)

    print('localizando barra de pesquisa')
    pyautogui.moveTo(403, 413, duration=2)
    pyautogui.click()
    time.sleep(1)

    print('inserindo nome do medico (FRANCISCO CANHETTI')
    pyautogui.typewrite("FRANCISCO CANHETTI")

    pyautogui.moveTo(511, 564, duration=2)
    pyautogui.click()
    time.sleep(1)

    print('confirmando e fechando')
    pyautogui.moveTo(716, 337, duration=2)
    pyautogui.click()
    pyautogui.click()
    time.sleep(2)

def pesquisaLaboratorio():

    print('indo para laboratorio')
    pyautogui.moveTo(482, 369, duration=2)
    pyautogui.click()
    time.sleep(1)

    pyautogui.moveTo(412, 442, duration=2)
    pyautogui.click()
    time.sleep(1)

    print('pesquisando nome da empresa (PRO OCUPACIONAL)')
    pyautogui.typewrite("PRO OCUPACIONAL")

    pyautogui.moveTo(384, 568, duration=2)
    pyautogui.click()
    time.sleep(1)

    print('confirma e fecha')
    pyautogui.moveTo(791, 336, duration=2)
    pyautogui.click()
    pyautogui.click()
    time.sleep(2)

def dataExameigual():

    print("indo para data")
    pyautogui.moveTo(453, 217, duration=2)
    pyautogui.click()
    pyautogui.click()
    time.sleep(0.3)

    print('copiando data')
    pyautogui.typewrite(str(data_exame_list[i]), interval=0.5)

    time.sleep(0.3)

    pyautogui.moveTo(658, 217, duration=2)
    pyautogui.click()
    pyautogui.click()
    time.sleep(0.3)

    print("colando data")
    pyautogui.typewrite(str(data_exame_list[i]), interval=0.5)

    time.sleep(0.3)

    pyautogui.moveTo(1059, 396, duration=2)
    pyautogui.click()
    pyautogui.click()
    time.sleep(0.3)

    data = str(data_exame_list[i])  # Converta a data para string
    ultimo_digito = int(data[-1])  # Obtenha o último dígito e converta para inteiro
    novo_ultimo_digito = ultimo_digito + 1  # Adicione 1 ao último dígito
    nova_data = data[:-1] + str(
    novo_ultimo_digito)  # Concatene a parte da data sem o último dígito com o novo último dígito
    print("Nova Data do Exame:", nova_data)

    pyautogui.typewrite(nova_data, interval=0.5)

    time.sleep(0.3)

def ordemExame():

    print('indo até ordem de exame')
    pyautogui.moveTo(503, 397, duration=2)
    pyautogui.click()
    time.sleep(0.3)

    print('selecionando sequencial')
    pyautogui.moveTo(460, 439, duration=2)
    pyautogui.click()
    time.sleep(0.3)

def statusExame():

    print('indo até status')
    pyautogui.moveTo(934, 218, duration=2)
    pyautogui.click()
    time.sleep(0.3)

    print('selecionando finalizado')
    pyautogui.moveTo(931, 246, duration=2)
    pyautogui.click()
    time.sleep(0.3)

def medicoExame():

    print('indo até medico ')
    pyautogui.moveTo(410,324, duration=2)
    pyautogui.click()
    time.sleep(0.3)

    print('campo de pesquisa')
    pyautogui.moveTo(415,444, duration=2)
    pyautogui.click()
    time.sleep(0.3)

    print('inserindo nome do medico responsavel (FRANCISCO CANHETTI')
    pyautogui.typewrite("FRANCISCO CANHETTI")
    pyautogui.moveTo(511, 564, duration=2)
    pyautogui.click()
    time.sleep(1)

    print('confirmando e fechando')
    pyautogui.moveTo(716, 337, duration=2)
    pyautogui.click()
    pyautogui.click()
    time.sleep(2)


buscaExcel()
time.sleep(1)

fazer_login_um()
time.sleep(1)

verificarCorSecundaria()
time.sleep(1)

verificar_cor(x, y)
time.sleep(1)

verificarCorSecundaria()
time.sleep(1)

exames_medicos()
time.sleep(1)

pesquisa_nome(nome_list, i)
time.sleep(1)

incluir()
time.sleep(1)

dataExameigual()
time.sleep(1)

ordemExame()
time.sleep(1)

statusExame()
time.sleep(1)

nomeMedico()
time.sleep(0.5)

medicoExame()
time.sleep(0.5)

pesquisaLaboratorio()
time.sleep(999999)








#369,796 gravar
