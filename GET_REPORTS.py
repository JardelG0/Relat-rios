from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.print_page_options import PrintOptions
from selenium.webdriver.chrome.options import Options
from base64 import b64decode
import pandas as pd
import time
import os


def login():
    driver.find_element(By.NAME, "user").send_keys(log)
    driver.find_element(By.NAME, "password").send_keys(passw)
    driver.find_element(By.NAME, "password").send_keys(Keys.RETURN)
    time.sleep(0.5)
    driver.find_element(By.XPATH, "/html/body/div[3]/div/div[1]/div/div[4]/div/ul/li[2]/a").send_keys(Keys.RETURN)


def dataHora(data_hora_ini, data_hora_fin):
    time.sleep(1)
    if plat == "GS": 
        driver.find_element(By.XPATH, "/html/body/div[3]/div/div[13]/form/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div/div/input").send_keys(Keys.RETURN)
        i = 0
        while i < 14:
            i+=1
            driver.find_element(By.XPATH, "/html/body/div[3]/div/div[13]/form/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div/div/input").send_keys(Keys.BACKSPACE)
        driver.find_element(By.XPATH, "/html/body/div[3]/div/div[13]/form/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div/div/input").send_keys(data_hora_ini)

        driver.find_element(By.XPATH, "/html/body/div[3]/div/div[13]/form/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div/div/input").send_keys(Keys.RETURN)
        i = 0
        while i < 14:
            i+=1
            driver.find_element(By.XPATH, "/html/body/div[3]/div/div[13]/form/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div/div/input").send_keys(Keys.BACKSPACE)
        driver.find_element(By.XPATH, "/html/body/div[3]/div/div[13]/form/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div/div/input").send_keys(data_hora_fin)
    elif plat == "MV": 
        driver.find_element(By.XPATH, "/html/body/div[3]/div/div[12]/form/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div/div/input").send_keys(Keys.RETURN)
        i = 0
        while i < 14:
            i+=1
            driver.find_element(By.XPATH, "/html/body/div[3]/div/div[12]/form/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div/div/input").send_keys(Keys.BACKSPACE)
        driver.find_element(By.XPATH, "/html/body/div[3]/div/div[12]/form/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div/div/input").send_keys(data_hora_ini)

        driver.find_element(By.XPATH, "/html/body/div[3]/div/div[12]/form/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div/div/input").send_keys(Keys.RETURN)
        i = 0
        while i < 14:
            i+=1
            driver.find_element(By.XPATH, "/html/body/div[3]/div/div[12]/form/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div/div/input").send_keys(Keys.BACKSPACE)
        driver.find_element(By.XPATH, "/html/body/div[3]/div/div[12]/form/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div/div/input").send_keys(data_hora_fin)


def placa(ind):
    time.sleep(1.5)
    print('\n\n\t===> ', dados['PLACA'][ind])
    print('\t===> ', dados['TURNO'][ind])
    print('\t===> ', dados['MUNICIPIO'][ind])
    print('\t===> ', dados['GRE'][ind])
    print('\tDia:', data[:2] + '/' + data[2:4])
    print('\tFalta:', qtd-1, '\n')
    if plat == 'GS':
        driver.find_element(By.XPATH, "/html/body/div[3]/div/div[13]/form/div[1]/div[2]/div[1]/div[2]/div/div[3]/div/span/span[1]/span").send_keys(Keys.RETURN)
    elif plat == 'MV':
        driver.find_element(By.XPATH, "/html/body/div[3]/div/div[12]/form/div[1]/div[2]/div[1]/div[2]/div/div[3]/div/span/span[1]/span").send_keys(Keys.RETURN)
    driver.find_element(By.XPATH, "/html/body/span/span/span[1]/input").send_keys(dados['PLACA'][ind][-6:])
    driver.find_element(By.XPATH, "/html/body/span/span/span[1]/input").send_keys(Keys.RETURN)



def informativo(ind, inform):
    dir_inf = os.getcwd()
    if dados['GRE'][ind] == '5°':
        dir_inf += "\\INFORMATIVO 5° GRE.xlsx"
        informativo_5 = pd.read_excel(dir_inf, sheet_name="Sheet1")
        for i in range(len(informativo_5)):
            if dados['PLACA'][ind] == informativo_5['PLACA'][i] and dados['TURNO'][ind] == informativo_5['TURNO'][i] and informativo_5['ROTA'][i] == inform:
                val = informativo_5['DIAS'][i]
                y = data[:2]
                x = str(val) + ', ' + y
                informativo_5.loc[i, 'DIAS'] = x
                informativo_5.to_excel(dir_inf, index=False)
    elif dados['GRE'][ind] == '18°':
        dir_inf += "\\INFORMATIVO 18° GRE.xlsx"
        informativo_18 = pd.read_excel(dir_inf, sheet_name="Sheet1")
        for i in range(len(informativo_18)):
            if dados['PLACA'][ind] == informativo_18['PLACA'][i] and dados['TURNO'][ind] == informativo_18['TURNO'][i] and informativo_18['ROTA'][i] == inform:
                val = informativo_18['DIAS'][i]
                y = data[:2]
                x = str(val) + ', ' + y
                informativo_18.loc[i, 'DIAS'] = x
                informativo_18.to_excel(dir_inf, index=False)
    elif dados['GRE'][ind] == '16°':
        dir_inf += "\\INFORMATIVO 16° GRE.xlsx"
        informativo_16 = pd.read_excel(dir_inf, sheet_name="Sheet1")
        for i in range(len(informativo_16)):
            if dados['PLACA'][ind] == informativo_16['PLACA'][i] and dados['TURNO'][ind] == informativo_16['TURNO'][i] and informativo_16['ROTA'][i] == inform:
                val = informativo_16['DIAS'][i]
                y = data[:2]
                x = str(val) + ', ' + y
                informativo_16.loc[i, 'DIAS'] = x
                informativo_16.to_excel(dir_inf, index=False)


def index(i):
    placa(i)

    # Espera para saber se tem rota e ver se o km da pra tirar o arquivo ou não
    if plat == "GS":
        # pega um elemento não carregado ainda só pra esperar os dados da rota aparecer
        driver.find_element(By.XPATH, "/html/body/div[3]/div/div[13]/form/div[1]/div[3]/button[2]").click()
        p = True
        while p:
            driver.implicitly_wait(300)
            plac_plataf = driver.find_element(By.XPATH, "/html/body/div[3]/div/div[13]/div[1]/div[2]/li[3]")
            time.sleep(1.5)
            if plac_plataf.text[-7:] == dados['PLACA'][i][-7:]:
                element = driver.find_element(By.XPATH, "/html/body/div[3]/div/div[13]/div[2]/div[2]/div[1]/div/a")
                p = False
            else:
                p2 = True
                while p2:
                    bo = int(input("\n\t! ERRO NAS PLACAS !\n\nTry again[1]\nNext Route[2]\n>_"))
                    if bo == 1:
                        p = True
                    elif bo == 2:
                        print('\n\tNext Route!')
                        driver.back()
                        permi = False
                        return True
                    else:
                        print('Valor inválido\n')
    elif plat == "MV":
        driver.find_element(By.XPATH, "/html/body/div[3]/div/div[12]/form/div[1]/div[3]/button").click()
        p = True
        while p:
            driver.implicitly_wait(300)
            plac_plataf = driver.find_element(By.XPATH, "/html/body/div[3]/div/div[12]/div[1]/div[2]/li[3]")
            time.sleep(1.5)
            if plac_plataf.text[-7:] == dados['PLACA'][i][-7:]:
                element = driver.find_element(By.XPATH, "/html/body/div[3]/div/div[12]/div[2]/div[2]/div[1]/div/a")
                p = False
            else:
                p2 = True
                while p2:
                    bo = int(input("\n\t! ERRO NAS PLACAS !\n\nTry again[1]\nNext Route[2]\n>_"))
                    if bo == 1:
                        p = True
                    elif bo == 2:
                        print('\n\tNext Route!')
                        driver.back()
                        permi = False
                        return True
                    else:
                        print('Valor inválido\n')
    
    if element.text == '':
        informativo(i, 'NÃO APRESENTA ROTA')
        print('\nNÃO APRESENTA ROTA\n')
        driver.back()
        return True
    else:
        g = element.text
        g = g[:4]
        h = ''
        for k in g:
            if k == '.':
                break
            else:
                h += k
        if int(h) < 2:
            informativo(i, 'NÃO FEZ A ROTA')
            print('\nNÃO FEZ A ROTA\n')
            driver.back()
            return True
        else:
            if plat == "MV":
                driver.find_element(By.XPATH, "/html/body/div[3]/div/div[12]/div[3]/div[2]/div[1]/label[1]/input").click()

            permi = True
            while permi:
                x = str(input('Next[N] | Print[P] | Close[C]\n>_'))
                if x.upper() == 'P':
                    # Pegar o diretório e concatenar com o diretório do arq juntamente com o próprio arq a ser gerado
                    dir_arq = os.getcwd()
                    gre = ''
                    if dados['GRE'][i] == '5°':
                        dir_arq += '\\5° GRE\\'
                    elif dados['GRE'][i] == '18°':
                        dir_arq += '\\18° GRE\\'
                    elif dados['GRE'][i] == '16°':
                        dir_arq += '\\16° GRE\\'
                    dir_arq += dados['MUNICIPIO'][i] +'\\'+ dados['PLACA'][i]+ ' ' + dados['TURNO'][i][0] + ' ' + data[:2] + '.pdf'

                    # printar o arq e mandar para o diretório criado acima
                    print_options = PrintOptions()
                    print_options.page_ranges = ['1-1']
                    with open(dir_arq, 'wb') as f:
                        f.write(b64decode(driver.print_page(print_options)))
                    
                    print('\n\t Arquivo Salvo!')
                    driver.back()
                    permi = False
                    return True
                elif x.upper() == 'N':
                    print('\n\tNext Route!')
                    driver.back()
                    permi = False
                    return True
                elif x.upper() == 'C':
                    permi = False
                else: 
                    print('Valor inválido\n')


print('\n\t==> ! WELCOME ! <==\n\n\t JÁ DEU O GIT PULL?')

log = str(input("\nLogin:\n>_"))
passw = str(input("\nPassword:\n>_"))

if log[:2] == 'gs':
    plat = 'GS'
elif log[:2] == 'ma':
    plat = 'MV'

pe = True
while pe:
    Turno = str(input("\nEscolha o turno:\nManhã[M] - Tarde[T] - Noite[N] - Integral[I]\n>_"))

    if Turno.upper() == "M":
        Turno = 'MANHÃ'
        h_inicial = '000000'
        h_final = '120000'
        pe = False
    elif Turno.upper() == "T":
        Turno = 'TARDE'
        h_inicial = '120000'
        h_final = '180000'
        pe = False
    elif Turno.upper() == "N":
        Turno = 'NOITE'
        h_inicial = '180000'
        h_final = '230000'
        pe = False
    elif Turno.upper() == "I":
        Turno = 'INTEGRAL'
        h_inicial = '070000'
        h_final = '180000'
        pe = False
    else:
        print("Valor inválido")

pe = True
while pe:
    Start = int(input("\nTodas[1] - Específica[2]\n>_"))

    if Start == 1:
        pe = False
    elif Start == 2:
        plac = str(input('\nQual a placa: '))
        pe = False
    else:
        print('Valor inválido')

data = str(input('\n[DDMMNNNN]\nQual a data: '))

# Abre o Navegador
options = Options()
options.add_argument("start-maximized")
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_experimental_option('excludeSwitches', ['enable-automation'])
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.get("https://web.hapolo.com.br/relatorio/?rel=admin")

login()

# Ler o arquivo contendo as placas, turnos e etc...
dir_dados = os.getcwd()
dir_dados += '\\dados_5_16_18_GRE.xlsx'
dados = pd.read_excel(dir_dados, sheet_name="5, 16 e 18 GRE")

# Contar a quantidade de rotas do turno.
# Quando não for rodas todas as placas aqui pega o índice da placa escolhida e conta as rotas restantes.
plac_ind = 0
qtd = 0

if Start == 2:
    for i in range(len(dados)):
        if dados['PLACA'][i] == plac.upper() and dados['TURNO'][i] == Turno:
            plac_ind = i
            break
    for j in range(len(dados)):
        if j >= plac_ind and dados['TURNO'][j] == Turno and dados['PLATAFORMA'][j] == plat:
            qtd += 1
else:
    for i in range(len(dados)):
        if dados['TURNO'][i] == Turno and dados['PLATAFORMA'][i] == plat:
            qtd += 1

# Corpo principal: Horário e troca de rotas
qtd_rota = 0
perm = True
dataHoraInicial = data + h_inicial
dataHoraFinal = data + h_final

for i in range(len(dados)):
    if Start == 2:
        if i >= plac_ind and dados['TURNO'][i][0] == Turno[0] and dados['PLATAFORMA'][i] == plat and perm:
            if qtd_rota == 0:
                dataHora(dataHoraInicial, dataHoraFinal)
                if plat == 'GS':
                    driver.find_element(By.XPATH, "/html/body/div[3]/div/div[13]/form/div[1]/div[2]/div[3]/div[2]/div[11]/div/div[4]/label/div/ins").click()
                elif plat == 'MV':
                    driver.find_element(By.XPATH, "/html/body/div[3]/div/div[12]/form/div[1]/div[2]/div[3]/div[2]/div[11]/div/div[4]/label/div/ins").click()
                perm = index(i)
                qtd_rota = 1
            elif qtd_rota > 0:
                perm = index(i)
            qtd -= 1
    else:
        if dados['TURNO'][i][0] == Turno[0] and dados['PLATAFORMA'][i] == plat and perm:
            if qtd_rota == 0:
                dataHora(dataHoraInicial, dataHoraFinal)
                if plat == 'GS':
                    driver.find_element(By.XPATH, "/html/body/div[3]/div/div[13]/form/div[1]/div[2]/div[3]/div[2]/div[11]/div/div[4]/label/div/ins").click()
                elif plat == 'MV':
                    driver.find_element(By.XPATH, "/html/body/div[3]/div/div[12]/form/div[1]/div[2]/div[3]/div[2]/div[11]/div/div[4]/label/div/ins").click()
                perm = index(i)
                qtd_rota = 1
            elif qtd_rota > 0:
                perm = index(i)
            qtd -= 1

print('\n\t! TURNO COMPLETO !\n\n\tDÊ O GIT PUSH')