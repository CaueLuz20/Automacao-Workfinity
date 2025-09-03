from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import zipfile
import time
import pyautogui as pya
from datetime import datetime
import os
import calendar

pasta_relatorios = r"c:\Users\Emprasel Daniel\Documents\relatorios"
options = Options()
prefs = {
    "download.default_directory": pasta_relatorios, # pasta de download
    "download.prompt_for_download": False,       # não pedir confirmação
    "directory_upgrade": True,
    "safebrowsing.enabled": True   
}

options.add_experimental_option("prefs", prefs)

navegador = webdriver.Chrome(options=options)
navegador.get("https://tefti.workfinity.com.br/login/auth")
navegador.maximize_window()
navegador.find_element("id", "username").send_keys("Caue.luz")
navegador.find_element("id", "password").send_keys("Caue!@2707")
navegador.find_element("tag name", "button").click()
navegador.find_element("class name", "wkf-brand").click()
esperar = WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable(("xpath", "/html/body/div[2]/div/nav/div/div[2]/div/div[1]/ul/li[2]/a"))).click()
esperar = WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable(("link text","Ordem de Serviço"))).click()
navegador.find_element("xpath", '//*[@id="filtersDiv"]/div[1]/div[3]/div/span/span[1]/span/ul/li/input').send_keys("Encaminhado")
time.sleep(5)
pya.press("enter")
navegador.find_element("xpath", '//*[@id="filtersDiv"]/div[1]/div[3]/div/span/span[1]/span/ul/li/input').send_keys("Reencaminhado")
time.sleep(5)
pya.press("enter")
navegador.find_element("xpath", '//*[@id="filtersDiv"]/div[1]/div[3]/div/span/span[1]/span/ul/li/input').send_keys("Encaminhar")
time.sleep(5)
pya.press("enter")
navegador.find_element("xpath", '//*[@id="filtersDiv"]/div[1]/div[4]/div/span/span[1]/span').click()
navegador.find_element("xpath", '/html/body/span/span/span[1]/input').click()
navegador.find_element("xpath", '/html/body/span/span/span[1]/input').send_keys("ba ogea salvador")
time.sleep(5)
pya.press("enter")
pya.scroll(-400)
navegador.find_element("xpath", '//*[@id="openingDateTo_value"]').click()
navegador.find_element("xpath", '//*[@id="ui-datepicker-div"]/div[3]/button[1]').click()
navegador.find_element("xpath", '//*[@id="ui-datepicker-div"]/div[3]/button[2]').click()

#criação do padrão de tempo
arquivo = "contador.txt"
hoje = datetime.now()

dia_para_clicar = 1
ultima_execucao = ""

# 1. Lê os dados do arquivo de controle
if os.path.exists(arquivo):
    with open(arquivo, "r") as f:
        linhas = f.readlines()
        if len(linhas) >= 2:
            dia_para_clicar = int(linhas[0].strip())
            ultima_execucao = linhas[1].strip()

# 2. Verifica se a data atual é diferente da última execução
if hoje.strftime("%Y-%m-%d") != ultima_execucao:
    # Adiciona esta verificação para evitar o erro "ValueError" na primeira execução
    if ultima_execucao and hoje.month != datetime.strptime(ultima_execucao, "%Y-%m-%d").month:
        dia_para_clicar = 1
    else:
        dia_para_clicar += 1
    
    # Garante que o contador não exceda o último dia do mês
    ultimo_dia_mes = calendar.monthrange(hoje.year, hoje.month)[1]
    if dia_para_clicar > ultimo_dia_mes:
        dia_para_clicar = 1

    # Atualiza a última execução para a data de hoje
    ultima_execucao = hoje.strftime("%Y-%m-%d")
navegador.find_element("xpath", '//*[@id="openingDateFrom_value"]').click()
navegador.find_element("xpath", '//*[@id="ui-datepicker-div"]/div[1]/a[1]/span').click()
try:
    botao_dia = navegador.find_element("xpath", f"//a[text()='{dia_para_clicar}']")
    botao_dia.click()
except Exception as e:
    print(f"⚠️ Não consegui clicar no dia {dia_para_clicar}: {e}")

navegador.find_element("xpath", '//*[@id="ui-datepicker-div"]/div[3]/button[2]').click()
# 6. Atualiza o contador e a data da última execução no arquivo
with open(arquivo, "w") as f:
    f.write(str(dia_para_clicar) + "\n")
    f.write(ultima_execucao)
pya.scroll(-500)
navegador.find_element("xpath", '//*[@id="filters_form"]/button').click()
navegador.find_element("xpath", '//*[@id="resultListTab"]/div/div/form/div[1]/div/div[2]/div/div[3]/button').click()
navegador.find_element("xpath", '//*[@id="resultListTab"]/div/div/form/div[1]/div/div[2]/div/div[3]/ul/li[2]/a').click()
time.sleep(60)
navegador.navigate().refresh()


# parte da extração de relatorio
def esperar_download():
    while any(nome_arquivo.endswith(".crdownload") for nome_arquivo in os.listdir(pasta_relatorios)):
        time.sleep(1)

esperar_download()

arquivo_relatorio = [os.path.join(pasta_relatorios, f) for f in os.listdir(pasta_relatorios)]
ultimo_arquivo = max(arquivo_relatorio, key=os.path.getctime)
if ultimo_arquivo.endswith(".zip"):
    with zipfile.ZipFile(ultimo_arquivo, 'r') as zip_ref:
        zip_ref.extractall(pasta_relatorios + r"\extraido")
        print("Arquivo extraido")

# parte não está funcionando, se faz necessário atualização da página para a extração do link
time.sleep(50)
