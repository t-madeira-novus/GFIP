import pyautogui
import pyperclip
import os
import time
import shutil

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as Wait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options


def _esperar(segundos):
    i = 0
    while i < segundos:
        time.sleep(1)
        if i < segundos:
            time.sleep(1)
        i += 1

def _pegar_ultimo_arquivo_modificado (search_dir):
    savedPath = os.getcwd()
    os.chdir(search_dir)
    files = filter(os.path.isfile, os.listdir(search_dir))
    files = [os.path.join(search_dir, f) for f in files] # add path to each file
    files.sort(key=lambda x: os.path.getmtime(x))
    os.chdir(savedPath)

    if "firebird.conf" not in files[-1]:
        return files[-1]
    else:
        return files[-2]

def _clicar_pela_imagem(imagem):
    confidence = 0.95
    aux = None
    tentativas = 0
    while aux == None and tentativas <= 10:
        aux = pyautogui.locateCenterOnScreen(imagem, confidence=confidence)
        if aux != None:
            pyautogui.click(aux)
            return True
        print("Procurando: ", imagem)
        time.sleep(1)
        tentativas += 1

    return False


def _clicar_pelo_xpath(browser, xpath):
    time.sleep(1)
    button = Wait(browser, 30).until(ec.presence_of_element_located((By.XPATH, xpath)))
    ActionChains(browser).move_to_element(button).click(button).perform()


def _postar_no_site(empresa, mes, ano, dictionary, app):
    login = "thiagomadeira.novus"
    password = "novus123"

    browser = webdriver.Chrome(executable_path="P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado"
                                               "\Sistemas Internos\chromedriver.exe")

    browser.get('https://novuscontabilidade.com.br/')
    browser.maximize_window()
    _clicar_pelo_xpath(browser, "/html/body/div[1]/div/header/div/div/section[1]/div/div/div[2]/div/div/div/div/"
                                "ul/li[3]/a/span[2]")

    browser.find_element_by_xpath("/html/body/div[1]/div/main/div/div/div/section/div/div/div[3]/div/div/div/div/"
                                  "div/form/div[1]/input").send_keys(login)
    browser.find_element_by_xpath("/html/body/div[1]/div/main/div/div/div/section/div/div/div[3]/div/div/div/div/"
                                  "div/form/div[2]/input").send_keys(password)

    _clicar_pelo_xpath(browser, "/html/body/div[1]/div/main/div/div/div/section/div/div/div[3]/div/div/div/div/"
                                "div/form/div[3]/input")  # Acessar
    _clicar_pelo_xpath(browser, "/html/body/div[2]/table/tbody/tr[2]/td/table/tbody/"
                                "tr/td/div/ul/li[4]/a")  # Publicação de Documentos
    _clicar_pelo_xpath(browser, "/html/body/div[2]/table/tbody/tr[2]/td/table/tbody/tr"
                                "/td/div/ul/li[4]/ul/li[1]/a")  # Documentos

    nome_empresa = dictionary[empresa]
    select_clientes = Select(browser.find_element_by_id("publicacaoForm:cliente"))

    for cliente in select_clientes.options:
        if nome_empresa in str(cliente.text):
            print ('ok')
            cliente.click()
            break


    _clicar_pelo_xpath(browser, "/html/body/div[3]/div/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table[1]/tbody/tr[2]/td[3]/input")

    _clicar_pelo_xpath(browser, "/html/body/div[3]/div/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table[1]/tbody/tr[3]/td[2]/div/select")
    pyautogui.typewrite("Pessoal")
    pyautogui.press("enter")

    _clicar_pelo_xpath(browser, "/html/body/div[3]/div/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table[1]/tbody/tr[4]/td[2]/div[1]/select")
    pyautogui.typewrite("/GFIP")
    pyautogui.press("enter")

    _clicar_pelo_xpath(browser, "/html/body/div[3]/div/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table[1]/tbody/tr[5]/td[2]/input")
    pyautogui.typewrite(mes+'/'+ano)


    _clicar_pelo_xpath(browser, "/html/body/div[3]/div/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table[3]/tbody/tr/td/table/tbody/tr/td/input")  #Anexar documento

    pyautogui.press('tab')
    pyautogui.typewrite(mes+'/'+ano)
    pyautogui.press(['tab', 'tab', 'space'])

    _clicar_pela_imagem("imgs/site_anexar_arquivo.png")
    _clicar_pela_imagem("imgs/site_escolher_arquivo.png")
    _clicar_pela_imagem("imgs/site_nome.png")

    path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Pessoal\Relatórios Sefip"
    origem = path + "\\" + str(empresa) + "-" + str(dictionary[empresa]) + "\\" + ano + "\\" + mes + "." + ano
    arquivo = _pegar_ultimo_arquivo_modificado(origem)
    pyperclip.copy(arquivo)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press("enter")

    _clicar_pela_imagem("imgs/site_anexar.png")
    _esperar(3)
    # _clicar_pela_imagem("imgs/site_anexar.png")



    # _clicar_pela_imagem("imgs/flash.png")
    # _clicar_pela_imagem("imgs/flash_bloqueado.png")
    #
    # pyautogui.press('tab', presses=3)
    # pyautogui.press('enter')
    # _clicar_pela_imagem("imgs/switch.png")
    # pyautogui.hotkey("ctrl", 'w')
    # _clicar_pela_imagem("imgs/flash.png")
    # time.sleep(1)
    # _clicar_pela_imagem("imgs/permitir.png")
    # time.sleep(3)
    #
    #
    # _clicar_pela_imagem("imgs/site_pessoal.png")
    # _clicar_pela_imagem("imgs/site_expandir_todas.png")



    time.sleep(60)
    #
    # try:
    #     _clicar_pelo_xpath(browser, "/html/body/div[3]/div/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table[2]/tbody/tr/td[1]/div/p/a/img")
    print(nome_empresa)


time.sleep(2)
dictionary = {}
empresas = [11]
path = "P:\\documentos\\OneDrive - Novus Contabilidade\\Doc Compartilhado\\Pessoal\\Relatórios Sefip"
for x in os.listdir(path):
    for i in range (0, len(empresas)):
        if "-" in str(x) and str(empresas[i]) == str(x)[:str(x).find("-")]:
            dictionary[empresas[i]] = str(x)[str(x).find("-")+1:]
print (dictionary)
_postar_no_site(11, "08", "2020", dictionary, "hehe")


