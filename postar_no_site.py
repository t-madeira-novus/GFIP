import pyautogui
import pyperclip
import os
import time
# import shutil

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as Wait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
# from selenium.webdriver.chrome.options import Options

from funcoes import clicar_pela_imagem, cria_dicionario, clicar_pelo_xpath


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

def get_fgts_namefile (search_dir):
    savedPath = os.getcwd()
    os.chdir(search_dir)
    files = filter(os.path.isfile, os.listdir(search_dir))
    files = [os.path.join(search_dir, f) for f in files] # add path to each file
    files.sort(key=lambda x: os.path.getmtime(x))
    os.chdir(savedPath)

    for file in files:
        file_aux = file.split('\\')[-1]
        if str(file_aux) == "FGTS.pdf" or str(file_aux) == "FGTS.PDF":
            return file_aux
    for file in files:
        file_aux = file.split('\\')[-1]
        if str(file_aux[:3]) == "GRF":
            return file_aux

def clicar_pelo_xpath(browser, xpath):
    time.sleep(1)
    button = Wait(browser, 10).until(ec.presence_of_element_located((By.XPATH, xpath)))
    ActionChains(browser).move_to_element(button).click(button).perform()


def postar_no_site(empresa, mes, ano, dictionary):
    login = "thiagomadeira.novus"
    password = "novus123"

    path = "P:\\documentos\\OneDrive - Novus Contabilidade\\Doc Compartilhado\\Pessoal\\Relatórios Sefip"
    origem = path + "\\" + str(empresa) + "-" + str(dictionary[empresa][0]) + "\\" + ano + "\\" + mes + "." + ano
    try:
        arquivo = origem+"\\"+get_fgts_namefile(origem)
    except TypeError: # can only concatenate str (not "NoneType") to str // Não achou guia do fgts pra postar
        return False
    print ("Arquivo para postar: ", arquivo)

    browser = webdriver.Chrome(executable_path="P:\\documentos\\OneDrive - Novus Contabilidade\\Doc Compartilhado"
                                               "\\Sistemas Internos\\chromedriver.exe")

    browser.get('https://novuscontabilidade.com.br/area-do-cliente/')
    browser.maximize_window()
    # clicar_pelo_xpath(browser, "/html/body/div[2]/div/header/div/div/section[1]/div/div/div[2]/div/div/div/div/ul/li[3]/a/span[2]")
    # clicar_pela_imagem('imgs/area_cliente.png')
    browser.find_element_by_xpath("/html/body/div[1]/div/main/div/div/div/section/div/div/div[3]/div/div/div/div/"
                                  "div/form/div[1]/input").send_keys(login)
    browser.find_element_by_xpath("/html/body/div[1]/div/main/div/div/div/section/div/div/div[3]/div/div/div/div/"
                                  "div/form/div[2]/input").send_keys(password)

    clicar_pelo_xpath(browser, "/html/body/div[3]/div/main/div/div/div/section/div/div/div[3]/div/div/div/div/"
                                "div/form/div[3]/input")  # Acessar
    clicar_pelo_xpath(browser, "/html/body/div[2]/table/tbody/tr[2]/td/table/tbody/"
                                "tr/td/div/ul/li[4]/a")  # Publicação de Documentos
    clicar_pelo_xpath(browser, "/html/body/div[2]/table/tbody/tr[2]/td/table/tbody/tr"
                                "/td/div/ul/li[4]/ul/li[1]/a")  # Documentos

    nome_empresa = dictionary[empresa][0]

    select_clientes = Select(browser.find_element_by_id("publicacaoForm:cliente"))

    for cliente in select_clientes.options:
        if nome_empresa in str(cliente.text):
            cliente.click()
            break

    # pyautogui.keyUp('shift')
    # pyautogui.keyUp('alt')
    # pyautogui.keyUp('ctrl')
    clicar_pelo_xpath(browser, "/html/body/div[3]/div/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table[1]/tbody/tr[2]/td[3]/input")
    clicar_pelo_xpath(browser, "/html/body/div[3]/div/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table[1]/tbody/tr[3]/td[2]/div/select")
    pyautogui.typewrite("Pessoal")
    pyautogui.press("enter")

    clicar_pelo_xpath(browser, "/html/body/div[3]/div/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table[1]/tbody/tr[4]/td[2]/div[1]/select")
    pyautogui.typewrite("/GFIP")
    pyautogui.press("enter")

    clicar_pelo_xpath(browser, "/html/body/div[3]/div/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table[1]/tbody/tr[5]/td[2]/input")
    pyautogui.typewrite(mes+'/'+ano)


    clicar_pelo_xpath(browser, "/html/body/div[3]/div/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table[3]/tbody/tr/td/table/tbody/tr/td/input")  #Anexar documento

    clicar_pela_imagem('imgs\\nome_dominio.png')

    pyautogui.press('tab')
    pyautogui.typewrite(mes+'/'+ano)
    # pyautogui.press(['tab', 'tab', 'space'])

    clicar_pela_imagem("imgs\\site_anexar_arquivo.png")

    clicar_pela_imagem("imgs\\site_escolher_arquivo.png")
    time.sleep(1)
    clicar_pela_imagem("imgs\\nome.png")

    pyperclip.copy(arquivo)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press("enter")

    clicar_pela_imagem("imgs/site_anexar.png")
    pyautogui.press(['tab', 'tab', 'space'])

    time.sleep(7)

    pyautogui.press(['tab', 'space'])
    clicar_pela_imagem("imgs/gravar.png")
    time.sleep(7)
    clicar_pela_imagem("imgs/gravar.png")
    clicar_pela_imagem("imgs/pop_up.png", offsetX=125, offsetY=35)
    # clicar_pela_imagem("imgs/pop_up.png", offsetX=100, offsetY=40)

    time.sleep(7)

    try:
        browser.close()
    except:
        time.sleep(5)
        try:
            browser.switch_to.alert.accept()
        except:
            pass
        browser.close()



    # clicar_pela_imagem("imgs/flash.png")
    # clicar_pela_imagem("imgs/flash_bloqueado.png")
    #
    # pyautogui.press('tab', presses=3)
    # pyautogui.press('enter')
    # clicar_pela_imagem("imgs/switch.png")
    # pyautogui.hotkey("ctrl", 'w')
    # clicar_pela_imagem("imgs/flash.png")
    # time.sleep(1)
    # clicar_pela_imagem("imgs/permitir.png")
    # time.sleep(3)
    #
    #P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Pessoal\Relatórios Sefip\761-BELEZA FEMININA\2020\11.2020\Sefip.re
    # clicar_pela_imagem("imgs/site_pessoal.png")
    # clicar_pela_imagem("imgs/site_expandir_todas.png")



    # time.sleep(60)
    #
    # try:
    #     clicar_pelo_xpath(browser, "/html/body/div[3]/div/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table[2]/tbody/tr/td[1]/div/p/a/img")
    # print(nome_empresa)


# time.sleep(2)
# dictionary = {}
# empresas = [11]
# path = "P:\\documentos\\OneDrive - Novus Contabilidade\\Doc Compartilhado\\Pessoal\\Relatórios Sefip"
# for x in os.listdir(path):
#     for i in range (0, len(empresas)):664
#         if "-" in str(x) and str(empresas[i]) == str(x)[:str(x).find("-")]:
#             dictionary[empresas[i]] = str(x)[str(x).find("-")+1:]
# print (dictionary)



# path = "P:\\documentos\\OneDrive - Novus Conta

