import time
import pyautogui
import threading
import os
import pyperclip
import shutil

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException


def _pegar_ultimo_arquivo_modificado(search_dir):
    saved_path = os.getcwd()
    os.chdir(search_dir)
    files = filter(os.path.isfile, os.listdir(search_dir))
    files = [os.path.join(search_dir, f) for f in files]  # add path to each file
    files.sort(key=lambda x: os.path.getmtime(x))
    os.chdir(saved_path)

    return files[-1]

def _clicar_pela_imagem (imagem):
    confidence = 0.95
    aux = None
    while aux == None:
        aux = pyautogui.locateCenterOnScreen(imagem, confidence=confidence)
        print("Procurando: ", imagem)
        time.sleep(2)
    pyautogui.click(aux)

def _pegar_arquivo(search_dir, string):
    """Retorna primeiro arquivo contendo string no diretorio passado"""
    saved_path = os.getcwd()
    os.chdir(search_dir)
    files = filter(os.path.isfile, os.listdir(search_dir))
    files = [os.path.join(search_dir, f) for f in files]  # add path to each file
    files.sort(key=lambda x: os.path.getmtime(x))
    os.chdir(saved_path)

    for file in files:
        if string in str(file):
            return file

    return "Nao tem"
    #return files[-1]


def _clicar_pelo_xpath(browser, xpath):
    time.sleep(1)
    button = wait(browser, 30).until(EC.presence_of_element_located((By.XPATH, xpath)))
    ActionChains(browser).move_to_element(button).click(button).perform()


def _selecionar_certificado():
    time.sleep(2)
    pyautogui.press('tab')
    pyautogui.press('enter')
    print("Certificado selecionado")


def _transmitir_gfips(empresa, mes, ano, dictionary, app):
    confidence = 0.8
    # Gerar os paths de cada empresa
    path = "P:\\documentos\\OneDrive - Novus Contabilidade\\Doc Compartilhado\\Pessoal\\Relatórios Sefip"

    # Abrir o site
    browser = webdriver.Ie()
    browser.get('https://conectividade.caixa.gov.br/')
    browser.maximize_window()

    # Thread pra selecionar o certificado
    # thread1 = threading.Thread(target=_selecionar_certificado)
    # thread1.start()

    _clicar_pela_imagem("imgs/site_conectividade/caixa_postal.png")

    path_busca = path + "\\" + str(empresa) + "-" + str(
        dictionary[empresa]) + "\\" + ano + "\\" + mes + "." + ano

    path_aux = _pegar_arquivo(path_busca, ".SFP")
    _clicar_pela_imagem("imgs/site_conectividade/nova_mensagem.png")
    _clicar_pela_imagem("imgs/site_conectividade/box.png")
    pyautogui.press('down')
    pyautogui.press('enter')
    _clicar_pela_imagem("imgs/site_conectividade/continuar.png")

    aux = None
    while aux == None:
        aux = pyautogui.locateCenterOnScreen("imgs/site_conectividade/nova_mensagem_tela_carregada.png", confidence=confidence)
        time.sleep(2)

    pyautogui.click(pyautogui.locateCenterOnScreen("imgs/site_conectividade/selecionar_municipio.png", confidence=confidence))
    pyautogui.typewrite("Juiz de Fora")
    time.sleep(1)
    pyautogui.press("enter")
    # time.sleep(1)
    pyautogui.press("tab")
    # time.sleep(1)
    pyautogui.press("space")
    time.sleep(1)
    pyautogui.click(
    pyautogui.locateCenterOnScreen("imgs/site_conectividade/anexar_arquivo.png", confidence=confidence))

    pyautogui.press('enter')
    time.sleep(2)
    pyperclip.copy(path_aux)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press("enter")

    time.sleep(2)

    pyautogui.click(pyautogui.locateCenterOnScreen("imgs/site_conectividade/salvar.png", confidence=confidence))
    time.sleep(5)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('enter')

    time.sleep(5)
    pyautogui.click(pyautogui.locateCenterOnScreen("imgs/site_conectividade/enviar.png", confidence=confidence))
    time.sleep(3)
    pyautogui.click(pyautogui.locateCenterOnScreen("imgs/site_conectividade/procurar.png", confidence=confidence))

    time.sleep(2)

    path_zip = _pegar_arquivo(path_busca, '.zip')
    pyperclip.copy(path_zip)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press("enter")

    time.sleep(5)

    _clicar_pela_imagem("imgs/site_conectividade/enviar_2.png")

    time.sleep(2)
    _clicar_pela_imagem("imgs/site_conectividade/clique_aqui.png")
    _clicar_pela_imagem("imgs/site_conectividade/salvar_2.png")
    _clicar_pela_imagem("imgs/site_conectividade/salvar_3.png")
    time.sleep(1)
    pyautogui.press('tab')
    pyautogui.press('esc')

    origem = _pegar_ultimo_arquivo_modificado("E:\\Users\\"+os.getlogin()+"\\Downloads")
    destino = path_busca
    print(origem)
    print(destino)
    try:
        shutil.move(origem, destino)
    except:
        pass

    _clicar_pela_imagem("imgs/site_conectividade/salvar_pdf.png")
    _clicar_pela_imagem("imgs/site_conectividade/salvar_3.png")
    time.sleep(1)
    pyautogui.press('tab')
    pyautogui.press('esc')
    pyautogui.hotkey('alt', 'f4')

    origem = _pegar_ultimo_arquivo_modificado("E:\\Users\\"+os.getlogin()+"\\Downloads")
    destino = path_busca
    try:
        shutil.move(origem, destino)
    except:
        pass
    browser.close()

dictionary = {}
path = "P:\\documentos\\OneDrive - Novus Contabilidade\\Doc Compartilhado\\Pessoal\\Relatórios Sefip"
empresas = [4]
for x in os.listdir(path):
    for i in range (0, len(empresas)):
        if "-" in str(x) and str(empresas[i]) == str(x)[:str(x).find("-")]:
            dictionary[empresas[i]] = str(x)[str(x).find("-")+1:]

# _gerar_sfps(4, "08", "2020", dictionary, "teste")