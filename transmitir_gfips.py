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
# from selenium.common.exceptions import TimeoutException

from funcoes import *

def _pegar_ultimo_arquivo_modificado(search_dir):
    saved_path = os.getcwd()
    os.chdir(search_dir)
    files = filter(os.path.isfile, os.listdir(search_dir))
    files = [os.path.join(search_dir, f) for f in files]  # add path to each file
    files.sort(key=lambda x: os.path.getmtime(x))
    os.chdir(saved_path)

    return files[-1]




def _clicar_pelo_xpath(browser, xpath):
    time.sleep(1)
    button = wait(browser, 30).until(EC.presence_of_element_located((By.XPATH, xpath)))
    ActionChains(browser).move_to_element(button).click(button).perform()


def _selecionar_certificado():
    time.sleep(2)
    pyautogui.press('tab')
    pyautogui.press('enter')
    print("Certificado selecionado")

def func_aux():
    time.sleep(5)
    pyautogui.click(300, 300)
    pyautogui.press(['tab', 'enter'])
    time.sleep(1)
    pyautogui.press('enter')


def transmitir_gfips(empresa, mes, ano, dictionary):
    confidence = 0.8
    # Gerar os paths de cada empresa
    path = "P:\\documentos\\OneDrive - Novus Contabilidade\\Doc Compartilhado\\Pessoal\\Relatórios Sefip"

    # Abrir o site
    browser = webdriver.Ie()
    time.sleep(2)
    browser.get('https://www.conectividade.caixa.gov.br/')

    while clicar_pela_imagem("imgs/site_conectividade/caixa_postal.png") is False:
        func_aux()
        pass
    clicar_pela_imagem("imgs/site_conectividade/nova_mensagem.png")

    clicar_pela_imagem("imgs/site_conectividade/box.png")
    time.sleep(1)
    pyautogui.press(['tab', 'down'], interval=1)
    # pyautogui.press('down')
    pyautogui.press('enter')
    clicar_pela_imagem("imgs/site_conectividade/continuar.png")

    aux = None
    while aux == None:
        aux = pyautogui.locateCenterOnScreen("imgs/site_conectividade/nova_mensagem_tela_carregada.png", confidence=confidence)
        time.sleep(2)

    pyautogui.click(pyautogui.locateCenterOnScreen("imgs/site_conectividade/selecionar_municipio.png", confidence=confidence))
    time.sleep((1))
    pyautogui.typewrite("Juiz de Fora")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("tab")
    time.sleep(1)
    pyautogui.press("space")
    time.sleep(1)
    pyautogui.click(
    pyautogui.locateCenterOnScreen("imgs/site_conectividade/anexar_arquivo.png", confidence=confidence))

    pyautogui.press('enter')
    time.sleep(2)

    path_busca = path + "\\" + str(empresa) + "-" + str(dictionary[empresa][0]) + "\\" + ano + "\\" + mes + "." + ano

    try:
        path_aux = pegar_arquivo(path_busca, ".SFP")
    except FileNotFoundError:
        time.sleep(2)
        path_aux = pegar_arquivo(path_busca, ".SFP")

    pyperclip.copy(path_aux)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press("enter")

    if clicar_pela_imagem('imgs/site_conectividade/erro_responsavel.png'):

        file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
        file.write(
            str(empresa) + "-" + dictionary[empresa] + ": Pulou pq Só é permitido envio de arquivo cujo responsável seja uma inscrição... \n")
        file.close()
        return False

    while clicar_pela_imagem('imgs/site_conectividade/municipio_recolhimento.png') is False:
        time.sleep(1)

    pyautogui.click(pyautogui.locateCenterOnScreen("imgs/site_conectividade/salvar.png", confidence=confidence))
    time.sleep(10)
    pyautogui.press('enter')
    time.sleep(3)
    pyautogui.press('enter')
    time.sleep(8)

    clicar_pela_imagem("imgs/site_conectividade/enviar.png")

    if clicar_pela_imagem('imgs/instabilidade_conectividade.png', tentativas=3):

        file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
        file.write(
            #str(empresa) + "-" + dictionary[empresa] + ": Pulou pq o site da caixa conectividade ruim\n")
            str(empresa) + "-" + ": Pulou pq o site da caixa conectividade ruim\n")
        file.close()
        pyautogui.press('esc', presses=3)
        browser.close()
        return False


    while clicar_pela_imagem("imgs/site_conectividade/procurar.png") is False:
        clicar_pela_imagem("imgs/site_conectividade/enviar.png")

    path_zip = pegar_arquivo(path_busca, '.zip')
    if path_zip == "Nao tem":
        path_zip = pegar_arquivo(path_busca, '.zip')
    if path_zip == "Nao tem":

        file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
        file.write(
            str(empresa) + "-" + dictionary[empresa] + ": Pulou pq n achou o arquivo .zip \n")
        file.close()

        pyautogui.press('esc', presses=3)
        browser.close()

        return False

    pyperclip.copy(path_zip)
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press("enter")

    while clicar_pela_imagem("imgs/site_conectividade/enviar_2.png", tentativas=2) is False:
        pass

    time.sleep(2)
    if clicar_pela_imagem("imgs/site_conectividade/clique_aqui.png", tentativas=5) is False:
        clicar_pela_imagem("imgs/site_conectividade/clique_aqui_roxo.png")
    if clicar_pela_imagem("imgs/site_conectividade/salvar_2.png") is False:
        time.sleep(2)

        pyautogui.click(172, 883)
        pyautogui.click(315, 786)

        time.sleep(2)

        # file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
        # file.write(
        #     str(empresa) + "-" + dictionary[empresa] + ": Pulou pq deu problema no fluxo do site, tenho q arrumar um jeito de arrumar o alt tab depois do clique aqqui\n")
        # file.close()
        # return False


    time.sleep(2)
    while clicar_pela_imagem("imgs/site_conectividade/salvar_3.png") is False:
        pass
    time.sleep(2)
    pyautogui.press('tab')
    pyautogui.press('esc')
    time.sleep(2)

    origem = _pegar_ultimo_arquivo_modificado("E:\\Users\\"+os.getlogin()+"\\Downloads")
    destino = path_busca
    # print(origem)
    # print(destino)
    try:
        shutil.move(origem, destino)
    except:
        pass

    while clicar_pela_imagem("imgs/site_conectividade/salvar_pdf.png") is False:
        pass
    time.sleep(2)
    while clicar_pela_imagem("imgs/site_conectividade/salvar_3.png") is False:
        pass

    time.sleep(4)
    pyautogui.press('tab')
    pyautogui.press('esc')
    browser.close()
    pyautogui.hotkey('alt', 'f4')

    origem = _pegar_ultimo_arquivo_modificado("E:/Users/"+os.getlogin()+"/Downloads")
    destino = path_busca
    try:
        shutil.move(origem, destino)
    except:
        pass

dictionary = {}
# path = "P:\\documentos\\OneDrive - Novus Contabilidade\\Doc Compartilhado\\Pessoal\\Relatórios Sefip"
# empresas = [4]
# for x in os.listdir(path):
#     for i in range (0, len(empresas)):
#         if "-" in str(x) and str(empresas[i]) == str(x)[:str(x).find("-")]:
#             dictionary[empresas[i]] = str(x)[str(x).find("-")+1:]

# _gerar_sfps(4, "08", "2020", dictionary, "teste")
# transmitir_gfips(4, "08", "2020", dictionary)