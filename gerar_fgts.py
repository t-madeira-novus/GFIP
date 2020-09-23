import pyautogui
import pyperclip
import os
import time
import shutil


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

def _gerar_fgts(empresa, mes, ano, dictionary, app):
    pyautogui.hotkey('alt', 'r')
    pyautogui.press('g')
    pyautogui.press('i')


    path = "P:/documentos/OneDrive - Novus Contabilidade/Doc Compartilhado/Pessoal/Relatórios Sefip"
    origem = path + "/" + str(empresa) + "-" + str(dictionary[empresa]) + "/" + ano + "/" + mes + "." + ano
    arquivo_xml = _pegar_arquivo(origem, ".xml")

    pyperclip.copy(arquivo_xml)
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(3)
    pyautogui.press("p")
    time.sleep(2)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("f")
    time.sleep(1)
    pyautogui.press("esc", presses=2, interval=1)
    pasta_caixa_path = 'C:/Program Files (x86)/CAIXA/SEFIP'
    origem = _pegar_ultimo_arquivo_modificado(pasta_caixa_path)
    destino = path + "/" + str(empresa) + "-" + str(dictionary[empresa]) + "/" + ano + "/" + mes + "." + ano
    shutil.move(origem, destino)

    return True


# time.sleep(2)
# dictionary = {}
# empresas = [11]
# path = "P:\\documentos\\OneDrive - Novus Contabilidade\\Doc Compartilhado\\Pessoal\\Relatórios Sefip"
# for x in os.listdir(path):
#     for i in range (0, len(empresas)):
#         if "-" in str(x) and str(empresas[i]) == str(x)[:str(x).find("-")]:
#             dictionary[empresas[i]] = str(x)[str(x).find("-")+1:]
# print (dictionary)
# _gerar_fgts(11, "08", "2020", dictionary, "hehe")