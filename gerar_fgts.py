import pyautogui
import pyperclip
import os
import time
import shutil
from funcoes import clicar_pela_imagem

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

def gerar_fgts(empresa, mes, ano, dictionary):
    path = "P:\\documentos\\OneDrive - Novus Contabilidade\\Doc Compartilhado\\Pessoal\\Relatórios Sefip"
    destino = path + "\\" + str(empresa) + "-" + str(dictionary[empresa][0]) + "\\" + ano + "\\" + mes + "." + ano
    arquivo_xml = _pegar_arquivo(destino, ".xml")

    while clicar_pela_imagem('imgs/arquivo_icp.png', tentativas = 2) is False:
        clicar_pela_imagem('imgs/relatorios_sefip.png', tentativas = 2)
        clicar_pela_imagem('imgs/grf.png', tentativas = 2)

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

    pyautogui.press("esc", presses=2, interval=1)
    pasta_caixa_path = 'E:\\Users\\thiago.madeira\\C\\SEFIP'
    time.sleep(10)
    origem = _pegar_ultimo_arquivo_modificado(pasta_caixa_path)
    print (origem)

    try:
        shutil.move(origem, destino)
    except:
        os.remove(origem)


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