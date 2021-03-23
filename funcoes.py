import os
import pyautogui
import time


def clicar_pelo_xpath(browser, xpath):
    button = wait(browser, 60).until(EC.presence_of_element_located((By.XPATH, xpath)))
    try:
        ActionChains(browser).move_to_element(button).click(button).perform()
    except:
        time.sleep(1)
        clicar_pelo_xpath(browser, xpath)

def cria_dicionario(empresas, faps, prolabores, path):
    dicionario = {}

    for x in os.listdir(path):
        for i in range(0, len(empresas)):
            if "-" in str(x) and str(empresas[i]) == str(x)[:str(x).find("-")]:
                dicionario[empresas[i]] = [str(x)[str(x).find("-") + 1:], faps[i], prolabores[i]]
    return dicionario


def procurar_imagem(imagem, tentativas=10):
    conf = 0.95
    aux = None
    c = 0
    # print("Procurando: ", imagem, end="... ")
    while aux is None and c <= tentativas:
        try:
            aux = pyautogui.locateCenterOnScreen(imagem, confidence=conf)
        except OSError:  # screen grab failed
            c = 0
            print("Screen grab error")
        if aux is not None:
            #print('Achou!')
            return True
        conf -= 0.01
        time.sleep(1)
        c += 1
   # print("Não achou!")
    return False


def clicar_pela_imagem(imagem, offsetX=0, offsetY=0, tentativas=10, right=False):
    conf = 0.9
    aux = None
    c = 0
    print("Procurando: ", imagem, end="... ")
    while aux is None and c <= tentativas:
        try:
            aux = pyautogui.locateCenterOnScreen(imagem, confidence=conf)
        except OSError:  # screen grab failed
            c = 0
            print("Screen grab error ", imagem)
        if aux is not None:
            x = aux[0] + offsetX
            y = aux[1] + offsetY
            if right:
                pyautogui.rightClick(x, y)
            else:
                pyautogui.click(x, y)
                print('Achou!')
            return True
        conf -= 0.01
        time.sleep(1)
        c += 1
    print("Não achou!")
    return False


def pegar_arquivo(search_dir, string):
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