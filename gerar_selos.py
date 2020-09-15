import pyperclip
import pyautogui
import os
import time

def _clicar_pela_imagem (imagem):
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

def _limpar_campo(X, Y):
    pyautogui.moveTo(X, Y)
    pyautogui.rightClick(X, Y)
    print (X, Y)

    time.sleep(10)
    pyautogui.press('down', presses=6)
    pyautogui.press('enter')
    pyautogui.press('delete')


def _gerar_selos(empresa, mes, ano, dictionary, app):
    confidence = 0.95
    path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Pessoal\Relatórios Sefip"

    # Clickar na janela da Damínio
    aux = _clicar_pela_imagem("imgs/dominio_folha_icon.png")
    if not aux:
        aux = _clicar_pela_imagem("imgs/dominio_folha_icon_2.png")
    if not aux:
        pyautogui.alert(text='O módulo do Domínio Folha não foi encontrado. Certifique-se de que ele esteja '
                             'aberto. Se estiver, chame Thiago Madeira para solucionar '
                             'este mistério misterioso.', title='Domínio Folha não encontrado', button='OK')
        return False

    try:
        path_aux = path + "\\" + str(empresa) + "-" + str(dictionary[empresa]) + "\\" + str(ano) + "\\" + str(mes) + "." + str(
            ano)
    except KeyError:
        file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
        file.write(str(empresa) + ": Não existe pasta" + "\n")
        file.close()
        pyautogui.hotkey('alt', 'tab')
        app.infoBox("Error", "Não existe pasta para a empresa " + str(empresa) + ". Confira relatório gerado.")
        return False

    try:
        os.mkdir(path_aux)
    except FileExistsError:
        pass

    pyautogui.press('esc', presses=5)
    pyautogui.press('f8') # Abrir busca de empresas
    pyautogui.typewrite(str(empresa)) # Digitar id da empresa
    pyautogui.press('enter')
    time.sleep(10) # Esperar empresa carregar

    if _clicar_pela_imagem("imgs/relatorios.png") is False:
        file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
        file.write(str(empresa) + "-" + dictionary[empresa] + ": Deu pau na hora de clicar em Relatórios\n")
        file.close()
        return False

    if _clicar_pela_imagem("imgs/informativos.png") is False:
        file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
        file.write(str(empresa) + "-" + dictionary[empresa] + ": Deu pau na hora de clicar em Informativos\n")
        file.close()
        return False

    if _clicar_pela_imagem("imgs/mensais.png") is False:
        file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
        file.write(str(empresa) + "-" + dictionary[empresa] + ": Deu pau na hora de clicar em Mensais\n")
        file.close()
        return False

    pyautogui.press('down')
    pyautogui.press('enter') # Entrar em GFIP
    time.sleep(1)

    # Limpar campo do caminho
    pyautogui.press('tab', presses=7)
    pyautogui.press('apps')
    pyautogui.press('down', presses=6)
    pyautogui.press('enter')
    # Digitar caminho
    pyperclip.copy(path_aux)
    pyautogui.hotkey("ctrl", "v")

    if empresa in [13, 879, 297, 312, 417, 800]:

        if _clicar_pela_imagem("imgs/codigo_reconhecimento.png") is False:
            file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
            file.write(str(empresa) + "-" + dictionary[empresa] + ": Deu pau na hora de clicar em Código de Reconhecimento\n")
            file.close()
            return False

        if _clicar_pela_imagem("imgs/mao_obra.png") is False:
            file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
            file.write(str(empresa) + "-" + dictionary[empresa] + ": Deu pau na hora de clicar em Mão de Obra\n")
            file.close()
            return False

    if _clicar_pela_imagem("imgs/ok.png") is False:
        file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
        file.write(str(empresa) + "-" + dictionary[empresa] + ": Deu pau na hora de clicar em OK (imgs/ok.png)\n")
        file.close()
        return False

    # Esperar gerar
    while   pyautogui.locateOnScreen("imgs/GFIP_gerada_sucesso.png", confidence=confidence) is None and \
            pyautogui.locateOnScreen("imgs/GFIP_gerada_sucesso_2.png", confidence=confidence) is None and \
            pyautogui.locateOnScreen("imgs/GFIP_gerada_sucesso_3.png", confidence=confidence) is None and \
            pyautogui.locateOnScreen("imgs/aviso_windows10.png", confidence=confidence) is None and \
            pyautogui.locateOnScreen("imgs/aviso_windows7.png", confidence=confidence) is None :
        time.sleep(1)

    pyautogui.press('esc', presses=5)
    time.sleep(5)

    file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
    file.write(str(empresa) + "-" + dictionary[empresa] + ": Sefip.re salvo na pasta" + "\n")
    file.close()

    return True

