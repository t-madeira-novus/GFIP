import time
import os
import pyperclip
import pyautogui
import shutil

# from pywinauto.application import Application
# from pywinauto.findwindows import ElementNotFoundError
# from pywinauto.findbestmatch import MatchError
# from pywinauto import Desktop


def _colocar_fap():
    _clicar_pela_imagem('imgs/fap_mao.png')
    _clicar_pela_imagem('imgs/fap_dados_do_movimento.png')
    _clicar_pela_imagem('imgs/fap.png', offsetX=150, right=True)
    pyautogui.press('down', presses=5)
    pyautogui.press('enter')
    pyautogui.typewrite('1')
    pyautogui.press('enter', presses=2, interval=1)
    _clicar_pela_imagem('imgs/fap_sair.png')

def _pegar_ultimo_arquivo_modificado(search_dir):
    saved_path = os.getcwd()
    os.chdir(search_dir)
    files = filter(os.path.isfile, os.listdir(search_dir))
    files = [os.path.join(search_dir, f) for f in files]  # add path to each file
    files.sort(key=lambda x: os.path.getmtime(x))
    os.chdir(saved_path)

    return files[-1]


def _clicar_pela_imagem(imagem, offsetX=0, offsetY=0, tentativas=10, right=False):
    conf = 0.95
    aux = None
    c = 0
    while aux is None and c <= tentativas:
        aux = pyautogui.locateCenterOnScreen(imagem, confidence=conf)
        if aux is not None:
            x = aux[0] + offsetX
            y = aux[1] + offsetY
            if right:
                pyautogui.rightClick(x, y)
            else:
                pyautogui.click(x, y)
            return True
        print("Procurando: ", imagem)
        conf -= 0.01
        time.sleep(1)
        c += 1

    return False


def _gerar_sfps(empresa, mes, ano, dictionary, app):
    path = "P:\\documentos\\OneDrive - Novus Contabilidade\\Doc Compartilhado\\Pessoal\\Relatórios Sefip"
    # try:
    #     aplication = Application(backend="uia").connect(best_match='SEFIP - Consulta Cadastro de Responsável')
    #     aplication.top_window().set_focus()
    # except ElementNotFoundError:
    #     app.infoBox("Erooou...", "SEFIP da Caixa não está aberto")
    #     file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
    #     file.write(str(empresa) + "-" + dictionary[empresa] + ": Sefip caixa não estava aberta" + "\n")
    #     file.close()
    #     return False
    # except MatchError:
    #     #erro aqui da sefip estar com menus abertos
    #     try:
    #         aplication = Application(backend="uia").connect(best_match='SEFIP - Consulta Movimento')
    #     except ElementNotFoundError:
    #         app.infoBox("Erooou...", "SEFIP da Caixa não está aberto")
    #         file = open("rel
    #         atorio_gfips_" + mes + "-" + ano + ".txt", "a+")
    #         file.write(str(empresa) + "-" + dictionary[empresa] + ": Sefip caixa não estava aberta" + "\n")
    #         file.close()
    #         return False

    # Clickar na janela da Sefip
    aux = _clicar_pela_imagem("imgs/icone_sefip_2.png")
    if not aux:
        aux = _clicar_pela_imagem("imgs/icone_sefip.png")
    if not aux:
        pyautogui.alert(text='O programa Sefip da Caixa não foi encontrado. Certifique-se de que ele esteja '
                             'aberto. Se estiver, chame Thiago Madeira para solucionar '
                             'este mistério misterioso.', title='Sefip não encontrado', button='OK')
        return False

    path_aux = path + "\\" + str(empresa) + "-" + str(dictionary[empresa]) + "\\" + ano + "\\" + mes + "." + ano + "\\Sefip.re"

    # try:
    #     dlg = aplication.window(best_match='SEFIP - Consulta Cadastro de Responsável')
    #     dlg.set_focus()
    # except MatchError:
    #     dlg = aplication.window(best_match='SEFIP - Consulta Movimento')
    #     dlg.set_focus()

    pyautogui.press('esc', presses=5)
    pyautogui.hotkey("alt", "a")
    pyautogui.press("i")

    time.sleep(3)

    pyperclip.copy(path_aux)
    time.sleep(1)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(1)
    pyautogui.press("enter")

    time.sleep(1)

    if _clicar_pela_imagem("imgs/sim.png") is False:
        # pyautogui.press('enter')
        file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
        file.write(
            str(empresa) + "-" + dictionary[empresa] + ": Deu pau na hora de clicar em Modalidade a ser importada\n")
        file.close()
        return False
    time.sleep(1)
    pyautogui.press("enter")# validacao

    _colocar_fap()

    # if _clicar_pela_imagem("imgs/modalidade_a_ser_importada.png") is False:
    #     pyautogui.press('enter')
    #     file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
    #     file.write(
    #         str(empresa) + "-" + dictionary[empresa] + ": Deu pau na hora de clicar em Modalidade a ser importada\n")
    #     file.close()
    #     return False
    #
    # if _clicar_pela_imagem("imgs/validacao.png") is False:
    #     pyautogui.press('enter')
    #     file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
    #     file.write(
    #         str(empresa) + "-" + dictionary[empresa] + ": Deu pau na hora de clicar em Validacao\n")
    #     file.close()
    #     return False
    #
    # pyautogui.press('enter')
    # time.sleep(1)
    # pyautogui.press('enter')
    # time.sleep(2)
    # pyautogui.press('enter')
    # time.sleep(2)

    if _clicar_pela_imagem("imgs/executar.png") is False:
        file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
        file.write(
            str(empresa) + "-" + dictionary[empresa] + ": Deu pau na hora de clicar em Executar\n")
        file.close()
        return False

    if _clicar_pela_imagem("imgs/compensacao.png"):
        pyautogui.press('enter')

        while _clicar_pela_imagem("imgs/espera_executar.png") is False:
            print("esperando executar")
            pyautogui.press('enter')
            pass

    while _clicar_pela_imagem("imgs/espera_executar.png") is False:
        print("esperando executar")

    # time.sleep(15)

    pyautogui.press('enter', presses=4, interval=1)
    pyautogui.press('tab')
    pyautogui.press('enter')
    pyautogui.press('esc', presses=5, interval=1)

    pasta_caixa_path = 'E:\\Users\\thiago.madeira\\C\\SEFIP'
    origem = _pegar_ultimo_arquivo_modificado(pasta_caixa_path)#("E:\\Users\\"+os.getlogin()+"\\C\\CAIXA\\SEFIP")
    destino = path + "/" + str(empresa) + "-" + str(dictionary[empresa]) + "/" + ano + "/" + mes + "." + ano
    try:
        shutil.move(origem, destino)
    except FileExistsError:
        pass

    file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
    file.write(str(empresa) + "-" + dictionary[empresa] + ": Arquivo .SFP salvo na pasta" + "\n")
    file.close()

    return True
