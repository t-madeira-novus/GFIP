import time
import os
import pyperclip
import pyautogui
import shutil
from funcoes import clicar_pela_imagem

# from pywinauto.application import Application
# from pywinauto.findwindows import ElementNotFoundError
# from pywinauto.findbestmatch import MatchError
# from pywinauto import Desktop


def _colocar_fap(fap):
    clicar_pela_imagem('imgs/fap_mao.png')
    clicar_pela_imagem('imgs/fap_dados_do_movimento.png')
    clicar_pela_imagem('imgs/fap.png', offsetX=150, right=True)
    pyautogui.press('down', presses=5)
    pyautogui.press('enter')
    pyautogui.typewrite(fap)
    pyautogui.press('enter', presses=2, interval=1)
    clicar_pela_imagem('imgs/fap_sair.png')

def _pegar_ultimo_arquivo_modificado(search_dir):
    saved_path = os.getcwd()
    os.chdir(search_dir)
    files = filter(os.path.isfile, os.listdir(search_dir))
    files = [os.path.join(search_dir, f) for f in files]  # add path to each file
    files.sort(key=lambda x: os.path.getmtime(x))
    os.chdir(saved_path)

    return files[-1]


def gerar_sfps(empresa, mes, ano, dictionary, fap, prolabore):
    path = "P:\\documentos\\OneDrive - Novus Contabilidade\\Doc Compartilhado\\Pessoal\\Relatórios Sefip"

    # Clicar na janela da Sefip
    aux = clicar_pela_imagem("imgs/icone_sefip_2.png")
    if not aux:
        aux = clicar_pela_imagem("imgs/icone_sefip.png")
    if not aux:
        pyautogui.alert(text='O programa Sefip da Caixa não foi encontrado. Certifique-se de que ele esteja '
                             'aberto. Se estiver, chame Thiago Madeira para solucionar '
                             'este mistério misterioso.', title='Sefip não encontrado', button='OK')
        return False

    path_aux = path + "\\" + str(empresa) + "-" + str(dictionary[empresa][0]) + "\\" + ano + "\\" + mes + "." + ano + "\\Sefip.re"

    pyautogui.press('esc', presses=5)

    while clicar_pela_imagem('imgs/importar_folha_sefip.png') is False:
        pass
    time.sleep(2)
    clicar_pela_imagem('imgs/nome_2.png')

    pyperclip.copy(path_aux)
    time.sleep(1)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(1)
    pyautogui.press("enter")

    time.sleep(1)

    clicar_pela_imagem("imgs/sim.png")

    if clicar_pela_imagem("imgs/inconsistencia_2.png", tentativas=3):
        pyautogui.hotkey('alt', 'f4')
        pyautogui.press('esc', presses=2)
        file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
        file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Inconsistência ao importar na SEFIP" + "\n")
        file.close()

        pyautogui.hotkey('win', 'd')
        return False # Empresa apresentou erro


    clicar_pela_imagem("imgs/ok_4.png")
    _colocar_fap(fap)

    if prolabore and mes == '13':
        clicar_pela_imagem('imgs/movimento_sefip.png')
        clicar_pela_imagem('imgs/novo.png')
        pyautogui.press('enter')
        pyautogui.typewrite('13'+ano)
        pyautogui.press('tab')
        pyautogui.typewrite('115')
        pyautogui.press('enter')
        clicar_pela_imagem('imgs/ausencia_sefip.png')
        clicar_pela_imagem('imgs/salvar.png')
        pyautogui.press('enter')
        # pyautogui.press('enter')
        time.sleep(2)
        print ('marcar')
        clicar_pela_imagem('imgs/fap_mao.png', right=True)
        time.sleep(1)
        pyautogui.press('down', presses=5)
        pyautogui.press('enter')
        clicar_pela_imagem('imgs/fap_sair.png', right=True)

    clicar_pela_imagem("imgs/executar.png")

    if clicar_pela_imagem("imgs/inconsistencia.png", tentativas=3):
        pyautogui.hotkey('alt', 'f')
        file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
        file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Inconsistência ao importar na SEFIP" + "\n")
        file.close()
        return False # Empresa apresentou erro

    if clicar_pela_imagem("imgs/compensacao.png"):
        pyautogui.press('enter')

    if clicar_pela_imagem("imgs/espera_executar.png"):
        pyautogui.press('enter')

    if clicar_pela_imagem("imgs/espera_executar.png"):
        pyautogui.press('enter')

    clicar_pela_imagem('imgs/salvar.png')

    # time.sleep(15)
    if clicar_pela_imagem("imgs/transmissao_obrigatoria.png"):
        pyautogui.press('enter')

    if clicar_pela_imagem("imgs/transmissao_obrigatoria.png"):
        pyautogui.press('enter')

    pyautogui.press('enter')

    clicar_pela_imagem('imgs/fechar.png')
    pyautogui.hotkey('alt', 'f')
    pyautogui.press('esc', presses=5, interval=1)

    pasta_caixa_path = 'E:\\Users\\thiago.madeira\\C\\SEFIP'
    origem = _pegar_ultimo_arquivo_modificado(pasta_caixa_path)#("E:\\Users\\"+os.getlogin()+"\\C\\CAIXA\\SEFIP")
    destino = path + "\\" + str(empresa) + "-" + dictionary[empresa][0] + "\\" + ano + "\\" + mes + "." + ano

    moveu = False
    while moveu is False:
        try:
            shutil.move(origem, destino)
            moveu = True
        except FileExistsError:
            pass
            moveu = True
        except FileNotFoundError:
            pass
            time.sleep(1)
        # print ('Movendo {} para {}'.format(origem, destino))

    file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
    file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Arquivo .SFP salvo na pasta" + "\n")
    file.close()

    return True

