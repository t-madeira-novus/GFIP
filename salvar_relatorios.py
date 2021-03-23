import pyperclip
import pyautogui
import os
import time
import shutil

from shutil import Error
from funcoes import *
# from pywinauto.application import Application
# from pywinauto.findwindows import ElementNotFoundError
# from pywinauto.findbestmatch import MatchError


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


def _limpar_campo(X, Y):
    pyautogui.doubleClick(X, Y)
    pyautogui.press('home')
    pyautogui.keyDown('shift')
    pyautogui.press('end')
    pyautogui.keyUp('shift')

def _abre_relatorios():
    pyautogui.press('esc', presses=3)
    pyautogui.hotkey('alt', 'r')
    pyautogui.press('down')
    pyautogui.press('right')

def _salva(destino, atalho_pdf = 'g'):
    if clicar_pela_imagem('imgs/gerar.png'):
        pyautogui.press('enter', presses=2, interval=3)
        clicar_pela_imagem('imgs/fechar_sefip.png')

        pasta_caixa_path = 'E:\\Users\\thiago.madeira\\C\\SEFIP'
        origem = _pegar_ultimo_arquivo_modificado(pasta_caixa_path)
        try:
            shutil.move(origem, destino)
        except Error:
            pass
        except FileNotFoundError:
            return False

        _abre_relatorios()
        return True
    elif clicar_pela_imagem('imgs/gerar_2.png'):
        pyautogui.press('enter', presses=2, interval=3)
        clicar_pela_imagem('imgs/fechar_sefip.png')

        pasta_caixa_path = 'E:\\Users\\thiago.madeira\\C\\SEFIP'
        origem = _pegar_ultimo_arquivo_modificado(pasta_caixa_path)
        try:
            shutil.move(origem, destino)
        except Error:
            pass
        except FileNotFoundError:
            return False

        _abre_relatorios()
        return True
    return False

def salvar_relatorios(empresa, mes, ano, dictionary, prolabore):
    path = "P:\\documentos\\OneDrive - Novus Contabilidade\\Doc Compartilhado\\Pessoal\\Relatórios Sefip"
    destino = path + "\\" + str(empresa) + "-" + str(dictionary[empresa][0]) + "\\" + ano + "\\" + mes + "." + ano

    # Clickar na janela da Sefip
    time.sleep(1)
    pyautogui.hotkey('winleft', 'd')
    time.sleep(1)
    aux = clicar_pela_imagem("imgs/icone_sefip_2.png")
    if not aux:
        aux = clicar_pela_imagem("imgs/icone_sefip.png")
    if not aux:
        pyautogui.alert(text='O programa Sefip da Caixa não foi encontrado. Certifique-se de que ele esteja '
                             'aberto. Se estiver, chame Thiago Madeira para solucionar '
                             'este mistério misterioso.', title='Sefip não encontrado', button='OK')
        return False

    #Salvar documento de que prolabore nao tem fgts pra decimo terceiro
    if prolabore and mes == '13':
        print ('entrou no if')
        clicar_pela_imagem('imgs/relatorios_sefip.png')
        pyautogui.press(['p', 'i', 'c'])
        clicar_pela_imagem('imgs/nome.png')
        pyperclip.copy(pegar_arquivo(destino, '.xml'))
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        _salva(destino)

    pyautogui.press('esc', presses=5, interval=1)
    _abre_relatorios()

    # Salvar arquivos
    teve_erro = False
    if clicar_pela_imagem('imgs/re.png'):
        if _salva(destino):
            file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
            file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Relatório RE salvo na pasta" + "\n")
            file.close()
    if clicar_pela_imagem('imgs/comprovante_declaracao.png'):
        if _salva(destino):
            file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
            file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Relatório comprovante_declaracao salvo na pasta" + "\n")
            file.close()
    if clicar_pela_imagem('imgs/gps.png'):
        if _salva(destino):
            file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
            file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Relatório gps salvo na pasta" + "\n")
            file.close()
    if clicar_pela_imagem('imgs/gps_rs.png'):
        if _salva(destino):
            file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
            file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Relatório gps_rs salvo na pasta" + "\n")
            file.close()
    if clicar_pela_imagem('imgs/analitico_gps.png'):
        if _salva(destino):
            file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
            file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Relatório analitico_gps salvo na pasta" + "\n")
            file.close()
    if clicar_pela_imagem('imgs/compensacao_relatorio.png'):
        if _salva(destino):
            file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
            file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Relatório compensacao_relatorio salvo na pasta" + "\n")
            file.close()
    if clicar_pela_imagem('imgs/analitico_grf.png'):
        time.sleep(1)
        if _salva(destino):
            file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
            file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Relatório grf salvo na pasta" + "\n")
            file.close()

    # Fazer backup
    clicar_pela_imagem('imgs/ferramentas.png', tentativas=2)
    while clicar_pela_imagem('imgs/fazer_backup.png', tentativas=2) is False:
        clicar_pela_imagem('imgs/ferramentas.png', tentativas=2)

    clicar_pela_imagem('imgs/ok_backup.png')
    pyautogui.press('home')
    pyautogui.press ('del', presses=20)
    destino = path + "\\" + str(empresa) + "-" + str(dictionary[empresa][0]) + "\\" + ano + "\\" + mes + "." + ano
    destino_aux = destino + "\\SEFIPBKP.ZIP"
    pyperclip.copy(destino_aux)
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'v')

    if clicar_pela_imagem("imgs/salvar_sefip.png") is False:
        file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
        file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Deu pau no programa da SEFIP\n")
        file.close()
        return False

    while clicar_pela_imagem('imgs/espera_executar.png') is False:
        pass
    pyautogui.press('enter')


    file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
    file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Fim" + "\n")
    file.close()

    if teve_erro:
        return False

# time.sleep(1)
# empresa = '669'
# mes = '02'
# ano = '2021'
# prolabore = False
# path = "P:\\documentos\\OneDrive - Novus Contabilidade\\Doc Compartilhado\\Pessoal\\Relatórios Sefip"
# dictionary = cria_dicionario([empresa], [1], [True], path)
# salvar_relatorios(empresa, mes, ano, dictionary, prolabore)