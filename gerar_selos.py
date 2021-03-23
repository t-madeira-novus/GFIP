import pyperclip
import pyautogui
import os
import time
from funcoes import *


def _limpar_campo(X, Y):
    pyautogui.moveTo(X, Y)
    pyautogui.rightClick(X, Y)
    print (X, Y)

    time.sleep(10)
    pyautogui.press('down', presses=6)
    pyautogui.press('enter')
    pyautogui.press('delete')

def set_folha_mensal (tipo_folha):
    clicar_pela_imagem("imgs/tipo_folha.png", offsetX=200)
    time.sleep(1)
    if tipo_folha == 'folha mensal':
        pyautogui.press('up', presses=3)
    elif tipo_folha == '13º integral':
        pyautogui.press('down', presses=3)
    pyautogui.press('enter')

def set_modalidade (modalidade):
    if modalidade == '1':
        clicar_pela_imagem('imgs/modalidade.png', offsetX=200)
        clicar_pela_imagem('imgs/modalidade_1.png')



def gerar_selos(empresa, mes, ano, dictionary, prolabore):
    confidence = 0.95
    path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Pessoal\Relatórios Sefip"
    path_aux = ""

    # Clicar na janela da Domínio
    aux = clicar_pela_imagem("imgs/dominio_folha_icon.png")
    if not aux:
        aux = clicar_pela_imagem("imgs/dominio_folha_icon_2.png")
    if not aux:
        pyautogui.alert(text='O módulo do Domínio Folha não foi encontrado. Certifique-se de que ele esteja '
                             'aberto. Se estiver, chame Thiago Madeira para solucionar '
                             'este mistério misterioso.', title='Domínio Folha não encontrado', button='OK')
        return False

    # Pasta Empresa
    try:
        path_aux = path + "\\" + str(empresa) + "-" + dictionary[empresa][0]
    except KeyError:
        file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
        file.write(str(empresa) + "-"  + ": Não tem pasta da empresa\n")
        file.close()
        return False

    if os.path.isdir(path_aux) is False:
        file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
        file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Não tem pasta da empresa\n")
        file.close()
        return False

    # Pasta ano
    path_aux += "\\" + str(ano)
    if os.path.isdir(path_aux) is False:
        os.mkdir(path_aux)

    # Pasta competencia
    path_aux += "\\" + str(mes) + "." + str(ano)
    if os.path.isdir(path_aux) is False:
        os.mkdir(path_aux)

    pyautogui.press('esc', presses=5)
    pyautogui.press('f8') # Abrir busca de empresas
    pyautogui.typewrite(str(empresa)) # Digitar id da empresa
    pyautogui.press('enter')

    # Esperar empresa abrir
    while clicar_pela_imagem("imgs/troca_empresa.png", tentativas=1) or clicar_pela_imagem("imgs/interrogacao.png", tentativas=1):
        pyautogui.hotkey('alt', 'n')
        pyautogui.hotkey('alt', 'n')
        pyautogui.press('esc')

    # Gerar extrato mensal
    # if mes != '13':
    #     clicar_pela_imagem("imgs/relatorios.png")
    #     pyautogui.press('f')sem_dados
    #     pyautogui.press('e')
    #     pyautogui.hotkey('alt', 'o')
    #     while procurar_imagem('imgs/extrato_mensal.png') is False:
    #         time.sleep(1)
    #         pass
    #     clicar_pela_imagem('imgs/pdf_dominio.png')
    #     # Digitar caminho
    #     clicar_pela_imagem('imgs/nome_2.png')
    #     pyperclip.copy(path_aux+'\\Extrato Mensal.pdf')aaaaa
    #     pyautogui.hotkey("ctrl", "v")
    #     pyautogui.press('enter')
    #     pyautogui.hotkey("alt", "s")
    #     while procurar_imagem('imgs/extrato_mensal_2.png') is False:
    #         pass
    #     pyautogui.click(500, 500)
    #     pyautogui.hotkey("ctrl", "w")aaaa
    #     pyautogui.click(500, 500)
    #     pyautogui.press('esc', presses=5)


    # Gerar selo - Navegar Menu
    if clicar_pela_imagem("imgs/relatorios.png") is False:
        file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
        file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Deu pau na hora de clicar em Relatórios\n")
        file.close()
        return False

    if clicar_pela_imagem("imgs/informativos.png") is False:
        file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
        file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Deu pau na hora de clicar em Informativos\n")
        file.close()
        return False

    if clicar_pela_imagem("imgs/mensais.png") is False:
        file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
        file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Deu pau na hora de clicar em Mensais\n")
        file.close()
        return False

    # Gerar selo - Entrar em GFIP
    pyautogui.press('down')
    pyautogui.press('enter')
    time.sleep(1)

    # Limpar campo do caminho
    clicar_pela_imagem('imgs/arquivo.png', offsetX=300)
    # pyautogui.press('tab', presses=7)
    pyautogui.press('apps')
    pyautogui.press('down', presses=6)
    pyautogui.press('enter')
    # Digitar caminho
    print (path_aux)
    pyperclip.copy(path_aux)
    pyautogui.hotkey("ctrl", "v")

    # Ajustar tipo de folha e modalidade
    if prolabore:
        tipo_folha = 'folha mensal'
        modalidade = '1'
    else:
        if mes == '13':
            tipo_folha = '13º integral'
            modalidade = '1'
        elif mes != '13':
            tipo_folha = 'folha mensal'
            modalidade = '0'
    if empresa == 767:
        modalidade = '1'
    # print (empresa, modalidade)
    set_modalidade(modalidade)
    set_folha_mensal(tipo_folha)


    # if mes == '13' and not prolabore: # se for décimo terceiro e não for prolabore
    #     #print ('mes: ', mes)
    #     clicar_pela_imagem("imgs/tipo_folha.png", offsetX=200)
    #     time.sleep(1)
    #     pyautogui.press('down', presses=3)
    #     pyautogui.press('enter')
    #
    # if prolabore or mes == '13':  # Trocar modalidade
    #     clicar_pela_imagem('imgs/modalidade.png', offsetX=200)
    #     clicar_pela_imagem('imgs/modalidade_1.png')
    #
    #     clicar_pela_imagem("imgs/tipo_folha.png", offsetX=200)
    #     time.sleep(1)
    #     pyautogui.press('down', presses=3)
    #     pyautogui.press('enter')


    # Colocar competência
    clicar_pela_imagem('imgs/competencia.png', offsetX=120)
    if mes == '13':
        competencia = str(12)+str(ano)
    else:
        competencia = str(mes)+str(ano)

    pyautogui.press('home')
    pyautogui.typewrite(competencia)

    if str(empresa) in ['13', '879', '684', '767']:

        if clicar_pela_imagem("imgs/codigo_reconhecimento.png") is False:
            file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
            file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Deu pau na hora de clicar em Código de Reconhecimento\n")
            file.close()
            return False

        if clicar_pela_imagem("imgs/mao_obra.png") is False:
            file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
            file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Deu pau na hora de clicar em Mão de Obra\n")
            file.close()
            return False

    elif str(empresa) in ['312', '417', '297', '800', '954']:

        if clicar_pela_imagem("imgs/codigo_reconhecimento.png") is False:
            file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
            file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Deu pau na hora de clicar em Código de Reconhecimento\n")
            file.close()
            return False

        if clicar_pela_imagem("imgs/construcao_civil.png") is False:
            file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
            file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Deu pau na hora de clicar em Mão de Obra\n")
            file.close()
            return False

    if clicar_pela_imagem("imgs/ok.png") is False:
        file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
        file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Deu pau na hora de clicar em OK (imgs/ok.png)\n")
        file.close()
        return False

    if clicar_pela_imagem("imgs/responsavel.png"):
        file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
        file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Deu pau a empresa n tem responsável na domínio\n")
        file.close()
        pyautogui.press('esc')
        clicar_pela_imagem('imgs/responsavel_dominio.png', offsetX=120)
        pyautogui.typewrite('1')
        pyautogui.press('enter')

        clicar_pela_imagem("imgs/ok.png")



    # Esperar gerar
    while   pyautogui.locateOnScreen("imgs/GFIP_gerada_sucesso.png", confidence=confidence) is None and \
            pyautogui.locateOnScreen("imgs/GFIP_gerada_sucesso_2.png", confidence=confidence) is None and \
            pyautogui.locateOnScreen("imgs/GFIP_gerada_sucesso_3.png", confidence=confidence) is None and \
            pyautogui.locateOnScreen("imgs/aviso_windows10.png", confidence=confidence) is None and \
            pyautogui.locateOnScreen("imgs/aviso_windows7.png", confidence=confidence) is None :
        time.sleep(1)
        if clicar_pela_imagem('imgs/sem_dados.png'):
            pyautogui.press('esc', presses=5)
            file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
            file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Sem dados para gerar GFIP na Domínio" + "\n")
            file.close()
            clicar_pela_imagem("imgs/dominio_folha_icon.png")
            return False

    pyautogui.press('esc', presses=5)
    time.sleep(5)

    file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
    file.write(str(empresa) + "-" + dictionary[empresa][0] + ": Sefip.re salvo na pasta" + "\n")
    file.close()

    return True


# clicar_pela_imagem('imgs/modalidade.png', offsetX=200)
# clicar_pela_imagem('imgs/modalidade_1.png')