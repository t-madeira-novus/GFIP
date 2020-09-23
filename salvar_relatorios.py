import pyperclip
import pyautogui
import os
import time
import shutil

from shutil import Error
# from pywinauto.application import Application
# from pywinauto.findwindows import ElementNotFoundError
# from pywinauto.findbestmatch import MatchError


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


def _salvar_relatorios(empresa, mes, ano, dictionary, app):
    path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Pessoal\Relatórios Sefip"
    # try:
    #     application = Application(backend="uia").connect(best_match='SEFIP - Consulta Cadastro de Responsável')
    #     application.top_window().set_focus()
    #
    # except ElementNotFoundError:
    #     app.infoBox("Erooou...", "SEFIP da Caixa não está aberto")
    #     file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
    #     file.write(str(empresa) + "-" + dictionary[empresa] + ": Sefip caixa não estava aberta" + "\n")
    #     file.close()
    #     return False
    #
    # except MatchError:
    #     try:
    #         application = Application(backend="uia").connect(best_match='SEFIP - Consulta Movimento')
    #     except ElementNotFoundError:
    #         app.infoBox("Erooou...", "SEFIP da Caixa não está aberto")
    #         file = open("relatorio_gfips_" + mes + "-" + ano + ".txt", "a+")
    #         file.write(str(empresa) + "-" + dictionary[empresa] + ": Sefip caixa não estava aberta" + "\n")
    #         file.close()
    #         return False

    # # Clickar na janela da Damínio
    # aux = _clicar_pela_imagem("imgs/icone_sefip.png")
    # if not aux:
    #     aux = _clicar_pela_imagem("imgs/icone_sefip_2.png")
    # if not aux:
    #     pyautogui.alert(text='O módulo do Domínio Folha não foi encontrado. Certifique-se de que ele esteja '
    #                          'aberto. Se estiver, chame Thiago Madeira para solucionar '
    #                          'este mistério misterioso.', title='Domínio Folha não encontrado', button='OK')
    #     return False

    # try:
    #     dlg = application.window(best_match='SEFIP - Consulta Cadastro de Responsável')
    #     dlg.set_focus()
    # except MatchError:
    #     dlg = application.window(best_match='SEFIP - Consulta Movimento')
    #     dlg.set_focus()

    # Clickar na janela da Sefip
    pyautogui.hotkey('winleft', 'd')
    aux = _clicar_pela_imagem("imgs/icone_sefip_2.png")
    if not aux:
        aux = _clicar_pela_imagem("imgs/icone_sefip.png")
    if not aux:
        pyautogui.alert(text='O programa Sefip da Caixa não foi encontrado. Certifique-se de que ele esteja '
                             'aberto. Se estiver, chame Thiago Madeira para solucionar '
                             'este mistério misterioso.', title='Sefip não encontrado', button='OK')
        return False

    pyautogui.press('esc', presses=5, interval=1)

    # Salvar arquivos
    # Analitico GRF
    # pyautogui.hotkey('alt', 'r')
    # pyautogui.press('down')
    # pyautogui.press('right')
    # pyautogui.press('enter')
    #
    # time.sleep(1)
    # _clicar_pela_imagem("imgs\gerar.png")
    # pyautogui.press('enter')
    # _clicar_pela_imagem("imgs\ok_4.png")
    # pyautogui.press('enter')
    # time.sleep(1)
    #
    # pasta_caixa_path = 'C:\Program Files (x86)\CAIXA\SEFIP'#"E:\\Users\\"+os.getlogin()+"\\C\\CAIXA\\SEFIP"
    # origem = _pegar_ultimo_arquivo_modificado(pasta_caixa_path)
    # destino = path + "/" + str(empresa) + "-" + str(dictionary[empresa]) + "/" + ano + "/" + mes + "." + ano
    # try:
    #     shutil.move(origem, destino)
    # except Error:
    #     pass
    #
    # pyautogui.hotkey('alt', 'f')

    # Comprovante de Declaração à Previdência

    pyautogui.hotkey('alt', 'r')
    pyautogui.press('down')
    pyautogui.press('right')
    time.sleep(2)
    _clicar_pela_imagem("imgs/comprovante_declaracao.png")
    pyautogui.hotkey('alt', 'g')
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.press('enter')
    pyautogui.hotkey('alt', 'f')
    pyautogui.press('enter')
    time.sleep(1)

    pasta_caixa_path = 'E:\\Users\\thiago.madeira\\C\\SEFIP'
    origem = _pegar_ultimo_arquivo_modificado(pasta_caixa_path)  # ("E:\\Users\\"+os.getlogin()+"\\C\\CAIXA\\SEFIP")
    destino = path + "/" + str(empresa) + "-" + str(dictionary[empresa]) + "/" + ano + "/" + mes + "." + ano
    try:
        shutil.move(origem, destino)
    except Error:
        pass

    time.sleep(1)

    # GPS
    pyautogui.press('esc', presses=3)
    pyautogui.hotkey('alt', 'r')
    pyautogui.press('down')
    pyautogui.press('right')

    _clicar_pela_imagem("imgs/gps.png")
    pyautogui.hotkey('alt', 'p')
    time.sleep(1)

    _clicar_pela_imagem("imgs\ok_3.png")
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)

    origem = _pegar_ultimo_arquivo_modificado(pasta_caixa_path)
    try:
        shutil.move(origem, destino)
    except Error:
        pass

    pyautogui.hotkey('alt', 'f')
    time.sleep(1)

    # Analítico
    pyautogui.press('down', presses=3)
    pyautogui.press('enter')
    pyautogui.hotkey('alt', 'g')
    # pyautogui.click(pyautogui.locateCenterOnScreen("imgs\ok_3.png", confidence=confidence))
    _clicar_pela_imagem("imgs\ok_3.png")
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)

    origem = _pegar_ultimo_arquivo_modificado(pasta_caixa_path)
    try:
        shutil.move(origem, destino)
    except Error:
        pass

    pyautogui.hotkey('alt', 'f')

    pyautogui.press('esc', presses=2)

    # RE
    pyautogui.press('esc', presses=3)
    pyautogui.hotkey('alt', 'r')
    pyautogui.press('down')
    pyautogui.press('right')

    _clicar_pela_imagem("imgs/re.png")
    pyautogui.hotkey('alt', 'g')
    pyautogui.press('enter')

    _clicar_pela_imagem("imgs\ok_4.png")
    time.sleep(1)

    origem = _pegar_ultimo_arquivo_modificado(pasta_caixa_path)
    try:
        shutil.move(origem, destino)
    except Error:
        pass

    pyautogui.hotkey('alt', 'f')
    time.sleep(1)
    pyautogui.press('esc', presses=3)



    # dictionary = {}
    # for x in os.listdir(path):
    #     for i in range(0, len(empresas)):
    #         if "-" in str(x) and str(empresas[i]) == str(x)[:str(x).find("-")]:
    #             dictionary[empresas[i]] = str(x)[str(x).find("-") + 1:]
    # path = "P:\\documentos\\OneDrive - Novus Contabilidade\\Doc Compartilhado\\Pessoal\\Relatórios Sefip"

    path = "P:\\documentos\\OneDrive - Novus Contabilidade\\Doc Compartilhado\\Pessoal\\Relatórios Sefip"
    # for empresa in empresas:
    #     if empresa in df["Codigos Empresas"].tolist():
    #
    #         # path_busca = path + "\\" + str(empresa) + "-" + str(
    #         #     dictionary[empresa]) + "\\" + ano + "\\" + mes + "." + ano
    #
    #         i = df["Codigos Empresas"].tolist().index(empresa)
    #         # destino = str(df.at[i, "Caminhos"]+str)
    #
    #         # Clicar em Arquivo e Importar Folha
    #         # pyautogui.hotkey("alt", "a")
    #         # pyautogui.press('i')
    #         # time.sleep(1)
    #
    #         # Digitar caminho do selo
    #         # kb.type_this(str(df.at[i, "Caminhos"])+"\Sefip.re")
    #         # pyautogui.typewrite(str(df.at[i, "Caminhos"])+"\Sefip.re")
    #         # print(str(df.at[i, "Caminhos"])+"\Sefip.re")
    #         # aux = pyautogui.locateCenterOnScreen("imgs/nome.png", confidence=confidence-0.2)
    #         # pyautogui.click(aux[0]+400, aux[1])
    #         # pyperclip.copy(str(df.at[i, "Caminhos"])+"\Sefip.re")
    #         # pyautogui.hotkey("ctrl", "v")
    #         # pyautogui.press('enter')
    #         # time.sleep(3)
    #
    #         # if pyautogui.locateCenterOnScreen("imgs/confirmar_repeticao.png", confidence=confidence) != None:
    #         #     pyautogui.press('enter')
    #
    #         # if pyautogui.locateCenterOnScreen("imgs/confirmar_importacao_1.png", confidence=confidence) != None\
    #         # or pyautogui.locateCenterOnScreen("imgs/confirmar_importacao_2.png", confidence=confidence) != None\
    #         # or pyautogui.locateCenterOnScreen("imgs/confirmar_importacao_3.png", confidence=confidence) != None:
    #         #
    #         #     pyautogui.press('enter')
    #         #     time.sleep(2)
    #         #     pyautogui.press('enter')
    #         #     time.sleep(2)
    #
    #         # Clicar em Executar e salvar arquivo no diretorio padrao
    #         #pyautogui.click(pyautogui.locateCenterOnScreen("imgs/executar.png", confidence=confidence))
    #         # _clicar_pela_imagem("imgs/executar.png")
    #         # time.sleep(7)
    #         # pyautogui.press('enter')
    #         # time.sleep(1)
    #         # pyautogui.press('enter')
    #         # time.sleep(1)
    #         # pyautogui.press('enter')
    #         # time.sleep(1)
    #         # pyautogui.press('enter')
    #         # time.sleep(1)
    #         # pyautogui.press('enter')
    #         # time.sleep(1)
    #         #
    #         # # Mover arquivo para diretorio da empresa
    #         # origem = _pegar_ultimo_arquivo_modificado(pasta_caixa_path)
    #         # destino = str(df.at[i, "Caminhos"])
    #         # destino += "\\" + ano + "\\" + mes + "." + ano
    #         # shutil.move(origem, destino)
    #
    #         # Importar folha
    #         pyautogui.hotkey('alt', 'a')
    #         pyautogui.press('i')
    #         time.sleep(1)
    #
    #         path_aux = str(df.at[i, "Caminhos"]) + "\\" + ano + "\\" + mes + "." + ano + "\\Sefip.re"
    #
    #         pyperclip.copy(path_aux)
    #         pyautogui.hotkey('ctrl', 'v')
    #         time.sleep(1)
    #         pyautogui.hotkey('alt', 'a')
    #         time.sleep(1)
    #         pyautogui.press('enter')
    #         # pyautogui.press('left')
    #         pyautogui.press('enter')
    #
    #         time.sleep(2)


            # Salvar arquivos
            # Analitico GRF
    #         pyautogui.hotkey('alt','r')
    #         pyautogui.press('down')
    #         pyautogui.press('right')
    #         pyautogui.press('enter')
    #
    #         time.sleep(1)
    #         _clicar_pela_imagem("imgs\gerar.png")
    #         pyautogui.press('enter')
    #         _clicar_pela_imagem("imgs\ok_4.png")
    #         pyautogui.press('enter')
    #         time.sleep(1)
    #
    #         origem = _pegar_ultimo_arquivo_modificado(pasta_caixa_path)
    #         destino = str(df.at[i, "Caminhos"])
    #         destino += "\\" + ano + "\\" + mes + "." + ano
    #         try:
    #             shutil.move(origem, destino)
    #         except Error:
    #             pass
    #
    #         pyautogui.hotkey('alt', 'f')
    #
    #         # Comprovante de Declaração à Previdência
    #         time.sleep(2)
    #         _clicar_pela_imagem("imgs/comprovante_declaracao.png")
    #         pyautogui.hotkey('alt', 'g')
    #         pyautogui.press('enter')
    #         time.sleep(2)
    #         pyautogui.press('enter')
    #         pyautogui.hotkey('alt', 'f')
    #         pyautogui.press('enter')
    #         time.sleep(1)
    #
    #         origem = _pegar_ultimo_arquivo_modificado(pasta_caixa_path)
    #         try:
    #             shutil.move(origem, destino)
    #         except Error:
    #             pass
    #
    #         time.sleep(1)
    #
    #         # GPS
    #         pyautogui.press('esc', presses=3)
    #         pyautogui.hotkey('alt','r')
    #         pyautogui.press('down')
    #         pyautogui.press('right')
    #
    #         _clicar_pela_imagem("imgs/gps.png")
    #         pyautogui.hotkey('alt', 'p')
    #         time.sleep(1)
    #
    #         _clicar_pela_imagem("imgs\ok_3.png")
    #         time.sleep(1)
    #         pyautogui.press('enter')
    #         time.sleep(1)
    #
    #         origem = _pegar_ultimo_arquivo_modificado(pasta_caixa_path)
    #         try:
    #             shutil.move(origem, destino)
    #         except Error:
    #             pass
    #
    #         pyautogui.hotkey('alt', 'f')
    #         time.sleep(1)
    #
    #         # Analítico
    #         pyautogui.press('down', presses=3)
    #         pyautogui.press('enter')
    #         pyautogui.hotkey('alt', 'g')
    #         pyautogui.click(pyautogui.locateCenterOnScreen("imgs\ok_3.png", confidence=confidence))
    #         time.sleep(1)
    #         pyautogui.press('enter')
    #         time.sleep(1)
    #
    #         origem = _pegar_ultimo_arquivo_modificado(pasta_caixa_path)
    #         try:
    #             shutil.move(origem, destino)
    #         except Error:
    #             pass
    #
    #         pyautogui.hotkey('alt', 'f')
    #
    #         pyautogui.press('esc', presses=2)
    #
    #         # RE
    #         pyautogui.press('esc', presses=3)
    #         pyautogui.hotkey('alt','r')
    #         pyautogui.press('down')
    #         pyautogui.press('right')
    #
    #         _clicar_pela_imagem("imgs/re.png")
    #         pyautogui.hotkey('alt','g')
    #         pyautogui.press('enter')
    #
    #         _clicar_pela_imagem("imgs\ok_4.png")
    #         time.sleep(1)
    #
    #         origem = _pegar_ultimo_arquivo_modificado(pasta_caixa_path)
    #         try:
    #             shutil.move(origem, destino)
    #         except Error:
    #             pass
    #
    #
    #         pyautogui.hotkey('alt','f')
    #         time.sleep(1)
    #         pyautogui.press('esc', presses=3)
    # pyautogui.hotkey('alt', 'tab')
    # app.infoBox("Fim", "Os relatórios foram salvos com sucesso.")


# _salvar_relatorios([854, 862, 863], "08", "2020", 'hehe')