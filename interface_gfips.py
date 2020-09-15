from appJar import gui
from comecar import _comecar
import pandas as pd


def _converte_mes(mes):
    if str(mes) == "Janeiro":
        return "01"
    elif str(mes) == "Fevereiro":
        return "02"
    elif str(mes) == "Março":
        return "03"
    elif str(mes) == "Abril":
        return "04"
    elif str(mes) == "Maio":
        return "05"
    elif str(mes) == "Junho":
        return "06"
    elif str(mes) == "Julho":
        return "07"
    elif str(mes) == "Agosto":
        return "08"
    elif str(mes) == "Setembro":
        return "09"
    elif str(mes) == "Outubro":
        return "10"
    elif str(mes) == "Novembro":
        return "11"
    elif str(mes) == "Dezembro":
        return "12"


def _ajuda(submenu):
    aux = 'Este sistema é capaz de gerar as GFIPs das empresas passadas como parâmetro. É possível adicionar as ' \
          'empresas carregando uma planilha contendo seus ids na primeira coluna ou digitando uma a uma.'
    if submenu == 'Como usar':
        app.infoBox('Como usar', aux)
    elif submenu == 'Versão':
        app.infoBox('Versão', 'Versão 1.0')


def thread_comecar():
    global empresas
    ano = app.getOptionBox("Ano: ")
    mes = app.getOptionBox("Mês: ")
    if mes is not None and ano is not None:
        mes = _converte_mes(mes)
        app.thread(_comecar(empresas, mes, ano, app))
    else:
        app.infoBox("Erooou...", "Mês ou Ano não selecionado")


def thread_gerar_selos():
    global empresas
    ano = app.getOptionBox("Ano: ")
    mes = app.getOptionBox("Mês: ")
    if mes is not None and ano is not None:
        mes = _converte_mes(mes)
        app.thread(_gerar_selos(empresas, mes, ano, app))
    else:
        app.infoBox("Erooou...", "Mês ou Ano não selecionado")


def thread_gerar_sfps():
    global empresas
    ano = app.getOptionBox("Ano: ")
    mes = app.getOptionBox("Mês: ")
    if mes is not None and ano is not None:
        mes = _converte_mes(mes)
        app.thread(_gerar_sfps(empresas, mes, ano, app))
    else:
        app.infoBox("Erooou...", "Mês ou Ano não selecionado")


def thread_salvar_relatorios_gfips():
    global empresas
    ano = app.getOptionBox("Ano: ")
    mes = app.getOptionBox("Mês: ")
    if mes is not None and ano is not None:
        mes = _converte_mes(mes)
        app.thread(_salvar_relatorios(empresas, mes, ano, app))
    else:
        app.infoBox("Erooou...", "Mês ou Ano não selecionado")


def thread_transmitir_gfips():
    global empresas
    ano = app.getOptionBox("Ano: ")
    mes = app.getOptionBox("Mês: ")
    if mes is not None and ano is not None:
        mes = _converte_mes(mes)
        app.thread(_transmitir_gfips(empresas, mes, ano, app))
    else:
        app.infoBox("Erooou...", "Mês ou Ano não selecionado")


def thread_add_empresa():
    app.thread(_add_empresa())


def _add_empresa():
    global empresas
    try:
        empresas.append(int(app.getEntry("empresa")))
        app.updateListBox("lista_empresas", empresas, select=False, callFunction=False)

    except ValueError:
        if app.getEntry("empresa") == "":
            app.infoBox("Erooou...", "Você está tentando inserir uma empresa sem digitar o ID dela antes!")
        else:
            app.infoBox("Erooou...", app.getEntry("empresa") + " não é uma empresa válida.")


def thread_remove_empresa():
    app.thread(_remove_empresa())


def _remove_empresa():
    global empresas
    try:
        aux = int(app.getEntry("empresa"))
        try:
            empresas.remove(aux)
            app.updateListBox("lista_empresas", empresas, select=False, callFunction=False)
        except ValueError:
            app.infoBox("Erooou...", app.getEntry("empresa") + " não está na lista.")

    except ValueError:
        if app.getEntry("empresa") == "":
            app.infoBox("Erooou...", "Você está tentando remover uma empresa sem digitar o ID dela antes!")
        else:
            app.infoBox("Erooou...", app.getEntry("empresa") + " não é uma empresa válida.")


def thread_carregar_empresas():
    app.thread(_carregar_empresas())


def _carregar_empresas():
    global empresas
    path = app.openBox(title=None, dirName=None, fileTypes=None, asFile=False, parent=None, multiple=False, mode='r')

    df = pd.read_csv(path, encoding="latin1", sep=";")
    primeira_coluna = df.columns[0]
    empresas = df[primeira_coluna].tolist()
    app.updateListBox("lista_empresas", df[primeira_coluna], select=False, callFunction=False)


# Variaveis globais
empresas = []

# Criando a interface Gráfica
app = gui("gFIP-FIP")
app.setFont(10)
app.addMenuList("Ajuda", ["Como usar", "Versão"], _ajuda)

#######################################################################################################################
coluna = 0
linha = 0

app.addLabelOptionBox("Mês: ", ["- Mês -", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 'Julho',
                                "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"],  row=linha, column=coluna)
linha += 1
app.addLabelOptionBox("Ano: ", ["- Ano -", "2020", "2021"],  row=linha, column=coluna)
linha += 1
app.addButton("Começar", thread_comecar, row=linha, column=coluna)
linha += 1
# app.addButton("Gerar .SFPs", thread_gerar_sfps, row=linha, column=coluna)
# linha += 1
# app.addButton("Transmitir GFIPs", thread_transmitir_gfips, row=linha, column=coluna)
# linha += 1
# app.addButton("Salvar Relatórios de GFIPs", thread_salvar_relatorios_gfips, row=linha, column=coluna)
# linha += 1

########################################################################################################################
coluna = 1
linha = 0

app.addLabel(str(linha)+"x"+str(coluna), "Lista de empresas", row=linha, column=coluna)
linha += 1
app.addListBox("lista_empresas", "", row=linha, column=coluna, colspan=0, rowspan=5)
app.addListItems("lista_empresas", empresas)
linha += 1

########################################################################################################################
coluna = 3
linha = 0

app.addLabel(str(linha)+"x"+str(coluna), "Digite uma empresa para adicionar ou remover", row=linha, column=coluna)
linha += 1
app.addEntry("empresa", row=linha, column=coluna)
linha += 1
app.addButton("Adicionar empresa", thread_add_empresa, row=linha, column=coluna)
linha += 1
app.addButton("Remover empresa", thread_remove_empresa, row=linha, column=coluna)
linha += 1
app.addButton("Carregar planilha com as empresas", thread_carregar_empresas, row=linha, column=coluna)
linha += 1

########################################################################################################################
# Inicializa a GUI
app.go()
