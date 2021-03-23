import pyperclip
import pywinauto

from appJar import gui
from comecar import start


def converte_mes(mes):
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
    elif str(mes) == "13º":
        return "13"



def ajuda(submenu):
    aux = 'Este sistema é capaz de gerar as GFIPs das empresas passadas como parâmetro. É possível adicionar as ' \
          'empresas carregando uma planilha contendo seus ids na primeira coluna ou digitando uma a uma.'
    if submenu == 'Como usar':
        app.infoBox('Como usar', aux)

    elif submenu == 'Versão':
        app.infoBox('Versão', 'Versão 1.0')


# def thread_comecar():
#     global empresas
#     ano = app.getOptionBox("Ano: ")
#     mes = app.getOptionBox("Mês: ")
#     if mes is not None and ano is not None:
#         mes = converte_mes(mes)
#         app.thread(comecar(empresas, mes, ano, app))
#     else:
#         app.infoBox("Erooou...", "Mês ou Ano não selecionado")

def copy_link():
    pyperclip.copy("https://docs.google.com/spreadsheets/d/1SFa0STOpjhB5b8Eqi92E40yDBCH9EFwqG_Mgv5gYdq0/")


def thread_start():
    global empresas
    ano = app.getOptionBox("Ano: ")
    mes = app.getOptionBox("Mês: ")
    if mes is not None and ano is not None:
        mes = converte_mes(mes)
        app.thread(start(mes, ano, app))
    else:
        app.infoBox("Erooou...", "Mês ou Ano não selecionado")


def start_dominio():
    app = pywinauto.Application(backend='win32').start("C:\\Contabil\\contabil.exe \\folha")
    login = app.window(title_re='.*Conectando*.')
    login['Nome do Usuario:Edit2'].set_text("novus123")
    login['&OKButton'].click()

def start_sefip():
    pywinauto.Application(backend='win32').start("E:\\Users\\thiago.madeira\\C\\SEFIP\\Sefip.exe")

# Variaveis globais
empresas = []

# Criando a interface Gráfica
app = gui("gFIP-FIP")
app.setResizable(canResize=True)
app.setLocation("CENTER")
app.setStretch("both")
app.setSticky("nesw")
app.setFont(10)
app.setSize(500, 200)
app.setIcon("logo.gif")
app.addMenuList("Ajuda", ["Como usar", "Versão"], ajuda)

#######################################################################################################################
coluna = 0
linha = 0

# app.addLabel("spreadsheet_link", "https://docs.google.com/spreadsheets/d/1SFa0STOpjhB5b8Eqi92E40yDBCH9EFwqG_Mgv5gYdq0/",
#              row=linha, column=coluna)
# linha += 1
app.addLabelOptionBox("Mês: ", ["- Mês -", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 'Julho',
                                "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro", "13º"],  row=linha, column=coluna)
linha += 1
app.addLabelOptionBox("Ano: ", ["- Ano -", "2021"],  row=linha, column=coluna)
linha += 1
app.addButton("Começar", thread_start, row=linha, column=coluna)
linha += 1
########################################################################################################################
coluna = 1
linha = 0

app.addButton("Abrir Domínio", start_dominio, row=linha, column=coluna)
linha += 1
app.addButton("Abrir Sefip", start_sefip, row=linha, column=coluna)
linha += 1
app.addButton("Copiar link da planilha", copy_link, row=linha, column=coluna)
linha += 1
########################################################################################################################
coluna = 3
linha = 0
########################################################################################################################
# Inicializa a GUI
app.go()
