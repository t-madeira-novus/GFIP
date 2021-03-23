import pandas as pd
import time
from datetime import datetime

# Read Google Spreadsheet
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Export para o Google Spreadsheet
from df2gspread import df2gspread as d2g

# Process' steps
from gerar_selos import gerar_selos
from gerar_sfps import gerar_sfps
from transmitir_gfips import transmitir_gfips
from salvar_relatorios import salvar_relatorios
from gerar_fgts import gerar_fgts
from postar_no_site import postar_no_site

# Auxiliar functions
from funcoes import cria_dicionario


def read_google_spreadsheet():
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']

    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        'gfip-294718-351dce3da4ae.json', scope)  # Authentication json file

    gc = gspread.authorize(credentials)

    wks = gc.open("Controle Empresas GFIP").sheet1
    data = wks.get_all_values()
    headers = data.pop(0)

    df = pd.DataFrame(data, columns=headers)

    return df  # Return the Google Spreadsheet in a Pandas' DataFrame


def start(month, year, app):
    # Load the dataframe
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        'gfip-294718-351dce3da4ae.json', scope)  # The credential json file here
    spreadsheet_key = "1SFa0STOpjhB5b8Eqi92E40yDBCH9EFwqG_Mgv5gYdq0"
    df = read_google_spreadsheet()
    companies = df['Empresas'].tolist()
    faps = df['Fap'].tolist()
    prolabores = df['Pró-labore'].tolist()
    path = "P:\\documentos\\OneDrive - Novus Contabilidade\\Doc Compartilhado\\Pessoal\\Relatórios Sefip"

    # Standardize the Pró-labore's format
    for i in range(0, len(prolabores)):
        if prolabores[i] != "Sim":
            prolabores[i] = "Não"

    # Make sure the codes are int values (sometimes they are imported as float)
    for i in range(0, len(companies)):
        try:
            companies[i] = int(companies[i])
        except ValueError:
            pass

    dictionary = cria_dicionario(companies, faps, prolabores, path)  # {Código: [Nome, Fap, Pró-labore]}
    rows_to_delete = []
    for i in df.index:
        if df.at[i, 'Fase 1'] != 'Ok':
            print ('Começando Fase 1', companies[i])
            phase_1(df, i, dictionary, companies[i], month, year)
        elif month != '13' and 'Verificado' in df.at[i, 'Verificado Fase 1'] and df.at[i, 'Fase 2'] != 'Ok':
            print ('Começando Fase 2', companies[i])
            if phase_2(df, i, dictionary, companies[i], month, year):
                df.at[i, 'Fase 2'] = 'Ok'
                d2g.upload(df, spreadsheet_key, credentials=credentials,
                           row_names=False)  # row_names=False removes the index 'column'
        elif df.at[i, 'Verificado Fase 1'] == 'Verificado' and  df.at[i, 'Verificado Fase 2'] == 'Verificado':
            rows_to_delete.append(i)
        elif df.at[i, 'Pró-labore'] == "Sim" and df.at[i, 'Verificado Fase 1'] == 'Verificado':
            rows_to_delete.append(i)
        else:
            print ("Conferir linha {} empresa {}".format(i, companies[i]))


def phase_1(df, i, dictionary, company,  month, year):
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        'gfip-294718-351dce3da4ae.json', scope)  # The credential json file here
    spreadsheet_key = "1SFa0STOpjhB5b8Eqi92E40yDBCH9EFwqG_Mgv5gYdq0"

    print ("Começando empresa ", company)
    now = datetime.now()
    start_time = now.strftime("%H:%M:%S")
    df.at[i, 'Início'] = start_time

    # Check if the company is the type prolabore
    prolabore = None
    try:
        if dictionary[company][2] == "Sim":
            prolabore = True
        elif dictionary[company][2] == "Não":
            prolabore = False
    except:
        df.at[i, 'Fase 1'] = 'Ok'
        df.at[i, 'Erros'] = 'Não existe pasta'

        now = datetime.now()
        df.at[i, 'Fim'] = now.strftime("%H:%M:%S")
        #df.at[i, 'Diferença de tempo'] = now.strftime("%H:%M:%S") - start_time
        d2g.upload(df, spreadsheet_key, credentials=credentials,
                   row_names=False)  # row_names=False removes the index 'column'
        return

    if gerar_selos(company, month, year, dictionary, prolabore) is False:
        df.at[i, 'Fase 1'] = 'Ok'
        df.at[i, 'Erros'] = 'Problema ao gerar selo'
        now = datetime.now()
        df.at[i, 'Fim'] = now.strftime("%H:%M:%S")
        #df.at[i, 'Diferença de tempo'] = now.strftime("%H:%M:%S") - start_time
        d2g.upload(df, spreadsheet_key, credentials=credentials,
                   row_names=False)  # row_names=False removes the index 'column'
        return
    if gerar_sfps(company, month, year, dictionary, str(dictionary[company][1]), prolabore) is False:
        df.at[i, 'Fase 1'] = 'Ok'
        df.at[i, 'Erros'] = 'Problema ao passar na SEFIP'
        now = datetime.now()
        df.at[i, 'Fim'] = now.strftime("%H:%M:%S")
        #df.at[i, 'Diferença de tempo'] = now.strftime("%H:%M:%S") - start_time
        d2g.upload(df, spreadsheet_key, credentials=credentials,
                   row_names=False)  # row_names=False removes the index 'column'
        return

    if transmitir_gfips(company, month, year, dictionary) is False:
        df.at[i, 'Fase 1'] = 'Ok'
        df.at[i, 'Erros'] = 'Problema ao passar no site da Conecitividade'
        now = datetime.now()
        df.at[i, 'Fim'] = now.strftime("%H:%M:%S")
       # df.at[i, 'Diferença de tempo'] = now.strftime("%H:%M:%S") - start_time
        d2g.upload(df, spreadsheet_key, credentials=credentials,
                   row_names=False)  # row_names=False removes the index 'column'
        return
    if salvar_relatorios(company, month, year, dictionary, prolabore) is False:
        df.at[i, 'Fase 1'] = 'Ok'
        df.at[i, 'Erros'] = 'Problema algum dos relatórios'
        now = datetime.now()
        df.at[i, 'Fim'] = now.strftime("%H:%M:%S")
        #df.at[i, 'Diferença de tempo'] = now.strftime("%H:%M:%S") - start_time
        d2g.upload(df, spreadsheet_key, credentials=credentials,
                   row_names=False)  # row_names=False removes the index 'column'
        return

    if prolabore is False and month != '13':
        print ("Gerando guia do FGTS...", month)
        if gerar_fgts(company,month, year, dictionary) is False:
            df.at[i, 'Fase 1'] = 'Ok'
            df.at[i, 'Erros'] = 'Problema ao gerar guia do FGTS'
            now = datetime.now()
            df.at[i, 'Fim'] = now.strftime("%H:%M:%S")
            #df.at[i, 'Diferença de tempo'] = now.strftime("%H:%M:%S") - start_time
            d2g.upload(df, spreadsheet_key, credentials=credentials,
                       row_names=False)  # row_names=False removes the index 'column'
            return

    df.at[i, 'Fase 1'] = 'Ok'
    now = datetime.now()
    df.at[i, 'Fim'] = now.strftime("%H:%M:%S")
#    df.at[i, 'Diferença de tempo'] = now.strftime("%H:%M:%S") - start_time
    d2g.upload(df, spreadsheet_key, credentials=credentials,
               row_names=False)  # row_names=False removes the index 'column'
    return

def phase_2(df, i, dictionary, company,  month, year):
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        'gfip-294718-351dce3da4ae.json', scope)  # The credential json file here
    spreadsheet_key = "1SFa0STOpjhB5b8Eqi92E40yDBCH9EFwqG_Mgv5gYdq0"
    if postar_no_site(company, month, year, dictionary) is False:
        df.at[i, 'Erros'] = 'Erro ao postar guia do FGTS'
        d2g.upload(df, spreadsheet_key, credentials=credentials,
                   row_names=False)
        return False
    return True

def postar_guias(df, mes, ano, app):
    empresas = df['Código'].tolist()
    faps = df['Fap'].tolist()
    path = "P:\\documentos\\OneDrive - Novus Contabilidade\\Doc Compartilhado\\Pessoal\\Relatórios Sefip"
    dicionario = cria_dicionario(empresas, faps, path)

    for i in df.index:
        empresa = df.at[i, 'Código']
        if df.at[i, 'Fase 1'] != 'Ok':# and df.at[i, 'Fase 2'] != 'Ok':
            print ("Começando ", empresa)
            if postar_no_site(empresa, mes, ano, dicionario, app) is False:
                continue


# Lendo google planilhas
# scope = ['https://spreadsheets.google.com/feeds',
#          'https://www.googleapis.com/auth/drive']
#
# credentials = ServiceAccountCredentials.from_json_keyfile_name(
#          'gfip-294718-351dce3da4ae.json', scope) # Your json file here
#
# gc = gspread.authorize(credentials)
# df = gc.open("Controle Empresas GFIP").sheet1
# data = wks.get_all_values()
# headers = data.pop(0)
#
# df = pd.DataFrame(data, columns=headers)


# df = pd.read_csv("E:\\Users\\thiago.madeira\\PycharmProjects\\GFIP\\empresas_gfip.csv", encoding="latin1", sep=";",
#                  converters={'Codigos Empresas': lambda x: str(x)})
#
# comecar(df, "10", "2020", "teste", prolabore=True)