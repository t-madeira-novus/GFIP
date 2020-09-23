import os
from gerar_selos import _gerar_selos
from gerar_sfps import _gerar_sfps
from transmitir_gfips import _transmitir_gfips
from salvar_relatorios import _salvar_relatorios
from gerar_fgts import _gerar_fgts
from postar_no_site import _postar_no_site

def _comecar (empresas, mes, ano, app):
    dictionary = {}
    path = "P:/documentos/OneDrive - Novus Contabilidade/Doc Compartilhado/Pessoal/Relatórios Sefip"
    for x in os.listdir(path):
        for i in range (0, len(empresas)):
            if "-" in str(x) and str(empresas[i]) == str(x)[:str(x).find("-")]:
                dictionary[empresas[i]] = str(x)[str(x).find("-")+1:]

    print (dictionary)

    for empresa in empresas:
        if _gerar_selos(empresa, mes, ano, dictionary, app) is False:
            continue
        if _gerar_sfps(empresa, mes, ano, dictionary, app) is False:
            continue
        if _transmitir_gfips(empresa, mes, ano, dictionary, app) is False:
            continue
        if _salvar_relatorios(empresa, mes, ano, dictionary, app) is False:
            continue
        # if _gerar_fgts(empresa, mes, ano, dictionary, app) is False:
        #     continue
        # if _postar_no_site(empresa, mes, ano, dictionary, app) is False:
        #     continue

import pandas as pd
df = pd.read_csv("E:\\Users\\thiago.madeira\\PycharmProjects\\GFIP\\empresas_gfip.csv", encoding="latin1", sep=";")
primeira_coluna = df.columns[0]
empresas = df[primeira_coluna].tolist()
_comecar(empresas, "09", "2020", "teste")