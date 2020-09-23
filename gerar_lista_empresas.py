import pandas as pd
import os

df = pd.read_csv("empresas_gfip.csv", encoding="ISO-8859-1", sep=";")

empresas = df["Codigos Empresas"].tolist()
print(empresas)

path = 'P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Pessoal\Relatórios Sefip'


for x in os.listdir(path):
    for i in range(0, len(empresas)):

        if "-" in str(x) and str(empresas[i]) == str(x)[:str(x).find("-")]:
            nome = str(x)[str(x).find("-") + 1:]
            aux = path + "\\" + str(empresas[i]) + "-" + nome
            print (nome)
            try:
                df.at[i, "Nomes Empresas"] = str(str(x)[str(x).find("-") + 1:])
                df.at[i, "Caminhos"] = str(aux)
            except:
                pass

            df.to_csv('empresas_gfip_completo.csv', encoding="ISO-8859-1", sep=";", index=False)