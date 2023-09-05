import os
import pandas as pd  # Trabalha com tabelas
import glob  # Encontra o nome dos arquivos no caminho especificado e lista arquivos pelo nome
from pyxlsb import open_workbook as open_xlsb  # Abrir aquivos no formato binario
import os.path  # Veifica o caminho de um diretório e arquivos
from datetime import datetime  # Trabalha com variáveis de tempo


def encontrar_arquivo_mais_recente(caminho_pasta):
    caminho_arquivos = os.path.join(
        caminho_pasta, "*nome do arquivo*") # Os asteriscos significam que estou procurando um arquivo que contenha como nome o texto interno
    arquivos = glob.glob(caminho_arquivos)  # Encontrando e listando arquivos
    caminho_arquivo_mais_recente = min(arquivos, key=os.path.getmtime, default=None)
    return caminho_arquivo_mais_recente

nome_da_pasta_dia = ""  # Variáveis Globais
nome_da_pasta_mes = ""
arquivo_phaseout = r"caminho do arquivo base.xlsx"
df_phaseout = pd.read_excel(arquivo_phaseout)
datas_phaseout = df_phaseout["data"].tolist()
print(datas_phaseout)

# Lógica para manipular os dados dentro das tabelas

def manipular_arquivo_excel(caminho_arquivo):
    # Adicionado a base de dados principal ao programa
    global arquivo_phaseout
    df_phaseout_load = pd.read_excel(arquivo_phaseout)
    df_consulta_estoque = pd.read_excel(caminho_arquivo)
    # Filtrando a coluna "Endereço" com os textos específicos
    filtros_phaseout = ["M2.C26", "M2.C27", "M2.C28"]
    filtros_estoque = ["M2.C26", "M2.C27", "M2.C28", "END_LOST_UZ", "END_LOST_UZ_PAI"]
    df_filtrado_phaseout = df_consulta_estoque[df_consulta_estoque["Endereço"].str.contains("|".join(filtros_phaseout),case=False)]
    df_filtrado_estoque = df_consulta_estoque[~df_consulta_estoque["Endereço"].apply(lambda endereço: any(endereço.startswith(texto) for texto in filtros_estoque))]
    # Somando os números da coluna "Qtd Atual"
    qtd_pecas_estoque = df_filtrado_estoque["Qtd Atual"].sum()
    qtd_pecas_phaseout = df_filtrado_phaseout["Qtd Atual"].sum()

    # Removendo itens duplicados na coluna "Item" e contando a quantidade de itens
    qtd_itensphaseout = df_filtrado_phaseout["Item"].nunique()
    qtd_itens_estoque = df_filtrado_estoque["Item"].nunique()
    # Bloco para converter o nome do mês da pasta para numero do mês
    numero_mes = ""

    if nome_da_pasta_mes == "Janeiro":
        numero_mes = "01"
    elif nome_da_pasta_mes == "Fevereiro":
        numero_mes = "02"
    elif nome_da_pasta_mes == "Março":
        numero_mes = "03"
    elif nome_da_pasta_mes == "Abril":
        numero_mes = "04"
    elif nome_da_pasta_mes == "Maio":
        numero_mes = "05"
    elif nome_da_pasta_mes == "Junho":
        numero_mes = "06"
    elif nome_da_pasta_mes == "Julho":
        numero_mes = "07"
    elif nome_da_pasta_mes == "Agosto":
        numero_mes = "08"
    elif nome_da_pasta_mes == "Setembro":
        numero_mes = "09"
    elif nome_da_pasta_mes == "Outubro":
        numero_mes = "10"
    elif nome_da_pasta_mes == "Novembro":
        numero_mes = "11"
    else:
        numero_mes = "12"
    # Fim do bloco

    data = nome_da_pasta_dia + "/" + numero_mes + "/" + "2023"
    # Adicionando dados nas linhas
    nova_linha = pd.DataFrame({ "data": [data],
                                "qt_itens_phaseout": [qtd_itensphaseout],
                                "qt_peças_phaseout": [qtd_pecas_phaseout],
                                "qt_itens_estoque": [qtd_itens_estoque],
                                "qt_peças_estoque": [qtd_pecas_estoque], })

    # Juntando os dados
    novo_df_phaseout = pd.concat([df_phaseout_load, nova_linha])
    # Visualizar progresso
    print(novo_df_phaseout)
    # Salvando os dados na base de dados
    novo_df_phaseout.to_excel(arquivo_phaseout, index=False)

def processar_pasta_mes(pasta_mes):
    caminho_pasta_mes = os.path.join(pasta_raiz, pasta_mes)
    numero_mes = ""
    global datas_phaseout
    if os.path.isdir(caminho_pasta_mes):
        for pasta_dia in os.listdir(caminho_pasta_mes):
            global nome_da_pasta_mes
            nome_da_pasta_mes = os.path.basename(caminho_pasta_mes)
            caminho_pasta_dia = os.path.join(caminho_pasta_mes, pasta_dia)
            if os.path.isdir(caminho_pasta_dia):
                caminho_arquivo_mais_recente = encontrar_arquivo_mais_recente(caminho_pasta_dia)
                global nome_da_pasta_dia
                nome_da_pasta_dia = os.path.basename(caminho_pasta_dia)

                # Bloco para converter o nome do mês da pasta para numero do mês
                if nome_da_pasta_mes == "Janeiro":
                    numero_mes = "01"
                elif nome_da_pasta_mes == "Fevereiro":
                    numero_mes = "02"
                elif nome_da_pasta_mes == "Março":
                    numero_mes = "03"
                elif nome_da_pasta_mes == "Abril":
                    numero_mes = "04"
                elif nome_da_pasta_mes == "Maio":
                    numero_mes = "05"
                elif nome_da_pasta_mes == "Junho":
                    numero_mes = "06"
                elif nome_da_pasta_mes == "Julho":
                    numero_mes = "07"
                elif nome_da_pasta_mes == "Agosto":
                    numero_mes = "08"
                elif nome_da_pasta_mes == "Setembro":
                    numero_mes = "09"
                elif nome_da_pasta_mes == "Outubro":
                    numero_mes = "10"
                elif nome_da_pasta_mes == "Novembro":
                    numero_mes = "11"
                else:
                    numero_mes = "12"
                # Fim do bloco
                data = nome_da_pasta_dia + "/" + numero_mes + "/" + "2023"

            if str(data) in str(datas_phaseout):
                continue
            else:
                if caminho_arquivo_mais_recente:
                    df_resultado = manipular_arquivo_excel(caminho_arquivo_mais_recente)
                    continue

# Definindo a pasta raiz onde estão as pastas dos meses
pasta_raiz = r"Pasta onde estão todos todas as pastas de meses e dentro dessas pastas as pastas de dias"
# Obtendo a lista de pastas dos meses na pasta raiz
pasta_meses = os.listdir(pasta_raiz)

for pasta_mes in pasta_meses:
    processar_pasta_mes(pasta_mes)

print(f"Atualização finalizada com sucesso")
