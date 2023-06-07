import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def divide_planilha():
    # Solicita o nome e caminho completo do arquivo de entrada
    arquivo_entrada = input("Digite o caminho completo e nome do arquivo de entrada: ")

    # Verifica se o arquivo existe
    if not os.path.isfile(arquivo_entrada):
        print("Arquivo não encontrado.")
        return

    # Solicita a quantidade de divisões desejada
    quantidade_divisoes = int(input("Digite a quantidade de divisões desejada: "))

    # Carrega a planilha Excel usando o pandas
    planilha = pd.read_excel(arquivo_entrada)

    # Verifica se a quantidade de divisões é maior que o número de linhas da planilha
    if quantidade_divisoes >= planilha.shape[0]:
        print("A quantidade de divisões deve ser menor que o número de linhas da planilha.")
        return

    # Calcula o tamanho de cada divisão
    tamanho_divisao = planilha.shape[0] // quantidade_divisoes

    # Cria uma pasta para armazenar os arquivos divididos
    pasta_saida = os.path.splitext(arquivo_entrada)[0] + "_dividido"
    os.makedirs(pasta_saida, exist_ok=True)

    # Divide a planilha em arquivos separados
    for i in range(quantidade_divisoes):
        # Cria um novo arquivo Excel usando a biblioteca openpyxl
        wb = Workbook()
        ws = wb.active

        # Obtém as linhas correspondentes à divisão atual
        inicio = i * tamanho_divisao
        fim = (i + 1) * tamanho_divisao
        divisao = planilha.iloc[inicio:fim]

        # Copia as linhas para a planilha do arquivo Excel
        for linha in dataframe_to_rows(divisao, index=False, header=True):
            ws.append(linha)

        # Salva o arquivo Excel na pasta de saída
        nome_arquivo_saida = f"{pasta_saida}/divisao_{i+1}.xlsx"
        wb.save(nome_arquivo_saida)

    print(f"A planilha foi dividida em {quantidade_divisoes} partes e salva na pasta '{pasta_saida}'.")

# Executa a função
divide_planilha()
