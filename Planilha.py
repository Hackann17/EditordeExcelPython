import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import tkinter as tk
from tkinter import messagebox


def main():
    root = tk.Tk()
    root.title('Divisor de Planilhas')

    # label
    tk.Label(root, text='Digite o diretorio exato da planilha que deseja editar:').grid(column=0, row=0)
    tk.Label(root, text='Digite a quantidade de vezes que a planilha deve ser dividida:').grid(column=0, row=2)

    # entry
    entDiretorio = tk.Entry(root, width=50)
    entDiretorio.grid(column=0, row=1, columnspan=19, padx=10, pady=10)

    entQtdDivisoes = tk.Entry(root, width=25)
    entQtdDivisoes.grid(column=0, row=3, padx=10, pady=10)

    def dividir():
        if entDiretorio.get() == '' or entQtdDivisoes.get() == '':
            messagebox.showinfo('Sem entrada', 'Digite o diretorio  e a quantidae de divisões pfvr')
            return

        for i in entQtdDivisoes.get():
            if i.isalpha() or i == '.' or i == ',' or i == ';':
                messagebox.showinfo('Número', 'A quantidade deve ser um numero inteiro pfvr')
                return

        divide_planilha(entDiretorio.get(), int(entQtdDivisoes.get()))

    # Button
    tk.Button(root, text='Dividir', command=dividir).grid(column=0, row=4)

    root.mainloop()


def divide_planilha(arquivo_entrada, quantidade_divisoes):
    # Solicita o nome e caminho completo do arquivo de entrada
    # arquivo_entrada = input("Digite o caminho completo e nome do arquivo de entrada: ")
    if arquivo_entrada == '':
        messagebox.showerror('ERRO', 'Digite um diretório valido pfvr')
        return

        # Verifica se o arquivo existe
    if not os.path.isfile(arquivo_entrada):
        messagebox.showerror('ERRO', 'Arquivo não encontrado.\n Verifique a extenção(tipo) do arquivo e se seu nome'
                                     ' está correto')
        print("Arquivo não encontrado.")
        return

    # Solicita a quantidade de divisões desejada
    # Carrega a planilha Excel usando o pandas
    planilha = pd.read_excel(arquivo_entrada)

    # Verifica se a quantidade de divisões é maior que o número de linhas da planilha
    if quantidade_divisoes >= planilha.shape[0]:
        messagebox.showinfo('Quantidadede', 'A quantidade de divisões deve ser menor que o número de linhas da '
                                            'planilha.')
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
        nome_arquivo_saida = f"{pasta_saida}/divisao_{i + 1}.xlsx"
        wb.save(nome_arquivo_saida)

    messagebox.showinfo('Finalizado!',
                        f"A planilha foi dividida em {quantidade_divisoes} partes e salva na pasta '{pasta_saida}'.")
    print(f"A planilha foi dividida em {quantidade_divisoes} partes e salva na pasta '{pasta_saida}'.")


main()
