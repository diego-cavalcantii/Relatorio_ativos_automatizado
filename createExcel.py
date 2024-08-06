import openpyxl
import os

def criar_arquivo_excel():
    # Cria uma nova instância de Workbook
    workbook = openpyxl.Workbook()

    # Adiciona uma nova planilha (a planilha padrão criada pelo Workbook é usada aqui)
    sheet1 = workbook.active
    sheet1.title = 'Planilha1'

    # Adiciona valores na célula A2
    sheet1['A2'] = 5

    # Define o caminho do arquivo
    nome_arquivo = 'exemplo.xlsx'
    caminho_arquivo = os.path.join(os.getcwd(), nome_arquivo)

    # Verifica se o arquivo já existe e gera um novo nome
    i = 1
    while os.path.exists(caminho_arquivo):
        nome_arquivo = f'exemplo{i}.xlsx'
        caminho_arquivo = os.path.join(os.getcwd(), nome_arquivo)
        i += 1

    # Salva o arquivo Excel
    try:
        workbook.save(caminho_arquivo)
        print(f'Arquivo Excel com uma planilha criado com sucesso: {nome_arquivo}')
    except Exception as error:
        print('Erro ao criar o arquivo Excel:', error)

# Chama a função para criar e atualizar o arquivo
criar_arquivo_excel()
