import os
import sys
import xlwings as xw

def atualizar_celulas():
    def obter_nome_arquivo():
        while True:
            nome_arquivo = input("Digite o nome do arquivo Excel : ") + ".xlsx"
            caminho_arquivo = os.path.join(os.getcwd(), nome_arquivo)
            if os.path.exists(caminho_arquivo):
                print(f"Arquivo '{nome_arquivo}' encontrado.")
                return caminho_arquivo
            else:
                print(f"O arquivo '{nome_arquivo}' não foi encontrado.")
                resp = int(input("Deseja tentar novamente? (1 - Sim, 0 - Não): "))
                if resp == 0:
                    print("Encerrando o programa.")
                    sys.exit()

    caminho_arquivo = obter_nome_arquivo()

    try:
        # Abre o arquivo Excel usando xlwings
        workbook = xw.Book(caminho_arquivo)

        def obter_planilha(workbook):
            while True:
                print("Planilhas disponíveis: ", [sheet.name for sheet in workbook.sheets])
                nome_planilha = input("Digite o nome exato da planilha desejada (ex: Dados): ")
                planilhas = [sheet.name for sheet in workbook.sheets]
                if nome_planilha in planilhas:
                    print(f"Planilha '{nome_planilha}' encontrada.")
                    return nome_planilha
                else:
                    print(f"A planilha '{nome_planilha}' não foi encontrada.")
                    resp = int(input("Deseja tentar novamente? (1 - Sim, 0 - Não): "))
                    if resp == 0:
                        print("Encerrando o programa.")
                        sys.exit()

        nome_planilha = obter_planilha(workbook)
        sheet = workbook.sheets[nome_planilha]

        resp = 1
        while resp == 1:
            print("[1] - Atualizar ativos líquidos")
            print("[2] - Atualizar ativos brutos")
            opcao = int(input("Digite a opção desejada: "))
            match opcao:
                case 1:
                    # Solicite a posição da primeira e última célula
                    primeira_celula_liquida = input("Digite a posição do primeiro ativo líquido (ex: B5): ")
                    ultima_celula_liquida = input("Digite a posição do último ativo líquido (ex: B10): ")

                    # Converta as coordenadas das células
                    coluna_primeira = primeira_celula_liquida[0]
                    linha_primeira = int(primeira_celula_liquida[1:])
                    coluna_ultima = ultima_celula_liquida[0]
                    linha_ultima = int(ultima_celula_liquida[1:])

                    # Verifique se as colunas são iguais
                    if coluna_primeira != coluna_ultima:
                        raise ValueError("As colunas das células inicial e final devem ser iguais.")

                    # Calcule a quantidade de linhas
                    quantidade_linhas = linha_ultima - linha_primeira + 1
                    valores = []

                    # Capture os valores das células no intervalo especificado
                    for i in range(quantidade_linhas):
                        cell_value = sheet.range(f'{coluna_primeira}{linha_primeira + i}').value
                        valores.append(cell_value)

                    primeiro_valor_ativo = input("Digite a posição da célula inicial (ex: B6): ")
                    coluna_ativo = primeiro_valor_ativo[0]
                    linha_inicial = int(primeiro_valor_ativo[1:])

                    # Percorra a lista e peça inputs ao usuário para cada célula
                    for i in range(len(valores)):
                        celula_atual = f'{coluna_ativo}{linha_inicial + i}'
                        valor = input(f"Digite o valor para {valores[i]}: ")
                        sheet.range(celula_atual).value = valor

                case _:
                    print("Opção inválida.")
                    
            resp = int(input("Deseja tentar novamente? (1 - Sim, 0 - Não): "))

        workbook.save()
        print('Informação adicionada com sucesso!')
    except Exception as error:
        print('Erro ao atualizar célula:', error)

# Chama a função para atualizar as células
atualizar_celulas()
