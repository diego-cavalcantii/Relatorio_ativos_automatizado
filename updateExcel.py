import os
import sys
import xlwings as xw

def main():
    caminho_arquivo = obter_nome_arquivo()
    try:
        # Abre o arquivo Excel usando xlwings
        workbook = xw.Book(caminho_arquivo)
        nome_planilha = obter_planilha(workbook)
        sheet = workbook.sheets[nome_planilha]

        resp = 1
        while resp == 1:
            print("[1] - Atualizar tabela de forma livre")
            print("[2] - Atualizar ativos líquidos")
            print("[3] - Atualizar ativos brutos")
            print("[4] - Sair")
            opcao = int(input("Digite a opção desejada: "))
            match opcao:
                case 1:
                    titulo_p = input("Digite a posição do primeiro título (ex: B5): ")
                    titulo_u = input("Digite a posição do último título (ex: B10): ")
                    valor_p = input("Digite a posição do primeiro valor a ser inserido (ex: B6): ")
                    atualizar_tabela(titulo_p, titulo_u, valor_p, sheet)
                case 2:
                    titulo_u = input("Digite a posição do último título (ex: B10): ")
                    atualizar_tabela('B6', titulo_u, 'AC6', sheet)
                case 3:
                    titulo_u = input("Digite a posição do último título (ex: B10): ")
                    atualizar_tabela('B6', titulo_u, 'AD6', sheet)
                case 4:
                    print("Encerrando o programa.")
                    sys.exit()
                case _:
                    print("Opção inválida.")
                    
            resp = int(input("Deseja tentar novamente? (1 - Sim, 0 - Não): "))

        workbook.save()
        print('Informação adicionada com sucesso!')
    except Exception as error:
        print('Erro ao atualizar célula:', error)


def obter_nome_arquivo():
    while True:
        nome_arquivo = input("Digite o nome do arquivo Excel: ") + ".xlsx"
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

def atualizar_tabela(titulo_p, titulo_u, valor_p, sheet):
    while True:
        try:
            # Converta as coordenadas das células
            col_primeira_conv = titulo_p.lower().rstrip('0123456789') # Ex: 'B'
            row_primeia_conv = int(titulo_p[len(col_primeira_conv):]) # Ex: 5
            col_ultima_conv = titulo_u.lower().rstrip('0123456789') # Ex: 'B'
            row_ultima_conv = int(titulo_u[len(col_ultima_conv):]) # Ex: 10
        except ValueError:
            print("Erro ao converter coordenadas das células. Verifique os valores de entrada.")
            if titulo_p == 0:
                titulo_p = input("Digite novamente a posição do primeiro título (ex: B5): ")
            titulo_u = input("Digite novamente a posição do último título (ex: B10): ")
            continue

        # Verifique se as colunas são iguais
        if col_primeira_conv != col_ultima_conv:
            print("As colunas das células inicial e final devem ser iguais. Tente novamente.")
            if titulo_p == 0:
                    titulo_p = input("Digite novamente a posição do primeiro título (ex: B5): ")
            titulo_u = input("Digite novamente a posição do último título (ex: B10): ")
            continue

        try:
            # Calcule a quantidade de linhas
            quantidade_linhas = row_ultima_conv - row_primeia_conv + 1
            valores = []

            # Capture os valores das células no intervalo especificado
            for i in range(quantidade_linhas):
                cell_value = sheet.range(f'{col_primeira_conv}{row_primeia_conv + i}').value
                valores.append(cell_value)
            break
        except Exception as e:
            raise RuntimeError(f"Erro ao capturar os valores das células: {e}")

    try:
        # Extraia a coluna e a linha da posição da célula inicial
        coluna_ativo = valor_p.rstrip('0123456789') # Ex: 'B'
        linha_inicial = int(valor_p[len(coluna_ativo):]) # Ex: 6

        # Percorra a lista e peça inputs ao usuário para cada célula
        for i in range(len(valores)):
            celula_atual = f'{coluna_ativo}{linha_inicial + i}'  # Ex: 'B6'
            while True:
                valor = input(f"Digite o valor para {valores[i]} (ou pressione Enter para manter o valor existente): ")
                if valor == "":
                    # Se o usuário pressionar Enter, simplesmente sai do loop sem fazer nada
                    break
                elif valor.isnumeric():
                    # Se o valor inserido for numérico, permite a inserção e substitui o valor na célula
                    sheet.range(celula_atual).value = valor
                    break
                else:
                    # Se o valor inserido não for numérico, exibe uma mensagem de erro
                    print("Valor inválido. Por favor, insira um número.")


    except ValueError:
        raise ValueError("Erro ao converter coordenadas das células de valores. Verifique os valores de entrada.")
    except Exception as e:
        raise RuntimeError(f"Erro ao atualizar as células: {e}")


if __name__ == "__main__":
    main()
