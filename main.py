# ==================================================================

# Essa função formata o arquivo txt que está sem padrão de colunas

# e o transforma em um padrão de 8 colunas e salva. Isso serve para

# que o pandas possa ler o arquivo como uma tabela. Lembrando que

# essa função recebe o nome e o caminho onde o arquivo txt está.

# ==================================================================

def formata_txt(nome_txt):
    linhas = []

    linha_splitada = []

    novas_linhas = []

    # Lê o arquivo e armazena cada linha do arquivo na lista linhas:

    # [H5X411       5X411       16MAR23MAR      0004000     00076F      EZE0915     1050BOGMIA  FF]

    arquivo = open(nome_txt, 'r')

    for linha in arquivo:
        linhas.append(linha)

    arquivo.close()

    # Apaga as 5 primeiras linhas e a última

    linhas = linhas[5:-1].copy()

    # Divide cada linha em uma lista de palavras:

    # ['H5X411', '5X411', '16MAR23MAR', '0004000', '00076F', 'EZE0915', '1050BOGMIA', 'FF']

    for linha in linhas:
        linha_splitada.append(linha.split())

    linhas = []

    # Dependendo da quantidade de palavras (colunas) inserir colunas faltantes

    for linha in linha_splitada:

        # Linha com 5 colunas

        if len(linha) == 5:
            linha.insert(1, '-')

            linha.insert(3, '-')

            linha.insert(6, '-')

            linha.append('\n')

            # Linha com 6 colunas

        if len(linha) == 6:

            if len(linha[1]) == 10:

                linha.insert(1, '-')

                linha.insert(6, '-')

            else:

                linha.insert(3, '-')

                linha.insert(5, '-')

            linha.append('\n')

        # Linha com 7 colunas

        if len(linha) == 7:

            if len(linha[2]) == 10:

                linha.insert(5, '-')

            else:

                linha.insert(3, '-')

            linha.append('\n')

        # Linha com 8 colunas

        if len(linha) == 8:
            linha.append('\n')

    # Volta a lista de palavras para uma linha completa com 8 colunas

    for linha in linha_splitada:
        linhas.append(' '.join(linha))

    # Volta a lista de palavras para uma linha completa com 8 colunas

    for linha in linhas:
        novas_linhas.append(''.join(linha))

    # Salva o arquivo com as novas linhas já configuradas

    arquivo_final = open(nome_txt, 'w')

    arquivo_final.truncate()

    for linha in novas_linhas:
        arquivo_final.write(linha)

    arquivo_final.close()


# fim da função ===================================================