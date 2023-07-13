def formata_txt(nome_txt):
    # Abre o arquivo no modo de leitura ('r')
    with open(nome_txt, 'r') as arquivo:
        # Lê todas as linhas do arquivo e as armazena na lista 'linhas'
        linhas = arquivo.readlines()

    # Remove as 5 primeiras linhas e a última linha da lista 'linhas'
    linhas = linhas[5:-1]

    # Itera sobre cada linha na lista 'linhas'
    for i, linha in enumerate(linhas):
        # Divide a linha em uma lista de palavras utilizando espaço como separador
        linha_splitada = linha.split()

        # Verifica o número de palavras na linha para determinar a formatação
        if len(linha_splitada) < 8:
            # Se a linha tiver menos de 8 palavras, insere '-' nas posições corretas
            if len(linha_splitada) == 5:
                linha_splitada.insert(1, '-')
                linha_splitada.insert(3, '-')
                linha_splitada.insert(6, '-')
            elif len(linha_splitada) == 6:
                if len(linha_splitada[1]) == 10:
                    linha_splitada.insert(1, '-')
                    linha_splitada.insert(6, '-')
                else:
                    linha_splitada.insert(3, '-')
                    linha_splitada.insert(5, '-')
            elif len(linha_splitada) == 7:
                if len(linha_splitada[2]) == 10:
                    linha_splitada.insert(5, '-')
                else:
                    linha_splitada.insert(3, '-')

        # Atualiza a linha na lista 'linhas' com a formatação correta
        linhas[i] = ' '.join(linha_splitada) + '\n'

    # Abre o arquivo no modo de escrita ('w')
    with open(nome_txt, 'w') as arquivo_final:
        # Escreve as linhas formatadas no arquivo
        arquivo_final.writelines(linhas)
