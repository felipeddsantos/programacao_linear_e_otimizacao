'''

Programação Linear e Otimização - Minimização da Diferença de Felicidade (Geração de Dados)
Felipe Daniel Dias dos Santos - 11711ECP004
Graduação em Engenharia de Computação - Faculdade de Engenharia Elétrica - Universidade Federal de Uberlândia

'''

#Importando as bibliotecas necessárias
import xlsxwriter
from random import randint

#Função de geração de dados referentes às demandas, ofertas e afinidades, todos aleatórios
def geracaoDeDados(qnt_professores):

    #Cada professor deve ser responsável por exatamente 3 disciplinas
    qnt_disciplinas = 3 * qnt_professores

    #Criando o arquivo de nome "Dados de Entrada", que conterá as planilhas de afinidades
    #oferta de professores e demanda de disciplinas
    dados = xlsxwriter.Workbook("entrada.xlsx")
    planilha_professores = dados.add_worksheet("Professores")
    planilha_disciplinas = dados.add_worksheet("Disciplinas")
    planilha_afinidades = dados.add_worksheet("Afinidades")

    #Nomeando as colunas das planilhas
    planilha_professores.write(0, 0, "Professor")
    planilha_professores.write(0, 1, "Oferta")
    planilha_disciplinas.write(0, 0, "Disciplina")
    planilha_disciplinas.write(0, 1, "Demanda")

    #Definindo o nome dos professores e a quantidade de turmas que cada um pode assumir
    for prof in range(qnt_professores):

        linha = prof + 1
        nome = "Professor " + str(linha)
        planilha_afinidades.write(linha, 0, nome)
        planilha_professores.write(linha, 0, nome)

        #O professor, nesse modelo, assume três turmas
        planilha_professores.write(linha, 1, 3)

    #Definindo o nome das disciplinas e a quantidade de professores demandada
    for disc in range(qnt_disciplinas):

        coluna = disc + 1
        nome = "Disciplina " + str(coluna)
        planilha_afinidades.write(0, coluna, nome)
        planilha_disciplinas.write(coluna, 0, nome)

        #Cada disciplina, nesse modelo, necessita somente de um professor
        planilha_disciplinas.write(coluna, 1, 1)

    #Preenchendo a planilha de afinidades: a matriz de afinidades conterá valores aleatórios
    for prof in range(qnt_professores):

        #Adicionando custos aleatórios à planilha de custos
        for disc in range(qnt_disciplinas):

            #O afinidade de um professor para com uma disciplina varia de 1 a 100,
            #determinado aleatoriamente
            planilha_afinidades.write(prof + 1, disc + 1, randint(1, 100))

    dados.close()

#Chamando a função principal com um argumento: a quantidade de professores, considerando que
#a quantidade de disciplinas é um múltiplo da quantidade de professores
geracaoDeDados(8)
