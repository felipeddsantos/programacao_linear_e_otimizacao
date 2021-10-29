'''

Programação Linear e Otimização - Minimização da Diferença de Felicidade (Obtenção de Dados)
Felipe Daniel Dias dos Santos - 11711ECP004
Graduação em Engenharia de Computação - Faculdade de Engenharia Elétrica - Universidade Federal de Uberlândia

'''

#Importando as bibliotecas necessárias
import os
import xlrd
import gurobipy as gp

#Extraindo e formatando os dados da planilha de professores
def obterProfessores(planilha_professores):

    professores = {}
    linha = 1

    #A planilha é percorrida uma excessão de índice inválido ocorrer
    while True:

        try:
        
            professores[planilha_professores.cell_value(linha, 0)] = planilha_professores.cell_value(linha, 1)
            linha += 1
    
        except IndexError:
    
            break

    #Transformando o dicionário em um multidict
    professores, qnt_professores = gp.multidict(professores)

    return professores, qnt_professores

#Extraindo e formatando os dados da planilha de disciplinas
def obterDisciplinas(planilha_disciplinas):

    disciplinas = {}
    linha = 1

    #A planilha é percorrida uma excessão de índice inválido ocorrer
    while True:

        try:
        
            disciplinas[planilha_disciplinas.cell_value(linha, 0)] = planilha_disciplinas.cell_value(linha, 1)
            linha += 1
    
        except IndexError:
    
            break

    #Transformando o dicionário em um multidict
    disciplinas, qnt_disciplinas = gp.multidict(disciplinas)

    return disciplinas, qnt_disciplinas

#Extraindo e formatando os dados da planilha de afinidades
def obterAfinidades(planilha_afinidades, professores, disciplinas):

    relacao_de_afinidade = []
    afinidades = {}

    for linha in range(len(professores)):

        for coluna in range(len(disciplinas)):
    
            #Extraindo as tuplas (professor, disciplina) disponíveis e armazenando em um 
            #dicionário cuja chave é a tupla e valor a afinidade associada
            tupla = (planilha_afinidades.cell_value(linha + 1, 0), planilha_afinidades.cell_value(0, coluna + 1))
            afinidades[tupla] = planilha_afinidades.cell_value(linha + 1, coluna + 1)   

            #Armazenando as relações de afinidade em uma lista
            relacao_de_afinidade.append(tupla)

    #Transformando o dicionário em um multidict e a lista em um tuplelist
    relacao_de_afinidade = gp.tuplelist(relacao_de_afinidade)
    _, afinidades = gp.multidict(afinidades)

    return relacao_de_afinidade, afinidades

#Função para extrair todos os dados das planilhas, armazenando-os em estruturas adequadas
#para utilização no solucionador do problema de programação linear
def obterDados():

    #Abrindo o arquivo e obtendo as planilhas
    dados = xlrd.open_workbook("entrada.xlsx")
    planilha_professores = dados.sheet_by_name("Professores")
    planilha_disciplinas = dados.sheet_by_name("Disciplinas")
    planilha_afinidades = dados.sheet_by_name("Afinidades")

    #Obtendo os dados de cada planilha
    professores, qnt_professores = obterProfessores(planilha_professores)
    disciplinas, qnt_disciplinas = obterDisciplinas(planilha_disciplinas)
    relacao_de_afinidade, afinidades = obterAfinidades(planilha_afinidades, professores, disciplinas)

    #Retornando todo o conteúdo necessário
    return professores, qnt_professores, disciplinas, qnt_disciplinas, relacao_de_afinidade, afinidades
