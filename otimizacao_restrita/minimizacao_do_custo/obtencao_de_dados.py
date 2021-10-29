'''

Programação Linear e Otimização - Minimização do Custo (Obtenção de Dados)
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

#Extraindo e formatando os dados da planilha de custos
def obterCustos(planilha_custos, professores, disciplinas):

    arestas = []
    custos = {}

    for linha in range(len(professores)):

        for coluna in range(len(disciplinas)):
    
            #Extraindo as tuplas (professor, disciplina) disponíveis e armazenando em um 
            #dicionário cuja chave é a tupla e valor o custo associado
            tupla = (planilha_custos.cell_value(linha + 1, 0), planilha_custos.cell_value(0, coluna + 1))
            custos[tupla] = planilha_custos.cell_value(linha + 1, coluna + 1)   

            #Armazenando as arestas em uma lista
            arestas.append(tupla)

    #Transformando o dicionário em um multidict e a lista em um tuplelist
    arestas = gp.tuplelist(arestas)
    _, custos = gp.multidict(custos)

    return arestas, custos

#Função para extrair todos os dados das planilhas, armazenando-os em estruturas adequadas
#para utilização no solucionador do problema de programação linear
def obterDados():

    #Abrindo o arquivo e obtendo as planilhas
    dados = xlrd.open_workbook("entrada.xlsx")
    planilha_professores = dados.sheet_by_name("Professores")
    planilha_disciplinas = dados.sheet_by_name("Disciplinas")
    planilha_custos = dados.sheet_by_name("Custos")

    #Obtendo os dados de cada planilha
    professores, qnt_professores = obterProfessores(planilha_professores)
    disciplinas, qnt_disciplinas = obterDisciplinas(planilha_disciplinas)
    arestas, custos = obterCustos(planilha_custos, professores, disciplinas)

    #Retornando todo o conteúdo necessário
    return professores, qnt_professores, disciplinas, qnt_disciplinas, arestas, custos
