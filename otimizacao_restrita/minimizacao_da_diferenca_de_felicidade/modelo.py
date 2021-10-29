'''

Programação Linear e Otimização - Minimização da Diferença de Felicidade (Modelo)
Felipe Daniel Dias dos Santos - 11711ECP004
Graduação em Engenharia de Computação - Faculdade de Engenharia Elétrica - Universidade Federal de Uberlândia

'''

#Importando as bibliotecas necessárias
import gurobipy as gp
import xlsxwriter

#Imprimindo solução e salvando em um arquivo
def imprimirSalvarSolucao(modelo, relacao_de_afinidade, afinidades, alocar):

    if modelo.status == gp.GRB.OPTIMAL:

        print("\nSolução existente para essa instância do problema. Imprimindo resultados:")

        #Criando o arquivo de nome "Resultados", que conterá os dados resultantes
        resultados = xlsxwriter.Workbook("saida.xlsx")
        planilha_alocacao = resultados.add_worksheet("Alocação")
        planilha_alocacao.write(0, 0, "Professor")
        planilha_alocacao.write(0, 1, "Disciplina")
        planilha_alocacao.write(0, 2, "Afinidade")

        linha = 1
        custo_total = 0
        
        print("\nValor mínima da diferença entre afinidade máxima e mínima:", modelo.objVal)
        print("\nAlocação para atingir o valor ótimo:\n")

        for relacao in relacao_de_afinidade:

            #Para cada relacao de afinidade alocada, imprimi-la e armazená-la na planilha, 
            #assim como sua afinidade
            if alocar[relacao].x:

                print("Alocação:", relacao, "Afinidade:", afinidades[relacao])
       
                planilha_alocacao.write(linha, 0, relacao[0])
                planilha_alocacao.write(linha, 1, relacao[1])
                planilha_alocacao.write(linha, 2, afinidades[relacao])
                linha += 1
        
        planilha_alocacao.write(linha + 1, 0, "Diferença entre afinidade máxima e mínima: ")
        planilha_alocacao.write(linha + 1, 1, modelo.objVal)
        resultados.close()

        print("\nSolução armazenada com sucesso no arquivo Resultados.xlsx\n")

    else:

        print("Não existe solução para essa instância do problema. Insira dados de entrada distintos.")

#Função de resolução do problema de programação linear de alocação de disciplinas a professores
def resolver(professores, qnt_professores, disciplinas, qnt_disciplinas, relacao_de_afinidade, afinidades):
  
    #Definição do modelo
    modelo = gp.Model("Alocação de Disciplinas")

    #Definição da variável de decisão no modelo (a alocação de uma disciplina para um 
    #professor)
    alocar = modelo.addVars(relacao_de_afinidade, ub = 1, vtype = gp.GRB.BINARY, name = 'x')

    #Definição da variável de soma de afinidades de cada professor
    soma_afinidades = modelo.addVars(professores)

    #Definição da variável de afinidade mínima
    afinidade_minima = modelo.addVar(name = 'y')

    #Definição da variável de afinidade máxima
    afinidade_maxima = modelo.addVar(name = 'z')

    #Restrição: cada disciplina e professor são alocados tantas vezes quanto requisitado
    #Nesse caso, deve existir um professor em cada disciplina e cada professor leciona
    #exatamente três disciplinas (ou conforme as informações disponíveis nas planilhas)
    modelo.addConstrs((alocar.sum('*', disc) == qnt_disciplinas[disc] for disc in disciplinas), '_')
    modelo.addConstrs((alocar.sum(prof, '*') == qnt_professores[prof] for prof in professores), '_')

    #Restrição: a soma de afinidades de cada professor é a soma do produto entre a 
    #alocação e a afinidade de cada disciplina
    modelo.addConstrs((soma_afinidades[prof] == (gp.quicksum(alocar[(prof, disc)] * afinidades[(prof, disc)] for disc in disciplinas)) for prof in professores), '_')

    #Restrição: a afinidade mínima é menor ou igual a todas as somas de afinidades
    modelo.addConstrs((afinidade_minima <= soma_afinidades[prof] for prof in professores), '_')

    #Restrição: a afinidade máxima é maior ou igual a todas as somas de afinidades
    modelo.addConstrs((afinidade_maxima >= soma_afinidades[prof] for prof in professores), '_')
    
    #Função objetivo: deseja-se minimizar a diferença entre a afinidade máxima e a afinidade
    #mínima, de forma que todos os professores tenham aproximadamente o mesmo "grau de felicidade" 
    modelo.setObjective(afinidade_maxima - afinidade_minima, gp.GRB.MINIMIZE)

    #Resolvendo e imprimindo o modelo
    modelo.optimize()
    imprimirSalvarSolucao(modelo, relacao_de_afinidade, afinidades, alocar)
