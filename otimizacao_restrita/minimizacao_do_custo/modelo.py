'''

Programação Linear e Otimização - Minimização do Custo (Modelo)
Felipe Daniel Dias dos Santos - 11711ECP004
Graduação em Engenharia de Computação - Faculdade de Engenharia Elétrica - Universidade Federal de Uberlândia

'''

#Importando as bibliotecas necessárias
import gurobipy as gp
import xlsxwriter

#Imprimindo solução e salvando em um arquivo
def imprimirSalvarSolucao(modelo, arestas, custos, fluxo_na_aresta):

    if modelo.status == gp.GRB.OPTIMAL:

        print("\nSolução existente para essa instância do problema. Imprimindo resultados:")

        #Criando o arquivo de nome "Resultados", que conterá os dados resultantes
        resultados = xlsxwriter.Workbook("saida.xlsx")
        planilha_alocacao = resultados.add_worksheet("Alocação")
        planilha_alocacao.write(0, 0, "Professor")
        planilha_alocacao.write(0, 1, "Disciplina")
        planilha_alocacao.write(0, 2, "Custo")

        linha = 1
        custo_total = 0
        
        print("\nCusto mínimo:", modelo.objVal)
        print("\nAlocação para atingir o custo mínimo:\n")

        for aresta in arestas:

            #Para cada relacao de afinidade alocada, imprimi-la e armazená-la na planilha, 
            #assim como sua afinidade
            if fluxo_na_aresta[aresta].x > 0.0001:

                print("Alocação:", aresta, "Custo:", custos[aresta])
       
                planilha_alocacao.write(linha, 0, aresta[0])
                planilha_alocacao.write(linha, 1, aresta[1])
                planilha_alocacao.write(linha, 2, custos[aresta])
                linha += 1
                
                custo_total += custos[aresta]
        
        planilha_alocacao.write(linha, 2, custo_total)
        resultados.close()

        print("\nSolução armazenada com sucesso no arquivo Resultados.xlsx\n")

    else:

        print("Não existe solução para essa instância do problema. Insira dados de entrada distintos.")

#Função de resolução do problema de programação linear de alocação de disciplinas a professores
def resolver(professores, qnt_professores, disciplinas, qnt_disciplinas, arestas, custos):
  
    #Definição do modelo
    modelo = gp.Model("Alocação de Disciplinas")

    #Definição da variável de decisão no modelo (o fluxo em cada aresta)
    fluxo_na_aresta = modelo.addVars(arestas, ub = 1, name = 'x')


    #Restrição: cada disciplina e professor são alocados tantas vezes quanto requisitado
    #Nesse caso, deve existir um professor em cada disciplina e cada professor leciona
    #exatamente uma disciplina (ou conforme as informações disponíveis nas planilhas)
    modelo.addConstrs((fluxo_na_aresta.sum('*', disc) == qnt_disciplinas[disc] for disc in disciplinas), '_')
    modelo.addConstrs((fluxo_na_aresta.sum(prof, '*') == qnt_professores[prof] for prof in professores), '_')
 
    #Função objetivo: minimizar os custos das arestas selecionadas 
    modelo.setObjective(gp.quicksum(custos[(prof, disc)] * fluxo_na_aresta[prof, disc] for prof, disc in arestas), gp.GRB.MINIMIZE)

    #Resolvendo e imprimindo o modelo
    modelo.optimize()
    imprimirSalvarSolucao(modelo, arestas, custos, fluxo_na_aresta)
