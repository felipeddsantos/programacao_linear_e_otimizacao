'''

Programação Linear e Otimização - Minimização da Diferença de Felicidade (Programa Principal)
Felipe Daniel Dias dos Santos - 11711ECP004
Graduação em Engenharia de Computação - Faculdade de Engenharia Elétrica - Universidade Federal de Uberlândia

'''

#Importanto as bibliotecas necessárias
from setup import setup
from modelo import resolver
from obtencao_de_dados import obterDados

#Função principal: instalação de pacotes, obtenção de dados e resolução do problema
def main():

  #Instalando os pacotes necessários, caso estejam faltando
  setup()

  #Obtendo os dados necessários, gerados e armazenados nas planilhas
  professores, qnt_professores, disciplinas, qnt_disciplinas, relacao_de_afinidade, afinidades = obterDados()

  #Resolvendo o problema com os dados obtidos
  resolver(professores, qnt_professores, disciplinas, qnt_disciplinas, relacao_de_afinidade, afinidades)

if __name__ == "__main__":

  main()
