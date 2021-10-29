'''

Programação Linear e Otimização - Minimização do Custo (Configuração)
Felipe Daniel Dias dos Santos - 11711ECP004
Graduação em Engenharia de Computação - Faculdade de Engenharia Elétrica - Universidade Federal de Uberlândia

'''

#Importando as bibliotecas necessárias
import sys
import subprocess
import pkg_resources

#Verificando os pacotes necessários que estão faltando e instalando-os na máquina
def setup():

    requeridos = {"xlrd == 1.2.0", "gurobipy", "xlsxwriter"} 
    instalados = {pkg.key for pkg in pkg_resources.working_set}
    faltando = requeridos - instalados

    if faltando:

        subprocess.check_call([sys.executable, "-m", "pip", "install", *faltando])
