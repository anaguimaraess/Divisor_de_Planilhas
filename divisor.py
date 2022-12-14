import pandas as pd
import os, glob

file_excel = "Entrada/arquivo.xlsx"

value_max = 4
posicao_num_linha = 0
value_column = 0

 

lendo_planilha = pd.read_excel (file_excel)
lendo_planilha.to_csv ("Entrada/Result.csv", 
                  index = None,
                  header=True)

   
reading_csv = pd.DataFrame(pd.read_csv("Entrada/Result.csv"))
 

lenght_lines = reading_csv.size
 

qntd_linhas_arredondado = lenght_lines // value_max
qntd_linhas_decimal = lenght_lines / value_max

py_files = glob.glob('Saida/*.csv')


for py_file in py_files:

    try:
        os.remove(py_file)

    except OSError as e:
        print(f"Error:{ e.strerror}")

 

def gerando_planilhas(numero_linhas, num_linha):

    for i in range(numero_linhas):
        lista_nova =[]       
        nome_planilha = "conclusao" + str(i) + ".csv"

        try:   

            for j in range(value_max):
                nome = reading_csv.iloc[num_linha, value_column]
                lista_nova += [nome]
                num_linha+=1

        except:
            print("Existe uma planilha com numero de itens menor do que você setou.")

        nova_planilha_split = pd.DataFrame(lista_nova)
        nova_planilha_split.to_csv("Saida/"+nome_planilha, index = None, header=False)
 

if(qntd_linhas_arredondado >= qntd_linhas_decimal):
    gerando_planilhas(qntd_linhas_arredondado, posicao_num_linha)

else:

    nova_qntd_linhas = qntd_linhas_arredondado + 1
    gerando_planilhas(nova_qntd_linhas, posicao_num_linha)
    
py_files = glob.glob('Entrada/*')

for py_file in py_files:

    try:
        os.remove(py_file)
    except OSError as e:
        print(f"Error:{ e.strerror}")
