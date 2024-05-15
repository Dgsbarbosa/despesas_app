import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
import math
import pandas as pd
import datetime
import calendar
import copy


meses = {1: 'janeiro', 2: 'fevereiro', 3: 'março', 4: 'abril', 5: 'maio', 6: 'junho', 7: 'julho', 8: 'agosto', 9: 'setembro', 10: 'outubro', 11: 'novembro', 12: 'dezembro'}



def ler_excel():
    with pd.ExcelFile("Contas Basicas.xlsx") as xlsx:
        
        
        df_receitas = pd.read_excel(xlsx, "Receitas")
        df_despesas = pd.read_excel(xlsx,"Despesas")

  
    
    return df_receitas, df_despesas

   
def organiza_dados(df):     
    
    
    dados = []
    
    for _, conta in df.iterrows():
        
        if pd.isna(conta["vencimento"]):
            conta["vencimento"] = 0
        dado = {"nome": conta['nome'],"vencimento":conta["vencimento"],"valor":conta["valor"],"tipo":conta["tipo"],"parcelas":conta["parcelas"], "primeira parcela":conta["primeira parcela"]}
        dados.append(dado)
        

    return dados

def calendario_do_ano():
    
    # global meses
    ano = datetime.datetime.now().year
    
    global meses
    dias_da_semana = ['segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira', 'sabado', 'domingo']
    
    calendario = calendar.Calendar(firstweekday=6)
    
   
    

    dias_do_ano = {}
    
    for mes in range(1,13):
        
        dias = []
        numero_semana = 1
        
        for dia in calendario.itermonthdates(ano, mes):
            
            print(dia)
            
            if dia.year == ano and dia.month == mes:    
                   
               
                
                dia_da_semana = dias_da_semana[dia.weekday()]
                
                # Obtenha o número da semana dentro do mês
                
                
               
                data_formatada = {"semana":dia_da_semana, "dia": dia.day ,"mes":meses[mes]}
                
                
                # print(numero_semana)
                # print(data_formatada)
                # print()

                # print(numero_da_semana)
                # print(data_formatada)
            
                dias.append(data_formatada)
               
             
         
        dias_do_ano[meses[mes]] = dias
    
    
    
    return dias_do_ano

def verifica_parcela(receitas, despesas):
    
    global meses        
    
    for receita in receitas:
        
               
        if math.isnan(receita["parcelas"]) == False:
            
            qtd_parcelas = receita["parcelas"]
            mes_primeira_parcela = receita["primeira parcela"]

            
            for chave_mes, mes_do_ano in meses.items():
                                
                if mes_do_ano == mes_primeira_parcela:
                    
                    ultima_parcela = int(chave_mes + qtd_parcelas - 1)
                    
                    parcelas_excedentes = "" 
                    if ultima_parcela > 12:
                        
                        parcelas_excedentes = ultima_parcela - 12
                        ultima_parcela = 12
                        
                    meses_parcela = []
                    
                    for mes in range(chave_mes,ultima_parcela + 1):
                        
                        meses_parcela.append(meses[mes])
                    
                    
                    if parcelas_excedentes:
                        meses_parcela.append(f" + {parcelas_excedentes}")
                   
                    receita["meses"] = meses_parcela
                 
            
        else:
            receita["parcelas"] = False
            if pd.isna(receita["primeira parcela"]) :
                receita["primeira parcela"] = False
            
            continue
        
 
    for despesa in despesas:
        
        if math.isnan(despesa["parcelas"]) == False:
            
            qtd_parcelas = despesa["parcelas"]
            mes_primeira_parcela = despesa["primeira parcela"]
            
            for chave_mes, mes_do_ano in meses.items():
                
                if mes_do_ano == mes_primeira_parcela:
                                  
                    chave_mes = int(chave_mes)
                    ultima_parcela = int(chave_mes + qtd_parcelas - 1)
                    
                    parcelas_excedentes = "" 
                    if ultima_parcela > 12:
                        
                        parcelas_excedentes = ultima_parcela - 12
                        ultima_parcela = 12        
                        
                    meses_parcela = []
                    
                    
                    for mes in range(chave_mes,ultima_parcela + 1):
                        
                        meses_parcela.append(meses[mes])
                    
                    if parcelas_excedentes:
                        meses_parcela.append(f" + {parcelas_excedentes}")
                    despesa["meses"] = meses_parcela
            
        else:
            continue
    
    return receitas, despesas
          

def verifica_vencimentos(mes):
    
    

    contas = verifica_parcela()
    dias_do_ano = calendario_do_ano()
    
    dias_do_mes = dias_do_ano[mes] 
    
    receitas = contas["receitas"]
    despesas = contas["despesas"]


       
def linhas_planilha(receitas, despesas):
    
    calendario = calendario_do_ano()
    
    receitas_copy = copy.deepcopy(receitas)
    despesas_copy = copy.deepcopy(despesas)
    

    linhas = {}
    
    for mes in calendario.keys():
        linhas[mes] = ""
        
        for dia in calendario[mes]:
            
            dia_mes = dia["dia"]
            
            for receita in receitas_copy:
                dia_vencimento = int(receita["vencimento"])
                
                if dia_mes == dia_vencimento:
                    # print(dia, receita)
                    ...
            
        break   
            
            


# verifica se o novo vencimento não tem na lista de contas
def valida_vencimento(vencimento, lista_contas):
    
    for conta in lista_contas:
        if int(conta["vencimento"]) == int(vencimento):
            return False
    return True
                    
    
def elimina_datas_iguais(contas,nome_da_conta):


    for conta in contas:
        
        vencimento = conta["vencimento"]
        
        vencimentos_iguais = [d for d in contas if d["vencimento"] == vencimento and d["vencimento"] != 0]
        
        if len(vencimentos_iguais) > 1:
                        
            print(f"\nForam encontrados {len(vencimentos_iguais)} {nome_da_conta} com vencimento no dia: {vencimento}\n")
            print("\nNecessário trocar as datas...")
            for i, vencimento_igual in enumerate(vencimentos_iguais):
                
                index = i + 1 
                nome = vencimento_igual['nome']
                print(f"{index}- {nome.capitalize()}")
            
            while True:
                escolha = input(f"\nQual você deseja alterar? (escolha entre 1 e {len(vencimentos_iguais)}): ")
                
                if escolha.isdigit() and 0 < int(escolha) <= len(vencimentos_iguais) :
                    break
                else:
                    print("\nEscolha um valor válido") 
                    
                        
            item_escolhido = vencimentos_iguais[int(escolha) - 1]
            
            print(f"\nConta escolhida:\n{escolha}- {item_escolhido['nome'].capitalize()}")
            
            
            check = False
            while check == False:
                
                novo_vencimento = input(f"\nQual o novo vencimento: ")
                resposta = valida_vencimento(novo_vencimento,contas)
                
                if novo_vencimento.isdigit() and 0 < int(novo_vencimento) < 31:
                    
                    if resposta == True:
                        for el in contas:
                            if el["nome"] == item_escolhido["nome"]:
                                el["vencimento"] = novo_vencimento 
                                check= True
                              
                    else:
                        print(f"Ja há um vencimento no dia {novo_vencimento} na lista")    
                    
               
                else:
                    print("\nDigite um numero de 1 a 31")
            
                
     
def main():
    
    # le o excel
    df_receitas, df_despesas = ler_excel()
    
    # organiza os dados
    dados_receitas = organiza_dados(df_receitas)
    dados_despesas = organiza_dados(df_despesas)   
    
    
    # verifica se o valor é parcelado e adiciona os meses da parcela nos dados
    
    receitas, despesas = verifica_parcela(dados_receitas,dados_despesas)
    
    # elimina datas iguais
    # elimina_datas_iguais(receitas,"Receitas")
    # elimina_datas_iguais(despesas,"Despesas")
    
    linhas = linhas_planilha(receitas,despesas)
    
    

main()

# calendario_do_ano()

# verifica_vencimentos("janeiro")    