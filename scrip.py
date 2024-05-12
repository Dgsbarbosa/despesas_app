import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
import math
import pandas as pd
import datetime
import calendar


meses = {1: 'janeiro', 2: 'fevereiro', 3: 'março', 4: 'abril', 5: 'maio', 6: 'junho', 7: 'julho', 8: 'agosto', 9: 'setembro', 10: 'outubro', 11: 'novembro', 12: 'dezembro'}


def main():
    ...

def ler_excel():
    with pd.ExcelFile("Contas Basicas.xlsx") as xlsx:
        
        
        df_receitas = pd.read_excel(xlsx, "Receitas")
        df_despesas = pd.read_excel(xlsx,"Despesas")

  
    
    return df_receitas, df_despesas

   

def organiza_dados():
        
    receitas, despesas = ler_excel()
    
    dados_receitas = []
    dados_despesas = []
    
    for index, receita in receitas.iterrows():
        
        dado = {"nome": receita['nome'],"vencimento":receita["vencimento"],"valor":receita["valor"],"tipo":receita["tipo"],"parcelas":receita["parcelas"], "primeira parcela":receita["primeira parcela"]}
        dados_receitas.append(dado)
        
    
    
    for index, despesa in despesas.iterrows():
        
        dado = {"nome": despesa["nome"], "vencimento": despesa["vencimento"],"valor":despesa["valor"],"tipo":despesa["tipo"],"parcelas":despesa["parcelas"], "primeira parcela":despesa["primeira parcela"]}  
        dados_despesas.append(dado) 
 

        
    # print("Receitas: ",dados_receitas,"Despesas: ", dados_despesas)

    return {"receitas":dados_receitas, "despesas":dados_despesas}

def calendario_do_ano():
    
    # global meses
    ano = datetime.datetime.now().year
    
    global meses
    # meses = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
    #      'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'] 
    
    dias_da_semana = ['segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira', 'sábado', 'domingo']
    
    calendario = calendar.Calendar()
    
    
    dias_do_ano = {}
    
    for mes in range(1,13):
        
        dias = []
        
        
        for dia in calendario.itermonthdates(ano, mes):
            
            if dia.year == ano and dia.month == mes:    
                   
               
                
                dia_da_semana = dias_da_semana[dia.weekday()]
                
                
                data_formatada = {"semana":dia_da_semana, "dia": dia.day ,"mes":meses[mes]}
                
                
                dias.append(data_formatada)
                
               
                
        dias_do_ano[meses[mes]] = dias
    
    
    
    return dias_do_ano

def verifica_parcela():
    
    global meses
    contas = organiza_dados()
    
    receitas = contas["receitas"]
    despesas = contas["despesas"]

    for receita in receitas:
        
        if math.isnan(receita["parcelas"]) == False:
            
            qtd_parcelas = receita["parcelas"]
            mes_primeira_parcela = receita["primeira parcela"]

            # print(qtd_parcelas)
            # print(mes_primeira_parcela)           
            # print(meses.get(1))
            
            for chave_mes, mes_do_ano in meses.items():
                
                if mes_do_ano == mes_primeira_parcela:
                    
                    ultima_parcela = chave_mes + qtd_parcelas - 1
                    
                    meses_parcela = []
                    
                    for mes in range(chave_mes,ultima_parcela + 1):
                        
                        meses_parcela.append(meses[mes])
                    
                    receita["meses"] = meses_parcela
                    # print(meses_parcela)
                    
            print("receita")
            print(receita)
            
            
            
            
            
        else:
            # print(receita["parcelas"])
            # print("esta vazio")
            continue
        
        # print(receita)
 
 
    for despesa in despesas:
        
        if math.isnan(despesa["parcelas"]) == False:
            
            qtd_parcelas = despesa["parcelas"]
            mes_primeira_parcela = despesa["primeira parcela"]
            
            for chave_mes, mes_do_ano in meses.items():
                
                if mes_do_ano == mes_primeira_parcela:
                                  
                    chave_mes = int(chave_mes)
                    ultima_parcela = int(chave_mes + qtd_parcelas - 1)
                    
                    if ultima_parcela > 12:
                        ultima_parcela = 12
                    
                    meses_parcela = []
                    
                    
                    for mes in range(chave_mes,ultima_parcela + 1):
                        
                        meses_parcela.append(meses[mes])
                    
                    despesa["meses"] = meses_parcela
                    # print(meses_parcela)
                    
            print("despesa")
            print(despesa)
            
            
            
            
            
        else:
            # print(receita["parcelas"])
            # print("esta vazio")
            continue
    # print(contas)
    # print(receitas)
    # print(despesas)
    
    

def verifica_vencimentos(mes):
    
    contas = verifica_parcela()
    dias_do_ano = calendario_do_ano()
    
    dias_do_mes = dias_do_ano[mes] 
    
    receitas = contas["receitas"]
    despesas = contas["despesas"]
    
    
    
    # print(contas)     
    # print(dias_do_ano)
    
    # print(dias_do_mes)
    # print(receitas)
    
    # print(despesas)
    
    # print(calendario)


# calendario_do_ano()
verifica_parcela()
# verifica_vencimentos("janeiro")    