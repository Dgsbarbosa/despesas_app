import pandas as pd

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
        
        dado = {"nome": receita['nome'],"vencimento":receita["vencimento"],"valor":receita["valor"],"tipo":receita["tipo"]}
        dados_receitas.append(dado)
        
    
    
    for index, despesa in despesas.iterrows():
        
        dado = {"nome": despesa["nome"], "vencimento": despesa["vencimento"],"valor":despesa["valor"],"tipo":despesa["tipo"],"parcelas":despesa["parcelas"]}  
        dados_despesas.append(dado) 
 


    print("Receitas: ",dados_receitas,"Despesas: ", dados_despesas)

    return dados_receitas, dados_despesas



organiza_dados()