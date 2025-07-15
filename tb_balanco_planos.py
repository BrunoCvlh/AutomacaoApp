import pandas as pd

def processar_segundo_arquivo(caminho_arquivo_entrada):
    try:
        df_bd = pd.read_excel(caminho_arquivo_entrada, sheet_name="BD - Receitas e Despesas")
        df_pp = pd.read_excel(caminho_arquivo_entrada, sheet_name="PP - Receitas e Despesas")

        colunas_selecionadas = [1, 2, 3] 

        linhas_receita_bd = [17, 18, 20, 22]
        linhas_despesa_bd = [37, 45, 47, 52]

        df_receita_bd = df_bd.iloc[linhas_receita_bd, colunas_selecionadas].copy()
        df_receita_bd.columns = ["Conta", "Orçado", "Realizado"]
        df_receita_bd["Descrição"] = "Receita Previdencial - BD"

        df_despesa_bd = df_bd.iloc[linhas_despesa_bd, colunas_selecionadas].copy()
        df_despesa_bd.columns = ["Conta", "Orçado", "Realizado"]
        df_despesa_bd["Descrição"] = "Despesa Previdencial - BD"

        df_bd_consolidado = pd.concat([df_receita_bd, df_despesa_bd], ignore_index=True)

        linhas_receita_pp = [6, 9, 11, 12, 13, 19, 21, 22]
        linhas_despesa_pp = [50, 52, 53, 57, 68, 69]

        df_receita_pp = df_pp.iloc[linhas_receita_pp, colunas_selecionadas].copy()
        df_receita_pp.columns = ["Conta", "Orçado", "Realizado"]
        df_receita_pp["Descrição"] = "Receita Previdencial - PostalPrev"

        df_despesa_pp = df_pp.iloc[linhas_despesa_pp, colunas_selecionadas].copy()
        df_despesa_pp.columns = ["Conta", "Orçado", "Realizado"]
        df_despesa_pp["Descrição"] = "Despesa Previdencial - PostalPrev"

        df_pp_consolidado = pd.concat([df_receita_pp, df_despesa_pp], ignore_index=True)

        dados_consolidados = pd.concat([df_bd_consolidado, df_pp_consolidado], ignore_index=True)
        
        dados_consolidados.dropna(subset=["Conta", "Orçado", "Realizado"], how='all', inplace=True)

        return dados_consolidados

    except FileNotFoundError:
        raise FileNotFoundError(f"Erro: O arquivo de entrada '{caminho_arquivo_entrada}' não foi encontrado.")
    except KeyError as e:
        raise KeyError(f"Erro: Uma das abas necessárias ('BD - Receitas e Despesas' ou 'PP - Receitas e Despesas') não foi encontrada no arquivo '{caminho_arquivo_entrada}'. Detalhes: {e}")
    except Exception as e:
        raise Exception(f"Ocorreu um erro ao processar o segundo arquivo: {e}")

