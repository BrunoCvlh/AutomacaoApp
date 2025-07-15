import pandas as pd

def processar_primeiro_arquivo(caminho_arquivo_entrada):
    try:
        df = pd.read_excel(caminho_arquivo_entrada, sheet_name="4.1 Variação Mês")

        colunas_selecionadas = [1, 2, 3]

        def processar_secao(dataframe, inicio, fim, descricao):
            dados = dataframe.iloc[inicio:fim, colunas_selecionadas].copy()
            dados.columns = ["Conta", "Orçado", "Realizado"]
            
            dados.dropna(subset=["Conta", "Orçado", "Realizado"], how='all', inplace=True)
            
            if not dados.empty:
                dados["Descrição"] = descricao
            return dados

        dados_secao1 = processar_secao(df, 4, 13, "Despesas Administrativas")
        dados_secao2 = processar_secao(df, 24, 38, "Pessoal e Encargos (Detalhado)")
        dados_secao3 = processar_secao(df, 44, 60, "Serviços de Terceiros")

        dados_consolidados = pd.concat([dados_secao1, dados_secao2, dados_secao3], ignore_index=True)
        
        return dados_consolidados

    except FileNotFoundError:
        raise FileNotFoundError(f"Erro: O arquivo de entrada '{caminho_arquivo_entrada}' não foi encontrado.")
    except KeyError:
        raise KeyError(f"Erro: A aba '4.1 Variação Mês' não foi encontrada no arquivo '{caminho_arquivo_entrada}'.")
    except Exception as e:
        raise Exception(f"Ocorreu um erro ao processar o primeiro arquivo: {e}")

