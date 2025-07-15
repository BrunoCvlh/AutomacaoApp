import pandas as pd
import os


from tb_admin_consolidado import processar_primeiro_arquivo
from tb_balanco_planos import processar_segundo_arquivo

def consolida_e_salva_excel(file1_path, file2_path, output_path):
    df_primeiro = processar_primeiro_arquivo(file1_path)
    df_segundo = processar_segundo_arquivo(file2_path)

    df_final = pd.concat([df_primeiro, df_segundo], ignore_index=True)

    df_final.to_excel(output_path, index=False)
    
    return output_path

def inclui_dados_na_base(consolidated_file_path, destination_sheet_path):
    """
    Função para incluir dados de um arquivo consolidado em uma planilha de destino local.
    Implemente a lógica de leitura e escrita aqui.
    Retorna True em caso de sucesso, False ou levanta exceção em caso de falha.
    """
    try:
        df_consolidado = pd.read_excel(consolidated_file_path)
        
        with pd.ExcelWriter(destination_sheet_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
            df_consolidado.to_excel(writer, sheet_name='Dados Inseridos', index=False)
        
        print(f"Dados anexados à planilha existente: {destination_sheet_path}")
        return True
    except FileNotFoundError:
        print(f"Erro: A planilha de destino '{destination_sheet_path}' não foi encontrada.")
        return False
    except Exception as e:
        print(f"Erro ao incluir dados na base: {e}")
        return False