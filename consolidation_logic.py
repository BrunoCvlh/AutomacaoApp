# consolidation_logic.py
import pandas as pd
import os
import datetime


from tb_admin_consolidado import processar_primeiro_arquivo
from tb_balanco_planos import processar_segundo_arquivo

def consolida_e_salva_excel(file1_path, file2_path, output_path, mes_competencia=None, ano_competencia=None):
    df_primeiro = processar_primeiro_arquivo(file1_path)
    df_segundo = processar_segundo_arquivo(file2_path)

    df_final = pd.concat([df_primeiro, df_segundo], ignore_index=True)

    df_final['Hora da Geração'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if mes_competencia is not None and ano_competencia is not None:
        try:
            data_competencia = datetime.date(ano_competencia, mes_competencia, 1)
            df_final['Data Competência'] = data_competencia.strftime("%d/%m/%Y")
        except ValueError:
            df_final['Data Competência'] = None

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
        
        # Verificar se o arquivo de destino existe
        if os.path.exists(destination_sheet_path):
            # Se existe, carregar a planilha e anexar os novos dados
            with pd.ExcelWriter(destination_sheet_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                # Carregar a planilha existente para obter a última linha
                # Se a aba 'Dados Inseridos' não existe, ela será criada automaticamente
                try:
                    df_existente = pd.read_excel(destination_sheet_path, sheet_name='Dados Inseridos')
                    # Anexar os novos dados, ignorando o cabeçalho para as linhas subsequentes
                    df_consolidado.to_excel(writer, sheet_name='Dados Inseridos', index=False, 
                                            startrow=len(df_existente) + 1, header=False)
                except ValueError:
                    # Se a aba 'Dados Inseridos' não existe no arquivo, escrevê-la com cabeçalho
                    df_consolidado.to_excel(writer, sheet_name='Dados Inseridos', index=False, header=True)
        else:
            # Se o arquivo não existe, criar um novo com a aba 'Dados Inseridos'
            with pd.ExcelWriter(destination_sheet_path, mode='w', engine='openpyxl') as writer:
                df_consolidado.to_excel(writer, sheet_name='Dados Inseridos', index=False, header=True)
        
        print(f"Dados anexados à planilha existente: {destination_sheet_path}")
        return True
    except FileNotFoundError:
        print(f"Erro: A planilha de destino '{destination_sheet_path}' não foi encontrada.")
        return False
    except Exception as e:
        print(f"Erro ao incluir dados na base: {e}")
        return False