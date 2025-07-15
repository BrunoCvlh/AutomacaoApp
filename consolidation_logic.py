import pandas as pd
import os
from tkinter import filedialog

from tb_admin_consolidado import processar_primeiro_arquivo
from tb_balanco_planos import processar_segundo_arquivo

def consolida_e_salva_excel(file1_path, file2_path):
    df_primeiro = processar_primeiro_arquivo(file1_path)
    df_segundo = processar_segundo_arquivo(file2_path)

    df_final = pd.concat([df_primeiro, df_segundo], ignore_index=True)

    caminho_saida = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os Arquivos", "*.*")],
        initialfile="dados_consolidados_finais.xlsx",
        title="Salvar Arquivo Consolidado Como"
    )

    if not caminho_saida:
        return None

    df_final.to_excel(caminho_saida, index=False)
    
    return caminho_saida
