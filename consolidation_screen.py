import tkinter as tk
from tkinter import filedialog, messagebox
import os
import datetime
from consolidation_logic import consolida_e_salva_excel

class ConsolidationScreen(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        self.frame1 = tk.LabelFrame(self, text="Primeiro Arquivo (Base do Relatório Operacional)", padx=10, pady=10)
        self.frame1.pack(pady=10, padx=10, fill="x")

        self.label1 = tk.Label(self.frame1, text="Nenhum arquivo selecionado.", wraplength=400)
        self.label1.pack(pady=5)

        self.btn_selecionar1 = tk.Button(self.frame1, text="Selecionar Arquivo 1", command=self.selecionar_arquivo1)
        self.btn_selecionar1.pack(pady=5)

        self.frame2 = tk.LabelFrame(self, text="Segundo Arquivo (Receitas e Despesas Previdenciárias)", padx=10, pady=10)
        self.frame2.pack(pady=10, padx=10, fill="x")

        self.label2 = tk.Label(self.frame2, text="Nenhum arquivo selecionado.", wraplength=400)
        self.label2.pack(pady=5)

        self.btn_selecionar2 = tk.Button(self.frame2, text="Selecionar Arquivo 2", command=self.selecionar_arquivo2)
        self.btn_selecionar2.pack(pady=5)

        self.btn_consolidar = tk.Button(self, text="Consolidar Arquivos", command=self.consolidar_arquivos,
                                        bg="#4CAF50", fg="white", font=("Arial", 10, "bold"), relief="raised")
        self.btn_consolidar.pack(pady=20)

        self.btn_ir_para_google_sheet = tk.Button(self, text="Ir para Envio de Dados",
                                                  command=lambda: self.controller.show_frame("GoogleSheetScreen"),
                                                  bg="#007BFF", fg="white", font=("Arial", 10, "bold"), relief="raised")
        self.btn_ir_para_google_sheet.pack(pady=5)

        self.status_label = tk.Label(self, text="", fg="blue", font=("Arial", 9))
        self.status_label.pack(pady=5)

    def selecionar_arquivo1(self):
        file_path = filedialog.askopenfilename(
            title="Selecione o Primeiro Arquivo",
            filetypes=[("Arquivos Excel", "*.xlsx")]
        )
        if file_path:
            self.controller.caminho_arquivo1 = file_path
            self.label1.config(text=os.path.basename(file_path))
            self.status_label.config(text="Primeiro arquivo selecionado.", fg="blue")
        else:
            self.status_label.config(text="Seleção do primeiro arquivo cancelada.", fg="red")

    def selecionar_arquivo2(self):
        file_path = filedialog.askopenfilename(
            title="Selecione o Segundo Arquivo",
            filetypes=[("Arquivos Excel", "*.xlsx")]
        )
        if file_path:
            self.controller.caminho_arquivo2 = file_path
            self.label2.config(text=os.path.basename(file_path))
            self.status_label.config(text="Segundo arquivo selecionado.", fg="blue")
        else:
            self.status_label.config(text="Seleção do segundo arquivo cancelada.", fg="red")

    def consolidar_arquivos(self):
        if not self.controller.caminho_arquivo1:
            messagebox.showwarning("Atenção", "Por favor, selecione o primeiro arquivo.")
            return
        if not self.controller.caminho_arquivo2:
            messagebox.showwarning("Atenção", "Por favor, selecione o segundo arquivo.")
            return

        self.status_label.config(text="Processando arquivos para consolidação...", fg="orange")
        self.controller.update_idletasks()

        try:
            downloads_path = os.path.expanduser('~/Downloads')
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"consolidado_automatico_{timestamp}.xlsx"
            
            full_output_path = os.path.join(downloads_path, output_filename)
            self.controller.caminho_arquivo3 = full_output_path 

            caminho_saida = consolida_e_salva_excel(
                self.controller.caminho_arquivo1,
                self.controller.caminho_arquivo2,
                self.controller.caminho_arquivo3
            )

            if caminho_saida:
                self.status_label.config(text=f"Salvo!", fg="green")
            else:
                self.status_label.config(text="Operação de salvamento cancelada ou falhou.", fg="red")

        except FileNotFoundError as e:
            self.status_label.config(text=f"Erro: Arquivo não encontrado - {e}", fg="red")
            messagebox.showerror("Erro de Arquivo", str(e))
        except KeyError as e:
            self.status_label.config(text=f"Erro: Coluna ou aba não encontrada - {e}", fg="red")
            messagebox.showerror("Erro de Dados", str(e))
        except Exception as e:
            self.status_label.config(text=f"Erro inesperado durante a consolidação: {e}", fg="red")
            messagebox.showerror("Erro", f"Ocorreu um erro inesperado: {str(e)}")