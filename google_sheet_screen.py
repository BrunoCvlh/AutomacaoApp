import tkinter as tk
from tkinter import messagebox, filedialog # Import filedialog
import os # Import os for basename

class GoogleSheetScreen(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        # Attribute to store the path of the consolidated file selected on this screen
        self.caminho_arquivo_consolidado_selecionado = None

        # Frame for selecting the Consolidated File
        self.frame_arquivo_consolidado = tk.LabelFrame(self, text="Arquivo Consolidado para Enviar", padx=10, pady=10)
        self.frame_arquivo_consolidado.pack(pady=10, padx=10, fill="x")

        self.label_arquivo_consolidado = tk.Label(self.frame_arquivo_consolidado, text="Nenhum arquivo consolidado selecionado.", wraplength=400)
        self.label_arquivo_consolidado.pack(pady=5)

        self.btn_selecionar_consolidado = tk.Button(self.frame_arquivo_consolidado, text="Selecionar Arquivo Consolidado", command=self.selecionar_arquivo_consolidado)
        self.btn_selecionar_consolidado.pack(pady=5)

        # Frame for the Google Sheet URL
        self.frame_google_sheet = tk.LabelFrame(self, text="URL da Planilha Google (Base de Dados para Inclusão)", padx=10, pady=10)
        self.frame_google_sheet.pack(pady=10, padx=10, fill="x")

        self.label_google_sheet_url = tk.Label(self.frame_google_sheet, text="Nenhuma URL definida.", wraplength=400)
        self.label_google_sheet_url.pack(pady=5)

        self.entry_google_sheet_url = tk.Entry(self.frame_google_sheet, width=60)
        self.entry_google_sheet_url.pack(pady=5)
        self.entry_google_sheet_url.bind("<Return>", self.definir_url_planilha_google_event)

        self.btn_definir_url_google = tk.Button(self.frame_google_sheet, text="Definir URL da Planilha Google", command=self.definir_url_planilha_google)
        self.btn_definir_url_google.pack(pady=5)

        # Button to send data to Google Sheet
        self.btn_incluir_na_base = tk.Button(self, text="Enviar Dados para a Planilha", command=self.incluir_dados_na_base,
                                             bg="#2196F3", fg="white", font=("Arial", 10, "bold"), relief="raised")
        self.btn_incluir_na_base.pack(pady=20)

        # Button to go back to consolidation screen
        self.btn_voltar = tk.Button(self, text="Voltar para Consolidação",
                                     command=lambda: self.controller.show_frame("ConsolidationScreen"),
                                     bg="#607D8B", fg="white", font=("Arial", 10, "bold"), relief="raised")
        self.btn_voltar.pack(pady=5)

        # Status Label
        self.status_label = tk.Label(self, text="", fg="blue", font=("Arial", 9))
        self.status_label.pack(pady=5)

    def selecionar_arquivo_consolidado(self):
        """Opens a file dialog to select the consolidated Excel file."""
        file_path = filedialog.askopenfilename(
            title="Selecione o Arquivo Consolidado",
            filetypes=[("Arquivos Excel", "*.xlsx")]
        )
        if file_path:
            self.caminho_arquivo_consolidado_selecionado = file_path
            self.label_arquivo_consolidado.config(text=os.path.basename(file_path))
            self.status_label.config(text="Arquivo consolidado selecionado.", fg="blue")
        else:
            self.status_label.config(text="Seleção de arquivo consolidado cancelada.", fg="red")

    def definir_url_planilha_google(self):
        url = self.entry_google_sheet_url.get().strip()
        if url:
            self.controller.url_planilha_google_base = url
            self.label_google_sheet_url.config(text=f"URL definida: {url}")
            self.status_label.config(text="URL da Planilha Google definida.", fg="blue")
        else:
            messagebox.showwarning("Atenção", "Por favor, insira uma URL válida para a Planilha Google.")
            self.status_label.config(text="Nenhuma URL definida.", fg="red")

    def definir_url_planilha_google_event(self, event):
        self.definir_url_planilha_google()

    def incluir_dados_na_base(self):
        # Use the file path selected on THIS screen first
        file_to_send = self.caminho_arquivo_consolidado_selecionado

        # If no file was selected on this screen, try to use the one from the controller (if consolidated previously)
        if not file_to_send and self.controller.caminho_arquivo3:
            file_to_send = self.controller.caminho_arquivo3
            self.status_label.config(text="Usando arquivo consolidado da tela anterior.", fg="blue")
            self.label_arquivo_consolidado.config(text=os.path.basename(file_to_send)) # Update label for clarity

        if not file_to_send:
            messagebox.showwarning("Atenção", "Por favor, selecione o arquivo consolidado ou consolide os arquivos na tela anterior.")
            return
        if not self.controller.url_planilha_google_base:
            messagebox.showwarning("Atenção", "Por favor, defina a URL da Planilha Google para inclusão.")
            return

        self.status_label.config(text="Incluindo dados na Planilha Google...", fg="orange")
        self.controller.update_idletasks()

        try:
            from consolidation_logic import inclui_dados_na_base # Import here to avoid circular dependency if not already imported

            url_saida_base = inclui_dados_na_base(
                file_to_send, # Use the determined file path
                self.controller.url_planilha_google_base
            )

            if url_saida_base:
                self.status_label.config(text=f"Sucesso! Dados incluídos na Planilha Google: {url_saida_base}", fg="green")
                messagebox.showinfo("Sucesso", f"Os dados consolidados foram incluídos na Planilha Google em:\n{url_saida_base}")
            else:
                self.status_label.config(text="Operação de inclusão na Planilha Google cancelada ou falhou.", fg="red")

        except FileNotFoundError as e:
            self.status_label.config(text=f"Erro: Arquivo consolidado não encontrado - {e}", fg="red")
            messagebox.showerror("Erro de Arquivo", str(e))
        except Exception as e:
            self.status_label.config(text=f"Erro inesperado durante a inclusão na Planilha Google: {e}", fg="red")
            messagebox.showerror("Erro", f"Ocorreu um erro inesperado: {str(e)}")

