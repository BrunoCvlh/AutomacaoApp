import tkinter as tk
from tkinter import messagebox, filedialog
import os

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

        # Frame for the Destination Spreadsheet
        # O texto foi alterado para refletir o uso de planilhas locais
        self.frame_planilha_destino = tk.LabelFrame(self, text="Planilha de Destino (Base de Dados para Inclusão)", padx=10, pady=10)
        self.frame_planilha_destino.pack(pady=10, padx=10, fill="x")

        self.label_planilha_destino_caminho = tk.Label(self.frame_planilha_destino, text="Nenhum caminho definido.", wraplength=400)
        self.label_planilha_destino_caminho.pack(pady=5)

        # Botão para localizar a planilha manualmente
        self.btn_localizar_planilha = tk.Button(self.frame_planilha_destino, text="Localizar Planilha de Destino", command=self.localizar_planilha_destino_manualmente)
        self.btn_localizar_planilha.pack(pady=5)

        # Button to send data to the spreadsheet
        # O texto do botão foi ajustado
        self.btn_incluir_na_base = tk.Button(self, text="Enviar Dados para a Planilha Local", command=self.incluir_dados_na_base,
                                             bg="#2196F3", fg="white", font=("Arial", 10, "bold"), relief="raised")
        self.btn_incluir_na_base.pack(pady=20)

        # Button to go back to consolidation screen
        self.btn_voltar = tk.Button(self, text="Voltar para Consolidação",
                                     command=lambda: self.controller.show_frame("ConsolidationScreen"),
                                     bg="#607D8B", fg="white", font=("Arial", 10, "bold"), relief="raised")
        self.btn_voltar.pack(pady=5)

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

    def localizar_planilha_destino_manualmente(self):
        """Abre uma caixa de diálogo para o usuário selecionar um arquivo de planilha,
        e insere o caminho no campo de caminho da planilha de destino."""
        file_path = filedialog.askopenfilename(
            title="Selecione a Planilha de Destino",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Arquivos CSV", "*.csv"), ("Todos os Arquivos", "*.*")]
        )
        if file_path:
            self.label_planilha_destino_caminho.delete(0, tk.END)
            self.label_planilha_destino_caminho.insert(0, file_path)
            self.status_label.config(text=f"Caminho da planilha preenchido: {os.path.basename(file_path)}", fg="blue")
        else:
            self.status_label.config(text="Seleção de planilha cancelada.", fg="red")

    def definir_caminho_planilha_destino(self):
        """Define o caminho da planilha de destino a partir do campo de entrada."""
        caminho = self.entry_planilha_destino_caminho.get().strip()
        if caminho:
            self.controller.caminho_planilha_base = caminho
            self.label_planilha_destino_caminho.config(text=f"Caminho definido: {caminho}")
            self.status_label.config(text="Caminho da Planilha de Destino definido.", fg="blue")
        else:
            messagebox.showwarning("Atenção", "Por favor, insira um caminho válido para a Planilha de Destino.")
            self.status_label.config(text="Nenhum caminho definido.", fg="red")

    def definir_caminho_planilha_destino_event(self, event):
        self.definir_caminho_planilha_destino()

    def incluir_dados_na_base(self):
        file_to_send = self.caminho_arquivo_consolidado_selecionado

        if not file_to_send and self.controller.caminho_arquivo3:
            file_to_send = self.controller.caminho_arquivo3
            self.status_label.config(text="Usando arquivo consolidado da tela anterior.", fg="blue")
            self.label_arquivo_consolidado.config(text=os.path.basename(file_to_send))

        if not file_to_send:
            messagebox.showwarning("Atenção", "Por favor, selecione o arquivo consolidado ou consolide os arquivos na tela anterior.")
            return
        
        caminho_planilha_destino = self.controller.caminho_planilha_base

        if not caminho_planilha_destino:
            messagebox.showwarning("Atenção", "Por favor, defina o caminho da Planilha de Destino para inclusão.")
            return

        self.status_label.config(text="Incluindo dados na Planilha de Destino...", fg="orange")
        self.controller.update_idletasks()

        try:
            from consolidation_logic import inclui_dados_na_base 

            resultado = inclui_dados_na_base(
                file_to_send,
                caminho_planilha_destino
            )

            if resultado:
                self.status_label.config(text=f"Sucesso! Dados incluídos na Planilha de Destino: {caminho_planilha_destino}", fg="green")
                messagebox.showinfo("Sucesso", f"Os dados consolidados foram incluídos na Planilha de Destino em:\n{caminho_planilha_destino}")
            else:
                self.status_label.config(text="Operação de inclusão na Planilha de Destino cancelada ou falhou.", fg="red")

        except FileNotFoundError as e:
            self.status_label.config(text=f"Erro: Arquivo não encontrado - {e}", fg="red")
            messagebox.showerror("Erro de Arquivo", str(e))
        except Exception as e:
            self.status_label.config(text=f"Erro inesperado durante a inclusão na Planilha de Destino: {e}", fg="red")
            messagebox.showerror("Erro", f"Ocorreu um erro inesperado: {str(e)}")