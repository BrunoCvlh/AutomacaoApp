import tkinter as tk
from tkinter import messagebox
import os

from consolidation_screen import ConsolidationScreen
from google_sheet_screen import GoogleSheetScreen

class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Consolidar e Enviar Dados Excel")
        self.geometry("550x550")
        self.resizable(False, False)

        self.caminho_arquivo1 = None
        self.caminho_arquivo2 = None
        self.caminho_arquivo3 = None 
        self.caminho_planilha_base = None

        self.container = tk.Frame(self)
        self.container.pack(side="top", fill="both", expand=True)
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        for F in (ConsolidationScreen, GoogleSheetScreen):
            page_name = F.__name__
            frame = F(parent=self.container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("ConsolidationScreen")

    def show_frame(self, page_name):
        """Mostra uma tela para o usuário."""
        frame = self.frames[page_name]
        frame.tkraise()

if __name__ == "__main__":
    app = MainApp()
    app.mainloop()