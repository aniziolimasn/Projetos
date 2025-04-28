# Aplicativo criado para dividir um arquivo PDF em múltiplos arquivos menores e renomeá-los com base em uma planilha Excel fornecida.
# O usuário pode especificar o número de páginas por arquivo e a pasta de saída para os arquivos gerados.
# A aplicação também inclui uma interface gráfica para facilitar a interação do usuário.
# Logs são gerados para rastrear o progresso e possíveis erros durante o processamento.

# Importação das bibliotecas necessárias
import pandas as pd  # Para manipulação de planilhas Excel
from PyPDF2 import PdfReader, PdfWriter  # Para manipulação de arquivos PDF
import tkinter as tk  # Para interface gráfica
from tkinter import filedialog, messagebox, ttk  # Componentes da interface
import os  # Para operações com arquivos e diretórios
from typing import Optional  # Para tipagem estática
import re  # Para expressões regulares
import logging

# Configuração de logs
logging.basicConfig(filename="app.log", level=logging.INFO, format="%(asctime)s - %(message)s")

class PDFSplitter:
    """
    Classe responsável pela lógica de divisão do PDF.
    Esta classe encapsula toda a lógica de negócio relacionada à divisão do PDF
    e manipulação dos arquivos.
    """
    
    def __init__(self):
        """Inicializa o divisor de PDF com flag de cancelamento"""
        self.cancel_operation = False

    def validate_excel_data(self, df: pd.DataFrame) -> bool:
        """
        Valida se a planilha Excel tem o formato correto.
        
        Args:
            df: DataFrame contendo os dados da planilha Excel
            
        Returns:
            bool: True se os dados são válidos, False caso contrário
        """
        if df.empty:
            messagebox.showerror("Erro", "A planilha está vazia.")
            return False
        if df.iloc[:, 0].isnull().any():
            messagebox.showerror("Erro", "Existem células vazias na primeira coluna da planilha.")
            return False
        return True

    def sanitize_filename(self, filename: str) -> str:
        """
        Remove caracteres inválidos do nome do arquivo.
        
        Args:
            filename: Nome do arquivo a ser sanitizado
            
        Returns:
            str: Nome do arquivo sanitizado
        """
        # Remove caracteres inválidos para nome de arquivo
        sanitized = re.sub(r'[<>:"/\\|?*]', '', filename)
        return sanitized.strip()

    def dividir_pdf(self, pdf_path: str, excel_path: str, paginas_por_arquivo: int, 
                    output_folder: str, progress_callback) -> bool:
        """
        Divide o PDF em múltiplos arquivos baseado na planilha Excel.
        
        Args:
            pdf_path: Caminho do arquivo PDF
            excel_path: Caminho da planilha Excel
            paginas_por_arquivo: Número de páginas por arquivo
            output_folder: Pasta de destino
            progress_callback: Função para atualizar o progresso
        
        Returns:
            bool: True se a operação foi bem sucedida, False caso contrário
        """
        try:
            # Carrega a planilha Excel
            df = pd.read_excel(excel_path)
            if not self.validate_excel_data(df):
                return False

            # Abre o arquivo PDF
            with open(pdf_path, 'rb') as pdf_file:
                reader = PdfReader(pdf_file)
                total_paginas = len(reader.pages)
                
                # Calcula o número total de arquivos que serão gerados
                total_steps = (total_paginas // paginas_por_arquivo) + (1 if total_paginas % paginas_por_arquivo != 0 else 0)
                
                # Verifica se há nomes suficientes na planilha
                if total_steps > len(df):
                    messagebox.showerror("Erro", 
                        f"A planilha contém menos nomes ({len(df)}) do que o número de arquivos que serão gerados ({total_steps}).")
                    return False
                
                # Processa cada bloco de páginas
                for i in range(0, total_paginas, paginas_por_arquivo):
                    if self.cancel_operation:
                        logging.info("Operação cancelada pelo usuário.")
                        return False

                    # Cria um novo PDF para o bloco atual
                    writer = PdfWriter()
                    for j in range(i, min(i + paginas_por_arquivo, total_paginas)):
                        writer.add_page(reader.pages[j])
                    
                    # Gera o nome do arquivo e salva
                    novo_nome = self.sanitize_filename(str(df.iloc[i // paginas_por_arquivo, 0])) + ".pdf"
                    output_path = os.path.join(output_folder, novo_nome)
                    
                    try:
                        with open(output_path, 'wb') as novo_pdf:
                            writer.write(novo_pdf)
                        logging.info(f"Arquivo {novo_nome} criado com sucesso.")
                    except Exception as e:
                        logging.error(f"Erro ao salvar arquivo {novo_nome}: {e}")
                        messagebox.showerror("Erro", f"Erro ao salvar arquivo {novo_nome}: {str(e)}")
                        return False
                    
                    # Atualiza o progresso
                    progress = ((i // paginas_por_arquivo + 1) / total_steps * 100)
                    progress_callback(progress)
                    
            logging.info("Processo concluído com sucesso.")
            return True
        except Exception as e:
            logging.error(f"Erro durante o processamento: {e}")
            messagebox.showerror("Erro", f"Erro durante o processamento: {str(e)}")
            return False

class Application:
    """
    Classe responsável pela interface gráfica.
    Esta classe gerencia toda a interface do usuário e interação com o PDFSplitter.
    """
    
    def __init__(self):
        """Inicializa a aplicação e configura a interface"""
        self.root = tk.Tk()
        self.root.title("Divisor de PDF")
        self.pdf_splitter = PDFSplitter()
        
        # Variáveis para armazenar os valores dos campos
        self.pdf_path = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.paginas_por_arquivo = tk.StringVar()
        self.progress_var = tk.DoubleVar()
        
        self.setup_ui()
        
    def setup_ui(self):
        """
        Configura a interface do usuário.
        Cria e organiza todos os elementos visuais da aplicação.
        """
        # Frame principal com padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Campo para selecionar o arquivo PDF
        ttk.Label(main_frame, text="Arquivo PDF:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(main_frame, textvariable=self.pdf_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="Selecionar", command=self.selecionar_pdf).grid(row=0, column=2, padx=5, pady=5)

        # Campo para selecionar o arquivo Excel
        ttk.Label(main_frame, text="Arquivo Excel:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(main_frame, textvariable=self.excel_path, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="Selecionar", command=self.selecionar_excel).grid(row=1, column=2, padx=5, pady=5)

        # Campo para número de páginas por arquivo
        ttk.Label(main_frame, text="Páginas por Arquivo:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(main_frame, textvariable=self.paginas_por_arquivo, width=50).grid(row=2, column=1, padx=5, pady=5)

        # Campo para selecionar a pasta de saída
        ttk.Label(main_frame, text="Pasta de Saída:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(main_frame, textvariable=self.output_folder, width=50).grid(row=3, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="Selecionar", command=self.selecionar_pasta_saida).grid(row=3, column=2, padx=5, pady=5)

        # Barra de progresso
        ttk.Label(main_frame, text="Progresso:").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100).grid(row=4, column=1, columnspan=2, padx=5, pady=5, sticky="we")

        # Frame para botões
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=20)

        # Botões de ação
        ttk.Button(button_frame, text="Iniciar", command=self.iniciar_divisao).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancelar", command=self.cancelar_operacao).pack(side=tk.LEFT, padx=5)

        self.centralizar_janela()
    
    def create_tooltip(self, widget, text):
        """
        Cria um tooltip para um widget.
        
        Args:
            widget: Widget ao qual o tooltip será associado
            text: Texto do tooltip
        """
        def show_tooltip(event):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
            
            label = ttk.Label(tooltip, text=text, background="#ffffe0", relief="solid", borderwidth=1)
            label.pack()
            
            def hide_tooltip():
                tooltip.destroy()
            
            widget.tooltip = tooltip
            widget.bind('<Leave>', lambda e: hide_tooltip())
            
        widget.bind('<Enter>', show_tooltip)
    
    def selecionar_pdf(self):
        """Abre diálogo para selecionar arquivo PDF"""
        self.pdf_path.set(filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")]))
    
    def selecionar_excel(self):
        """Abre diálogo para selecionar arquivo Excel"""
        self.excel_path.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")]))
    
    def selecionar_pasta_saida(self):
        """Abre diálogo para selecionar pasta de saída"""
        self.output_folder.set(filedialog.askdirectory())
    
    def update_progress(self, value):
        """
        Atualiza a barra de progresso.
        
        Args:
            value: Valor do progresso (0-100)
        """
        self.progress_var.set(value)
        self.root.update_idletasks()
    
    def cancelar_operacao(self):
        """Cancela a operação em andamento"""
        self.pdf_splitter.cancel_operation = True
    
    def validar_campos(self) -> bool:
        """
        Valida os campos antes de iniciar a operação.
        
        Returns:
            bool: True se todos os campos são válidos, False caso contrário
        """
        if not all([self.pdf_path.get(), self.excel_path.get(), self.output_folder.get()]):
            messagebox.showerror("Erro", "Todos os campos devem ser preenchidos.")
            return False
        
        if not self.paginas_por_arquivo.get().isdigit():
            messagebox.showerror("Erro", "O número de páginas deve ser um valor numérico.")
            return False
        
        return True
    
    def iniciar_divisao(self):
        """Inicia o processo de divisão do PDF"""
        if not self.validar_campos():
            return
        
        self.pdf_splitter.cancel_operation = False
        success = self.pdf_splitter.dividir_pdf(
            self.pdf_path.get(),
            self.excel_path.get(),
            int(self.paginas_por_arquivo.get()),
            self.output_folder.get(),
            self.update_progress
        )
        
        if success:
            messagebox.showinfo("Concluído", "Processo concluído com sucesso!")
        elif not self.pdf_splitter.cancel_operation:
            messagebox.showerror("Erro", "Ocorreu um erro durante o processamento.")
    
    def centralizar_janela(self):
        """Centraliza a janela principal na tela"""
        self.root.update_idletasks()
        largura = self.root.winfo_width()
        altura = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (largura // 2)
        y = (self.root.winfo_screenheight() // 2) - (altura // 2)
        self.root.geometry(f"{largura}x{altura}+{x}+{y}")

    def configurar_arrastar_soltar(self):
        """Configura suporte para arrastar e soltar arquivos"""
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.arrastar_arquivo)

    def arrastar_arquivo(self, event):
        caminho = event.data.strip()
        if caminho.endswith(".pdf"):
            self.pdf_path.set(caminho)
        elif caminho.endswith(".xlsx"):
            self.excel_path.set(caminho)

    def run(self):
        """Inicia a aplicação"""
        self.root.mainloop()

# Ponto de entrada da aplicação
if __name__ == "__main__":
    app = Application()
    app.run()