import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import shutil
import os
from datetime import datetime
from pathlib import Path
import platform
import sys
import traceback
import logging
import subprocess
import time

# Importações locais
from main import main
from separarRelatorio.main import processar_todos_arquivos_simplificado as processar_arquivos

# ============================================================================
# CONFIGURAÇÃO DE LOGGING
# ============================================================================

logging.basicConfig(
    filename='./app.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def setup_exception_handler():
    """Configura o handler global de exceções"""
    def handle_exception(exc_type, exc_value, exc_traceback):
        logging.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))
        messagebox.showerror("Erro", f"Ocorreu um erro: {exc_value}")
    
    sys.excepthook = handle_exception

# ============================================================================
# CONSTANTES E CONFIGURAÇÕES
# ============================================================================

COLORS = {
    'primary': '#3a7ca5',
    'secondary': '#2c3e50',
    'accent': '#1abc9c',
    'light': '#ecf0f1',
    'dark': '#2c3e50',
    'success': '#27ae60',
    'warning': '#f39c12',
    'danger': '#e74c3c',
    'background': '#f0f5ff'
}

FONTS = {
    'title': ("Segoe UI", 24, "bold"),
    'subtitle': ("Segoe UI", 12),
    'label': ("Segoe UI", 11),
    'button': ("Segoe UI", 10, "bold"),
    'entry': ("Segoe UI", 10),
    'small': ("Segoe UI", 9)
}

FILE_CONFIGS = [
    {"title": "📋 NEOMATER", "key": "neomater", "dialog_title": "Selecione o arquivo Neomater"},
    {"title": "📋 NEOTIN", "key": "neotin", "dialog_title": "Selecione o arquivo Neotin"},
    {"title": "📋 PEDIÁTRICO", "key": "pediatrico", "dialog_title": "Selecione o arquivo Pediátrico"}
]

DESTINATION_MAPPING = {
    "neomater": "separarNeomater",
    "neotin": "separarNeotin",
    "pediatrico": "separarPediatrico"
}

# ============================================================================
# CLASSES AUXILIARES
# ============================================================================

class AbrirPasta:
    """Classe utilitária para abrir pastas no explorador de arquivos"""
    
    @staticmethod
    def abrir(caminho):
        """Abre uma pasta no explorador de arquivos do sistema"""
        if not os.path.isdir(caminho):
            try:
                os.makedirs(caminho, exist_ok=True)
            except Exception as e:
                print(f"Erro ao criar pasta: {e}")
                return
        
        sistema = platform.system()
        
        try:
            if sistema == "Windows":
                os.startfile(caminho)
            elif sistema == "Darwin":  # macOS
                subprocess.Popen(["open", caminho])
            else:  # Linux
                subprocess.Popen(["xdg-open", caminho])
        except Exception as e:
            print(f"Erro ao abrir a pasta: {e}")


class FileFrame:
    """Classe para criar um frame de seleção de arquivo"""
    
    def __init__(self, parent, title, dialog_title, file_key):
        self.parent = parent
        self.title = title
        self.dialog_title = dialog_title
        self.file_key = file_key
        self.file_path = tk.StringVar()
        self.create_widgets()
    
    def create_widgets(self):
        """Cria os widgets do frame de arquivo"""
        self.frame = tk.Frame(
            self.parent, 
            bg=COLORS['light'], 
            relief='flat', 
            padx=15, 
            pady=15
        )
        
        # Título
        self.title_label = tk.Label(
            self.frame,
            text=self.title,
            font=("Segoe UI", 12, "bold"),
            bg=COLORS['light'],
            fg=COLORS['primary']
        )
        self.title_label.grid(row=0, column=0, sticky='w', pady=(0, 10))
        
        # Campo de entrada
        self.entry = tk.Entry(
            self.frame,
            font=FONTS['entry'],
            bg='white',
            fg=COLORS['dark'],
            relief='flat',
            borderwidth=1,
            highlightbackground='#bdc3c7',
            highlightthickness=1,
            highlightcolor=COLORS['accent']
        )
        self.entry.grid(row=1, column=0, sticky='ew', padx=(0, 10))
        
        # Botão de busca
        self.button = tk.Button(
            self.frame,
            text="📁 Procurar",
            command=self.select_file,
            font=FONTS['button'],
            bg=COLORS['primary'],
            fg='white',
            activebackground=COLORS['accent'],
            activeforeground='white',
            relief='flat',
            padx=20,
            pady=5,
            cursor='hand2'
        )
        self.button.grid(row=1, column=1, padx=(0, 10))
        
        # Label de status
        self.status_label = tk.Label(
            self.frame,
            text="❌ Aguardando seleção",
            font=FONTS['small'],
            bg=COLORS['light'],
            fg=COLORS['danger']
        )
        self.status_label.grid(row=2, column=0, sticky='w', pady=(5, 0))
        
        # Configurar efeitos hover
        self.setup_hover_effects()
    
    def setup_hover_effects(self):
        """Configura efeitos hover para o botão"""
        def on_enter(e):
            e.widget['background'] = COLORS['accent']
        
        def on_leave(e):
            e.widget['background'] = COLORS['primary']
        
        self.button.bind("<Enter>", on_enter)
        self.button.bind("<Leave>", on_leave)
    
    def select_file(self):
        """Abre diálogo para seleção de arquivo"""
        file_path = filedialog.askopenfilename(
            title=self.dialog_title,
            filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
        )
        if file_path:
            self.entry.delete(0, tk.END)
            self.entry.insert(0, file_path)
            self.status_label.config(text="✓ Arquivo selecionado", fg=COLORS['success'])
    
    def get_path(self):
        """Retorna o caminho do arquivo selecionado"""
        return self.entry.get()
    
    def reset(self):
        """Reseta o frame para o estado inicial"""
        self.entry.delete(0, tk.END)
        self.status_label.config(text="❌ Aguardando seleção", fg=COLORS['danger'])


class ResultsSection:
    """Classe para a seção de resultados"""
    
    def __init__(self, parent):
        self.parent = parent
        self.create_widgets()
    
    def create_widgets(self):
        """Cria a seção de resultados"""
        # Frame principal
        self.frame = tk.Frame(
            self.parent,
            bg=COLORS['light'],
            relief='flat',
            borderwidth=1,
            highlightbackground='#d1d8e0',
            highlightthickness=1
        )
        self.frame.pack(fill='x', pady=(1, 1))
        
        # Título
        self.title_label = tk.Label(
            self.frame,
            text="📁 ABRIR PASTAS DE RESULTADOS",
            font=("Segoe UI", 11, "bold"),
            bg=COLORS['light'],
            fg=COLORS['primary'],
            padx=15,
        )
        self.title_label.pack(anchor='w')
        
        # Container para botões
        self.buttons_container = tk.Frame(self.frame, bg=COLORS['light'], padx=15)
        self.buttons_container.pack(fill='x')
        
        # Criar botões
        self.create_result_buttons()
        
        # Texto informativo
        self.info_label = tk.Label(
            self.frame,
            text="Clique em qualquer botão acima para abrir a pasta com os relatórios gerados.",
            font=FONTS['small'],
            bg=COLORS['light'],
            fg=COLORS['dark'],
            padx=15,
        )
        self.info_label.pack(anchor='w')
    
    def create_result_buttons(self):
        """Cria os botões para abrir pastas de resultados"""
        button_configs = [
            {
                "text": "Neomater",
                "command": self.abrir_resultados_neomater,
                "color": "#2c3e50"
            },
            {
                "text": "Neotin",
                "command": self.abrir_resultados_neotin,
                "color": "#34495e"
            },
            {
                "text": "Pronto Baby",
                "command": self.abrir_resultados_prontobaby,
                "color": "#7f8c8d"
            }
        ]
        
        self.buttons = []
        for i, config in enumerate(button_configs):
            btn = self.create_button(
                self.buttons_container,
                config["text"],
                config["command"],
                config["color"]
            )
            btn.pack(side='left', padx=(0, 10) if i < len(button_configs) - 1 else (0, 0))
            self.buttons.append(btn)
    
    @staticmethod
    def create_button(parent, text, command, color):
        """Cria um botão estilizado"""
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            font=FONTS['entry'],
            bg=color,
            fg='white',
            relief='flat',
            padx=20,
            pady=10,
            cursor='hand2',
            width=15
        )
        
        # Efeitos hover
        def on_enter(e):
            e.widget['background'] = COLORS['accent']
        
        def on_leave(e):
            e.widget['background'] = color
        
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        
        return btn
    
    @staticmethod
    def abrir_resultados_neomater():
        """Abre pasta de resultados do Neomater"""
        caminho = Path("../Prestador/neomater/resultado").resolve(strict=False)
        os.makedirs(caminho, exist_ok=True)
        AbrirPasta.abrir(caminho)
    
    @staticmethod
    def abrir_resultados_neotin():
        """Abre pasta de resultados do Neotin"""
        caminho = Path("../Prestador/neotin/resultado").resolve(strict=False)
        os.makedirs(caminho, exist_ok=True)
        AbrirPasta.abrir(caminho)
    
    @staticmethod
    def abrir_resultados_prontobaby():
        """Abre pasta de resultados do Pronto Baby"""
        caminho = Path("../Prestador/prontobaby/resultado").resolve(strict=False)
        os.makedirs(caminho, exist_ok=True)
        AbrirPasta.abrir(caminho)


# ============================================================================
# FUNÇÕES PRINCIPAIS
# ============================================================================

def copy_and_rename_files(file_frames):
    """Copia e renomeia os arquivos selecionados"""
    destination_folder = "../separarRelatorio"
    os.makedirs(destination_folder, exist_ok=True)
    
    files_to_process = []
    
    # Coletar arquivos válidos
    for file_frame, config in zip(file_frames, FILE_CONFIGS):
        file_path = file_frame.get_path()
        if file_path:
            if not file_path.lower().endswith(('.xlsx', '.xls')):
                raise ValueError(
                    f"O arquivo {os.path.basename(file_path)} não é um arquivo Excel válido!"
                )
            
            new_name = DESTINATION_MAPPING[config["key"]]
            files_to_process.append((file_path, new_name))
    
    # Processar cada arquivo
    for original_file, new_name in files_to_process:
        ext = os.path.splitext(original_file)[1]
        new_path = os.path.join(destination_folder, new_name + ext)
        shutil.copy(original_file, new_path)
        print(f"✅ Arquivo copiado: {os.path.basename(original_file)} -> {new_name + ext}")
    
    return len(files_to_process) > 0  # Retorna True se pelo menos um arquivo foi processado


def process_submit(file_frames, submit_button):
    """Função principal de processamento dos arquivos"""
    # Coletar informações dos arquivos
    file_info = []
    for i, file_frame in enumerate(file_frames):
        file_path = file_frame.get_path()
        if file_path:
            file_info.append(f"• {FILE_CONFIGS[i]['title']}: {os.path.basename(file_path)}")
    
    if not file_info:
        messagebox.showwarning("Atenção", "Por favor, selecione pelo menos um arquivo!")
        return
    
    # Confirmar com o usuário
    confirm_message = "Deseja processar os arquivos selecionados?\n\n" + "\n".join(file_info)
    confirm_message += "\n\nEsta operação pode levar alguns minutos."
    
    if not messagebox.askyesno("Confirmar Processamento", confirm_message):
        return
    
    # Desabilitar botão durante processamento
    submit_button.config(state='disabled', text="Processando...")
    submit_button.update_idletasks()
    
    try:
        # Copiar e renomear arquivos
        if not copy_and_rename_files(file_frames):
            messagebox.showwarning("Atenção", "Nenhum arquivo válido para processar!")
            return
        time.sleep(2)
        
        # Processar arquivos
        print("Iniciando processamento de arquivos...")
        processar_arquivos()
        time.sleep(4)
        
        print("Iniciando análise...")
        main()
        time.sleep(4)
        
        print("Processamento finalizado!")
        
        
        # Mensagem de sucesso
        messagebox.showinfo(
            "Sucesso!",
            "✅ Processamento concluído com sucesso!\n\n"
            "Os relatórios foram gerados na pasta:\n"
            "'relatorios_simplificados'"
        )

        
        
        # Resetar frames
        for file_frame in file_frames:
            file_frame.reset()
        
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar arquivos:\n{str(e)}")
        logging.error(f"Erro no processamento: {e}", exc_info=True)
    
    finally:
        # Reabilitar botão
        submit_button.config(state='normal', text="🚀 Processar Arquivos")


def create_main_window():
    """Cria e configura a janela principal"""
    root = tk.Tk()
    root.title("OPERA FÁCIL - Analisador de Relatórios")
    root.geometry("700x600")
    root.configure(bg=COLORS['background'])
    root.resizable(True, True)
    
    return root


def create_scrollable_canvas(root):
    """Cria um canvas com barra de rolagem"""
    # Canvas principal
    canvas = tk.Canvas(root, bg=COLORS['background'], highlightthickness=0)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    
    # Barra de rolagem
    scrollbar = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    # Configurar canvas
    canvas.configure(yscrollcommand=scrollbar.set)
    
    # Frame principal dentro do canvas
    main_frame = tk.Frame(canvas, bg=COLORS['background'], padx=30, pady=20)
    canvas_frame = canvas.create_window((0, 0), window=main_frame, anchor="nw")
    
    # Configurar rolagem
    def configure_scrollregion(event):
        canvas.configure(scrollregion=canvas.bbox("all"))
    
    main_frame.bind("<Configure>", configure_scrollregion)
    
    # Configurar rolagem com mouse wheel
    def on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    
    canvas.bind_all("<MouseWheel>", on_mousewheel)
    
    return main_frame, canvas


def create_header(main_frame):
    """Cria o cabeçalho da aplicação"""
    header_frame = tk.Frame(main_frame, bg=COLORS['background'])
    header_frame.pack(fill='x', pady=(0, 30))
    
    # Título principal
    title_label = tk.Label(
        header_frame,
        text="📊 OPERA FÁCIL",
        font=FONTS['title'],
        bg=COLORS['background'],
        fg=COLORS['primary']
    )
    title_label.pack()
    
    # Subtítulo
    subtitle_label = tk.Label(
        header_frame,
        text="Analisador de Relatórios",
        font=FONTS['subtitle'],
        bg=COLORS['background'],
        fg=COLORS['secondary']
    )
    subtitle_label.pack(pady=(5, 0))
    
    return header_frame


def create_instructions(main_frame):
    """Cria a seção de instruções"""
    instructions_frame = tk.Frame(
        main_frame,
        bg=COLORS['light'],
        relief='flat',
        borderwidth=1,
        highlightbackground='#d1d8e0',
        highlightthickness=1
    )
    instructions_frame.pack(fill='x', pady=(0, 25))
    
    instructions_text = """Selecione os 3 arquivos Excel para análise:
1. Neomater
2. Neotin  
3. Pediátrico

Os arquivos serão copiados, renomeados e processados automaticamente."""
    
    instructions_label = tk.Label(
        instructions_frame,
        text=instructions_text,
        font=FONTS['entry'],
        bg=COLORS['light'],
        fg=COLORS['dark'],
        justify='left',
        padx=15,
        pady=15
    )
    instructions_label.pack()
    
    return instructions_frame


def create_file_selection(main_frame):
    """Cria a seção de seleção de arquivos"""
    files_frame = tk.Frame(main_frame, bg=COLORS['background'])
    files_frame.pack(fill='x', pady=(0, 25))
    
    # Criar frames para cada arquivo
    file_frames = []
    for config in FILE_CONFIGS:
        file_frame = FileFrame(
            files_frame,
            config["title"],
            config["dialog_title"],
            config["key"]
        )
        file_frame.frame.grid(row=len(file_frames), column=0, columnspan=3, sticky='ew', pady=8)
        file_frames.append(file_frame)
    
    return files_frame, file_frames


def create_action_button(main_frame, file_frames):
    """Cria o botão de ação principal"""
    action_frame = tk.Frame(main_frame, bg=COLORS['background'])
    action_frame.pack(fill='x', pady=(10, 0))
    
    # Botão de processamento
    submit_button = tk.Button(
        action_frame,
        text="🚀 Processar Arquivos",
        command=lambda: process_submit(file_frames, submit_button),
        font=("Segoe UI", 13, "bold"),
        bg=COLORS['accent'],
        fg='white',
        activebackground='#16a085',
        activeforeground='white',
        relief='flat',
        padx=40,
        pady=15,
        cursor='hand2',
        borderwidth=0
    )
    submit_button.pack()
    
    # Configurar efeitos hover
    def on_enter(e):
        e.widget['background'] = '#16a085'
    
    def on_leave(e):
        e.widget['background'] = COLORS['accent']
    
    submit_button.bind("<Enter>", on_enter)
    submit_button.bind("<Leave>", on_leave)
    
    return submit_button


def create_footer(main_frame):
    """Cria o rodapé da aplicação"""
    footer_frame = tk.Frame(main_frame, bg=COLORS['background'])
    footer_frame.pack(fill='x', pady=(25, 0))
    
    footer_label = tk.Label(
        footer_frame,
        text="© 2025 Opera Fácil - Sistema de Análise de Dados Feito por Davi",
        font=FONTS['small'],
        bg=COLORS['background'],
        fg=COLORS['secondary']
    )
    footer_label.pack()
    
    return footer_frame


def center_window(root):
    """Centraliza a janela na tela"""
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')


# ============================================================================
# FUNÇÃO PRINCIPAL
# ============================================================================

def app():
    """Função principal da aplicação"""
    # Configurar handler de exceções
    setup_exception_handler()
    
    # Criar janela principal
    root = create_main_window()
    
    # Criar canvas com rolagem
    main_frame, _ = create_scrollable_canvas(root)
    
    # Criar componentes da interface
    create_header(main_frame)
    create_instructions(main_frame)
    
    # Seção de resultados
    results_section = ResultsSection(main_frame)
    
    # Seção de seleção de arquivos
    _, file_frames = create_file_selection(main_frame)
    
    # Botão de ação
    submit_button = create_action_button(main_frame, file_frames)
    
    # Rodapé
    create_footer(main_frame)
    
    # Centralizar janela
    center_window(root)
    
    # Iniciar loop principal
    root.mainloop()

if __name__ == "__main__":
    app()