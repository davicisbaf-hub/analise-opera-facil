import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import shutil
import os
from main import main as exe_main
from separarRelatorio.main import processar_todos_arquivos_simplificado as processar_arquivos
from tkinter import font as tkfont


import os
import platform
import subprocess

class AbrirPasta:
    @staticmethod
    def abrir(caminho):
        """Abre uma pasta no explorador de arquivos do sistema"""
        if not os.path.isdir(caminho):
            # Tenta criar a pasta se não existir
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


# Cores modernas
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

# Fontes personalizadas
title_font = ("Segoe UI", 24, "bold")
label_font = ("Segoe UI", 11)
button_font = ("Segoe UI", 10, "bold")
entry_font = ("Segoe UI", 10)

# Configuração da janela principal
root = tk.Tk()
root.title("OPERA FÁCIL - Analisador de Relatórios")
root.geometry("700x600")
root.configure(bg='#f0f5ff')
root.resizable(True, True)  # Permitir redimensionamento

# ========== CRIAR CANVAS COM BARRA DE ROLAGEM ==========
# Criar um canvas
canvas = tk.Canvas(root, bg=COLORS['background'], highlightthickness=0)
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Adicionar barra de rolagem vertical
scrollbar = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Configurar o canvas
canvas.configure(yscrollcommand=scrollbar.set)

# Frame principal DENTRO do canvas
main_frame = tk.Frame(canvas, bg=COLORS['background'], padx=30, pady=20)
canvas_frame = canvas.create_window((0, 0), window=main_frame, anchor="nw")

# Configurar rolagem do canvas
def configure_scrollregion(event):
    canvas.configure(scrollregion=canvas.bbox("all"))

main_frame.bind("<Configure>", configure_scrollregion)

# Adicionar bind para rolagem com mouse wheel
def on_mousewheel(event):
    canvas.yview_scroll(int(-1*(event.delta/120)), "units")

canvas.bind_all("<MouseWheel>", on_mousewheel)
# ======================================================

# Cabeçalho
header_frame = tk.Frame(main_frame, bg=COLORS['background'])
header_frame.pack(fill='x', pady=(0, 30))

title_label = tk.Label(
    header_frame,
    text="📊 OPERA FÁCIL",
    font=title_font,
    bg=COLORS['background'],
    fg=COLORS['primary']
)
title_label.pack()

subtitle_label = tk.Label(
    header_frame,
    text="Analisador de Relatórios",
    font=("Segoe UI", 12),
    bg=COLORS['background'],
    fg=COLORS['secondary']
)
subtitle_label.pack(pady=(5, 0))

# Frame de instruções
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
    font=("Segoe UI", 10),
    bg=COLORS['light'],
    fg=COLORS['dark'],
    justify='left',
    padx=15,
    pady=15
)
instructions_label.pack()

# ========== SEÇÃO DE RESULTADOS ==========
results_section_frame = tk.Frame(
    main_frame,
    bg=COLORS['light'],
    relief='flat',
    borderwidth=1,
    highlightbackground='#d1d8e0',
    highlightthickness=1
)
results_section_frame.pack(fill='x', pady=(1, 1))

# Título da seção
results_title = tk.Label(
    results_section_frame,
    text="📁 ABRIR PASTAS DE RESULTADOS",
    font=("Segoe UI", 11, "bold"),
    bg=COLORS['light'],
    fg=COLORS['primary'],
    padx=15,
    
)
results_title.pack(anchor='w')

# Frame para os botões de resultados
results_buttons_container = tk.Frame(results_section_frame, bg=COLORS['light'], padx=15)
results_buttons_container.pack(fill='x')

def abrir_resultados_neomater():
    caminho = os.path.abspath("./Prestador/neomater/resultado")
    os.makedirs(caminho, exist_ok=True)
    AbrirPasta.abrir(caminho)

def abrir_resultados_neotin():
    caminho = os.path.abspath("./Prestador/neotin/resultado")
    os.makedirs(caminho, exist_ok=True)
    AbrirPasta.abrir(caminho)

def abrir_resultados_prontobaby():
    caminho = os.path.abspath("./Prestador/prontobaby/resultado")
    os.makedirs(caminho, exist_ok=True)
    AbrirPasta.abrir(caminho)

# Botões para abrir resultados
def criar_botao_resultado(parent, text, command, color):
    btn = tk.Button(
        parent,
        text=text,
        command=command,
        font=("Segoe UI", 10),
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
        btn['background'] = COLORS['accent']
    
    def on_leave(e):
        btn['background'] = color
    
    btn.bind("<Enter>", on_enter)
    btn.bind("<Leave>", on_leave)
    
    return btn

# Criar botões para cada pasta
btn_result_neomater = criar_botao_resultado(
    results_buttons_container,
    "Neomater",
    abrir_resultados_neomater,
    "#2c3e50"
)

btn_result_neotin = criar_botao_resultado(
    results_buttons_container,
    "Neotin",
    abrir_resultados_neotin,
    "#34495e"
)

btn_result_prontobaby = criar_botao_resultado(
    results_buttons_container,
    "Pronto Baby",
    abrir_resultados_prontobaby,
    "#7f8c8d"
)

# Posicionar botões lado a lado com espaçamento
btn_result_neomater.pack(side='left', padx=(0, 10))
btn_result_neotin.pack(side='left', padx=(0, 10))
btn_result_prontobaby.pack(side='left')

# Texto informativo
results_info = tk.Label(
    results_section_frame,
    text="Clique em qualquer botão acima para abrir a pasta com os relatórios gerados.",
    font=("Segoe UI", 9),
    bg=COLORS['light'],
    fg=COLORS['dark'],
    padx=15,
)
results_info.pack(anchor='w')
# ========================================

# Funções para seleção de arquivos
def select_file1():
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo Neomater",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
    )
    if file_path:
        entry1.delete(0, tk.END)
        entry1.insert(0, file_path)
        status_label1.config(text="✓ Arquivo selecionado", fg=COLORS['success'])

def select_file2():
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo Neotin",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
    )
    if file_path:
        entry2.delete(0, tk.END)
        entry2.insert(0, file_path)
        status_label2.config(text="✓ Arquivo selecionado", fg=COLORS['success'])

def select_file3():
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo Pediátrico",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
    )
    if file_path:
        entry3.delete(0, tk.END)
        entry3.insert(0, file_path)
        status_label3.config(text="✓ Arquivo selecionado", fg=COLORS['success'])

# Frame para os arquivos
files_frame = tk.Frame(main_frame, bg=COLORS['background'])
files_frame.pack(fill='x', pady=(0, 25))

# Estilo para os frames de arquivo
# Dentro da função create_file_frame, RETORNAR o botão também
def create_file_frame(parent, title, select_command, row):
    frame = tk.Frame(parent, bg=COLORS['light'], relief='flat', padx=15, pady=15)
    frame.grid(row=row, column=0, columnspan=3, sticky='ew', pady=8)
    frame.grid_columnconfigure(0, weight=1)
    
    # Título
    title_label = tk.Label(
        frame,
        text=title,
        font=("Segoe UI", 12, "bold"),
        bg=COLORS['light'],
        fg=COLORS['primary']
    )
    title_label.grid(row=0, column=0, sticky='w', pady=(0, 10))
    
    # Campo de entrada
    entry = tk.Entry(
        frame,
        font=entry_font,
        bg='white',
        fg=COLORS['dark'],
        relief='flat',
        borderwidth=1,
        highlightbackground='#bdc3c7',
        highlightthickness=1,
        highlightcolor=COLORS['accent']
    )
    entry.grid(row=1, column=0, sticky='ew', padx=(0, 10))
    
    # Botão - AGORA DEFINIDO AQUI
    button = tk.Button(
        frame,
        text="📁 Procurar",
        command=select_command,
        font=button_font,
        bg=COLORS['primary'],
        fg='white',
        activebackground=COLORS['accent'],
        activeforeground='white',
        relief='flat',
        padx=20,
        pady=5,
        cursor='hand2'
    )
    button.grid(row=1, column=1, padx=(0, 10))
    
    # Label de status
    status_label = tk.Label(
        frame,
        text="❌ Aguardando seleção",
        font=("Segoe UI", 9),
        bg=COLORS['light'],
        fg=COLORS['danger']
    )
    status_label.grid(row=2, column=0, sticky='w', pady=(5, 0))
    
    # RETORNAR também o botão
    return frame, entry, status_label, button  # <-- ADICIONAR button aqui
  
# Criar frames para cada arquivo
file1_frame, entry1, status_label1, button1 = create_file_frame(files_frame, "📋 NEOMATER", select_file1, 0)
file2_frame, entry2, status_label2, button2 = create_file_frame(files_frame, "📋 NEOTIN", select_file2, 1)
file3_frame, entry3, status_label3, button3 = create_file_frame(files_frame, "📋 PEDIÁTRICO", select_file3, 2)
# Função de envio (submit)
def submit():
    file1 = entry1.get()
    file2 = entry2.get()
    file3 = entry3.get()
    
    # Verificar se todos os arquivos foram selecionados
    if not all([file1, file2, file3]):
        messagebox.showwarning("Atenção", "Por favor, selecione todos os 3 arquivos!")
        return
    
    # Confirmar com o usuário
    confirm = messagebox.askyesno(
        "Confirmar Processamento",
        "Deseja processar os arquivos selecionados?\n\n"
        f"• Neomater: {os.path.basename(file1)}\n"
        f"• Neotin: {os.path.basename(file2)}\n"
        f"• Pediátrico: {os.path.basename(file3)}\n\n"
        "Esta operação pode levar alguns minutos."
    )
    
    if not confirm:
        return
    
    # Desabilitar botão durante processamento
    submit_button.config(state='disabled', text="Processando...")
    root.update()
    
    try:
        destination_folder = "./separarRelatorio"
        os.makedirs(destination_folder, exist_ok=True)
        
        # Copiar e renomear arquivos
        files_to_process = [
            (file1, "separarNeomater"),
            (file2, "separarNeotin"),
            (file3, "separarPediatrico")
        ]


        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        for original_file, new_name in files_to_process:
            if original_file:
        # Verificar se é arquivo Excel
              if not original_file.lower().endswith(('.xlsx', '.xls')):
                  messagebox.showerror("Erro", f"O arquivo {os.path.basename(original_file)} não é um arquivo Excel válido!")
                  submit_button.config(state='normal', text="Processar Arquivos")
                  return
              
              # Copiar arquivo
              try:
                  ext = os.path.splitext(original_file)[1]
                  # Adicionar timestamp ao nome do arquivo
                  new_name_with_timestamp = f"{new_name}"
                  new_path = os.path.join(destination_folder, new_name_with_timestamp + ext)
                  
                  # Copiar o arquivo
                  shutil.copy(original_file, new_path)
                  print(f"✅ Arquivo copiado: {os.path.basename(original_file)} -> {new_name_with_timestamp + ext}")
                  
              except Exception as copy_error:
                  messagebox.showerror("Erro", f"Erro ao copiar arquivo:\n{str(copy_error)}")
                  submit_button.config(state='normal', text="Processar Arquivos")
                  return
        
        # Processar arquivos
        try:
            processar_arquivos()
            exe_main()
            
            # Mostrar mensagem de sucesso
            messagebox.showinfo(
                "Sucesso!",
                "✅ Processamento concluído com sucesso!\n\n"
                "Os relatórios foram gerados na pasta:\n"
                "'relatorios_simplificados'"
            )
            
            # Resetar status
            status_label1.config(text="❌ Aguardando seleção", fg=COLORS['danger'])
            status_label2.config(text="❌ Aguardando seleção", fg=COLORS['danger'])
            status_label3.config(text="❌ Aguardando seleção", fg=COLORS['danger'])
            entry1.delete(0, tk.END)
            entry2.delete(0, tk.END)
            entry3.delete(0, tk.END)
            
        except Exception as process_error:
            messagebox.showerror("Erro no Processamento", f"Erro ao processar arquivos:\n{str(process_error)}")
            
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar arquivos:\n{str(e)}")
    
    finally:
        # Reabilitar botão
        submit_button.config(state='normal', text="Processar Arquivos")

# Frame para o botão de ação
action_frame = tk.Frame(main_frame, bg=COLORS['background'])
action_frame.pack(fill='x', pady=(10, 0))

# Botão de processamento
submit_button = tk.Button(
    action_frame,
    text="🚀 Processar Arquivos",
    command=submit,
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

# Footer
footer_frame = tk.Frame(main_frame, bg=COLORS['background'])
footer_frame.pack(fill='x', pady=(25, 0))

footer_label = tk.Label(
    footer_frame,
    text="© 2025 Opera Fácil - Sistema de Análise de Dados Feito por Davi",
    font=("Segoe UI", 9),
    bg=COLORS['background'],
    fg=COLORS['secondary']
)
footer_label.pack()

# Configurar estilo do botão hover
def on_enter(e):
    e.widget['background'] = '#16a085' if e.widget == submit_button else COLORS['accent']

def on_leave(e):
    e.widget['background'] = COLORS['accent'] if e.widget == submit_button else COLORS['primary']

# Aplicar efeitos hover
submit_button.bind("<Enter>", on_enter)
submit_button.bind("<Leave>", on_leave)

# Aplicar efeitos nos botões de busca
for button in [button1, button2, button3]:
    button.bind("<Enter>", on_enter)
    button.bind("<Leave>", on_leave)

# Centralizar a janela
root.update_idletasks()
width = root.winfo_width()
height = root.winfo_height()
x = (root.winfo_screenwidth() // 2) - (width // 2)
y = (root.winfo_screenheight() // 2) - (height // 2)
root.geometry(f'{width}x{height}+{x}+{y}')

root.mainloop()
