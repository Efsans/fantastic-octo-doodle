import win32com.client
import tkinter as tk
from tkinter import ttk, messagebox

# === Janela principal ===
janela = tk.Tk()
janela.title("SolidWorks - Ferramentas")
janela.attributes("-topmost", True)
janela.geometry("1200x600")
janela.configure(bg="#f0f0f0")

# === Estilo de botões nas abas ===
ESTILO_BOTAO = {
    "height": 2,
    "width": 30,
    "bg": "#e1e1e1",
    "activebackground": "#d0d0d0",
    "relief": "raised",
    "bd": 2,
    "font": ("Segoe UI", 10)
}

# === Notebook (abas) ===
notebook = ttk.Notebook(janela)
notebook.pack(expand=True, fill="both", padx=10, pady=10)

# === Importar funções ===
from sw_funcs.atualizar_codigo import atualizar_codigo
from sw_funcs.enviar_dados_para_zoho import enviar_dados_para_zoho
from sw_funcs.monitorar_propriedade import monitorar_propriedade
from sw_funcs.editor import abrir_editor
from sw_funcs.tabela_dados import tabela_dados

# === Criar frames e associar funções ===
abas_info = [
    ("Modificar Dados", abrir_editor, True),
    ("Tabela de Propriedades", tabela_dados, True)
]
frame_botoes = tk.Frame(janela, bg="#f0f0f0")
frame_botoes.pack(pady=10)

from sw_funcs.atualizar_codigo import atualizar_codigo
tk.Button(frame_botoes, text="Atualizar Código", command=atualizar_codigo, height=2, width=25).grid(row=0, column=0, padx=10, pady=5)

from sw_funcs.enviar_dados_para_zoho import enviar_dados_para_zoho
tk.Button(frame_botoes, text="Enviar dados para Zoho", command=enviar_dados_para_zoho, height=2, width=25).grid(row=0, column=1, padx=10, pady=5)

from sw_funcs.monitorar_propriedade import monitorar_propriedade
tk.Button(frame_botoes, text="Consultar e Atualizar", command=monitorar_propriedade, height=2, width=25).grid(row=0, column=2, padx=10, pady=5)


frames_abas = []  # armazenar os frames para acesso posterior

for nome, funcao, autoexec in abas_info:
    frame = tk.Frame(notebook, bg="#f9f9f9")
    notebook.add(frame, text=nome)
    frames_abas.append((frame, funcao, autoexec))

    # Se NÃO for autoexec, adiciona botão manual
    if not autoexec:
        botao = tk.Button(frame, text=nome, command=funcao, **ESTILO_BOTAO)
        botao.pack(row=0, column=0, padx=10, pady=30)

# === Evento para executar função ao mudar de aba ===
def on_tab_changed(event):
    aba_index = notebook.index(notebook.select())
    frame, funcao, autoexec = frames_abas[aba_index]
    
    if autoexec:
        # Limpa a aba antes (opcional)
        for widget in frame.winfo_children():
            widget.destroy()
        # Executa a função passando o frame como destino
        funcao(frame)

notebook.bind("<<NotebookTabChanged>>", on_tab_changed)

# === Rodapé ===
label_rodape = tk.Label(janela, text="Versão 1.0.2 • Ferramentas SW", bg="#f0f0f0", font=("Segoe UI", 8))
label_rodape.pack(side="bottom", pady=10)

janela.mainloop()
