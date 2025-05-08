import win32com.client
import tkinter as tk
from tkinter import messagebox, ttk

# === Janela principal ===
janela = tk.Tk()
janela.title("SolidWorks - Ferramentas")
janela.attributes("-topmost", True)
janela.geometry("1200x600")
janela.configure(bg="#f0f0f0")  # fundo claro tipo VCL

# === Frame superior para os botões ===
frame_botoes = tk.Frame(janela, bg="#f0f0f0")
frame_botoes.pack(fill="x", pady=10)

  # Usando grid para alinhar a label

# === Estilo comum aos botões ===
ESTILO_BOTAO = {
    "height": 2,
    "width": 28,
    "bg": "#e1e1e1",
    "activebackground": "#d0d0d0",
    "relief": "raised",
    "bd": 2,
    "font": ("Segoe UI", 9)
}

# Organizando os botões com grid
from sw_funcs.atualizar_codigo import atualizar_codigo
botao_atualizar = tk.Button(frame_botoes, text="*Atualizar Código*", command=atualizar_codigo, **ESTILO_BOTAO)
botao_atualizar.grid(row=1, column=0, pady=5, padx=5)

from sw_funcs.enviar_dados_para_zoho import enviar_dados_para_zoho
botao_processar = tk.Button(frame_botoes, text="-", command=enviar_dados_para_zoho, **ESTILO_BOTAO)
botao_processar.grid(row=1, column=1, pady=5, padx=5)

from sw_funcs.monitorar_propriedade import monitorar_propriedade
botao_consultar = tk.Button(frame_botoes, text="Preencher Dados", command=monitorar_propriedade, **ESTILO_BOTAO)
botao_consultar.grid(row=1, column=2, pady=5, padx=5)

from sw_funcs.editor import abrir_editor
botao_editar = tk.Button(frame_botoes, text="Modificar Dados", command=lambda: abrir_editor(janela), **ESTILO_BOTAO)
botao_editar.grid(row=1, column=3, pady=5, padx=5)

from sw_funcs.tabela_dados import tabela_dados
botao_tabelar = tk.Button(frame_botoes, text="Tabela de Propriedades", command=lambda: tabela_dados(janela), **ESTILO_BOTAO)
botao_tabelar.grid(row=1, column=4, pady=5, padx=5)

# === Rodapé opcional ===
label_rodape = tk.Label(janela, text="Versão 1.0 • Ferramentas SW", bg="#f0f0f0", font=("Segoe UI", 8))
label_rodape.pack(side="bottom", pady=10)

janela.mainloop()
