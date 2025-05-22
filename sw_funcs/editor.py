import tkinter as tk
from tkinter import messagebox, ttk, StringVar, Label
from sw_funcs.gereprop import get_mgr
from sw_funcs.prefixo import prefixo
import os
from sw_funcs.validador import validar_todos_campos
from sw_funcs.descri_tipo import regras
from sw_funcs.descri_tipo import tipo_prefix
from sw_funcs.aba_adicionar import criar_aba_adicionar

def abrir_editor(janela):
    campos_fixos = ["Codigo", "Descrição", "Tipo", "Unidade", "Pos.IPI/NCM", "Origem"]
    entradas_fixas = {}
    outras_props = {}
    salvar_adicionar = {}
    historico_campos = {}

    for widget in janela.pack_slaves():
        if getattr(widget, "editor_embutido", False):
            widget.destroy()

    editor = tk.Frame(janela, height=300, bg="#f0f0f0")
    editor.pack(side="bottom", fill="both", pady=5)
    editor.editor_embutido = True

    notebook = ttk.Notebook(editor)
    notebook.pack(expand=True, fill="both", padx=5, pady=5)

    # ================== Aba: Campos fixos ==================
    aba_fixos = tk.Frame(notebook)
    notebook.add(aba_fixos, text="Campos fixos")

    try:
        mgr, _ = get_mgr()
        nomes = mgr.GetNames
        propriedades_existentes = {
            nome: mgr.Get(nome)[0] if isinstance(mgr.Get(nome), tuple) else mgr.Get(nome)
            for nome in nomes or []
        }
    except:
        propriedades_existentes = {}

    def recarregar_campos_fixos():
        # Re-obtem propriedades e atualiza os campos
        try:
            mgr, _ = get_mgr()
            nomes = mgr.GetNames
            propriedades_atualizadas = {
                nome: mgr.Get(nome)[0] if isinstance(mgr.Get(nome), tuple) else mgr.Get(nome)
                for nome in nomes or []
            }
        except:
            propriedades_atualizadas = {}

        for nome, entry in entradas_fixas.items():
            entry.delete(0, tk.END)
            if nome in propriedades_atualizadas:
                entry.insert(0, propriedades_atualizadas[nome])

    # Botão recarregar no canto superior esquerdo
    btn_recarregar = tk.Button(aba_fixos, text="⟳ Recarregar", command=recarregar_campos_fixos)
    btn_recarregar.grid(row=0, column=0, sticky="w", padx=10, pady=(8, 2), columnspan=2)

    # Comece os campos na linha 1!
    for idx, nome in enumerate(campos_fixos):
        tk.Label(aba_fixos, text=nome + ":", anchor="w", width=16).grid(row=idx+1, column=0, sticky="w", padx=(10, 5), pady=6)
        entry = tk.Entry(aba_fixos)
        entry.grid(row=idx+1, column=1, sticky="ew", padx=(0, 10), pady=6)
        aba_fixos.columnconfigure(1, weight=1)
        if nome in propriedades_existentes:
            entry.insert(0, propriedades_existentes[nome])
        entradas_fixas[nome] = entry
    # ================== Aba: Campos SolidWorks ==================
    aba_outros = tk.Frame(notebook)
    notebook.add(aba_outros, text="Campos SolidWorks")

    canvas = tk.Canvas(aba_outros)
    scrollbar = tk.Scrollbar(aba_outros, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    frame_interno = tk.Frame(canvas)
    canvas_window = canvas.create_window((0, 0), window=frame_interno, anchor="nw")

    def ajustar_largura(event):
        canvas.itemconfig(canvas_window, width=event.width)
        canvas.configure(scrollregion=canvas.bbox("all"))

    canvas.bind("<Configure>", ajustar_largura)

    # Adicionar novo campo (linha de adicionar)
    frame_add = tk.Frame(frame_interno)
    frame_add.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(8, 12), padx=12)
    frame_add.columnconfigure(1, weight=1)
    frame_add.columnconfigure(3, weight=1)

    tk.Label(frame_add, text="Nome:").grid(row=0, column=0, sticky="w")
    entrada_nome_novo = tk.Entry(frame_add, width=20)
    entrada_nome_novo.grid(row=0, column=1, sticky="ew", padx=6)

    tk.Label(frame_add, text="Valor:").grid(row=0, column=2, sticky="w")
    entrada_valor_novo = tk.Entry(frame_add, width=25)
    entrada_valor_novo.grid(row=0, column=3, sticky="ew", padx=6)

    btn_add = tk.Button(frame_add, text="Adicionar Campo")
    btn_add.grid(row=0, column=4, padx=(12, 0))

    # Função para adicionar campos na interface
    def adicionar_campo_interface(nome, valor):
        idx = adicionar_campo_interface.linha_idx
        linha = (idx // 2) + 1   # Começa em 1 para pular frame_add na linha 0
        coluna = idx % 2

        frame = tk.Frame(frame_interno)
        frame.grid(row=linha, column=coluna, sticky="ew", padx=11, pady=4)
        frame.grid_columnconfigure(1, weight=1)

        rotulo = tk.Label(frame, text=nome, width=16, anchor=tk.E)
        rotulo.grid(row=0, column=0, sticky="w", padx=(0, 6))

        entry = tk.Entry(frame, width=50)
        entry.grid(row=0, column=1, sticky="ew", padx=(0, 6))
        entry.insert(0, valor)
        outras_props[nome] = entry
        historico_campos[nome] = [valor]

        def limpar_campo():
            valor_atual = entry.get()
            if valor_atual:
                historico_campos[nome].append(valor_atual)
                entry.delete(0, tk.END)

        botao_limpar = tk.Button(frame, text="Limpar", width=8, command=limpar_campo)
        botao_limpar.grid(row=0, column=2, padx=(6, 0))

        adicionar_campo_interface.linha_idx += 1

    adicionar_campo_interface.linha_idx = 0  # Começa depois da linha de adicionar

    def desfazer_ultima_acao(event=None):
        for nome, entry in outras_props.items():
            if nome in historico_campos and len(historico_campos[nome]) > 1:
                historico_campos[nome].pop()
                valor_anterior = historico_campos[nome][-1]
                entry.delete(0, tk.END)
                entry.insert(0, valor_anterior)
    janela.bind_all("<Control-z>", desfazer_ultima_acao)

    def adicionar_novo_campo():
        nome_novo = entrada_nome_novo.get().strip()
        valor_novo = entrada_valor_novo.get().strip()
        if not nome_novo:
            messagebox.showerror("Erro", "O nome do campo não pode estar vazio.")
            return
        if nome_novo in outras_props or nome_novo in campos_fixos:
            messagebox.showwarning("Aviso", f"O campo '{nome_novo}' já existe.")
            return

        try:
            mgr, _ = get_mgr()
            if not mgr.Get(nome_novo)[0]:
                mgr.Add3(nome_novo, 30, valor_novo, 2)
            mgr.Set(nome_novo, valor_novo)

            adicionar_campo_interface(nome_novo, valor_novo)
            entrada_nome_novo.delete(0, tk.END)
            entrada_valor_novo.delete(0, tk.END)
            entrada_nome_novo.focus_set()
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    btn_add.config(command=adicionar_novo_campo)

    # Adicionar campos existentes já cadastrados
    outros_existentes = {k: v for k, v in propriedades_existentes.items() if k not in campos_fixos}
    if outros_existentes:
        for nome, val in sorted(outros_existentes.items()):
            adicionar_campo_interface(nome, val)
    else:
        idx = adicionar_campo_interface.linha_idx
        tk.Label(frame_interno, text="Nenhuma outra propriedade encontrada.").grid(row=idx, column=0, sticky="w", padx=12, pady=(10, 0))

    # ================== Aba: Adicionar ==================
    from sw_funcs.query_edit import consultar_protheu
    from sw_funcs.nome_arq import nome_arquivo
    try:
        codigo = nome_arquivo()
        validação = consultar_protheu(codigo)
        if validação == True:
            None
        else:
            aba_adicionar = criar_aba_adicionar(notebook, janela)
    except Exception as e:
        print("Erro a gerar aba de adicionar: ", e)

    # Botão salvar
    def salvar_propriedades():
        try:
            mgr, _ = get_mgr()
            for nome, entry in entradas_fixas.items():
                valor = entry.get().strip()
                if not valor:
                    continue
                if not mgr.Get(nome)[0]:
                    mgr.Add3(nome, 30, valor, 2)
                mgr.Set(nome, valor)

            for nome, entry in outras_props.items():
                valor = entry.get().strip()
                if not valor:
                    continue
                if not mgr.Get(nome)[0]:
                    mgr.Add3(nome, 30, valor, 2)
                mgr.Set(nome, valor)

            messagebox.showinfo("Sucesso", "Propriedades salvas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    botao_salvar = tk.Button(editor, text="Salvar Propriedades", command=salvar_propriedades)
    botao_salvar.pack(pady=5)