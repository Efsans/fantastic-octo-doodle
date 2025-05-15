import tkinter as tk
from tkinter import messagebox, ttk,  StringVar, Label
from sw_funcs.gereprop import get_mgr
from sw_funcs.prefixo import prefixo
import os
from sw_funcs.validador import validar_todos_campos 


def abrir_editor(janela):
    campos_fixos = ["Codigo", "Descrição", "Tipo", "Unidade", "Pos.IPI/NCM", "Origem"]
    entradas_fixas = {}
    outras_props = {}
    salvar_adicionar = {}

    for widget in janela.pack_slaves():
        if getattr(widget, "editor_embutido", False):
            widget.destroy()

    editor = tk.Frame(janela, height=300, bg="#f0f0f0")
    editor.pack(side="bottom", fill="both", pady=5)
    editor.editor_embutido = True

    notebook = ttk.Notebook(editor)
    notebook.pack(expand=True, fill="both", padx=5, pady=5)

    
   # ================== Aba: Adicionar ==================
    aba_adicionar = tk.Frame(notebook)
    aba_adicionar.pack(fill="both")
    notebook.add(aba_adicionar, text="Adicionar")


    def salvar_em_txt():
        # Caminho para a área de trabalho
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        caminho_arquivo = os.path.join(desktop, "dados_adicionados.txt")

        with open(caminho_arquivo, "w", encoding="utf-8") as f:
            for nome, entry in salvar_adicionar.items():
                valor = entry.get() if isinstance(entry, tk.Entry) or isinstance(entry, ttk.Combobox) else ""
                f.write(f"{nome}: {valor}\n")

        print(f"Dados salvos em: {caminho_arquivo}")


    def verificar_e_habilitar_botao(event=None):
        """
        Verifica se os campos que precisam de validação são válidos 
        e habilita o botão "Adicionar" se estiverem.
        """
        campos_para_validar = {nome: entry.get() for nome, entry in salvar_adicionar.items() if nome in campos_a_validar}
        labels_status = {nome: label for nome, label in status_labels.items() if nome in campos_a_validar}

        # Valida os campos e atualiza os ícones de status
        todos_validos = True
        for nome, valor in campos_para_validar.items():
            if validar_todos_campos({nome: valor}):
                status_labels[nome].config(text="✅", fg="green")
            else:
                status_labels[nome].config(text="❌", fg="red")
                todos_validos = False

        # Habilita o botão "Adicionar" apenas se todos os campos validados forem válidos
        if todos_validos:
            botao_A["state"] = "normal"
        else:
            botao_A["state"] = "disabled"


    # Interface tkinter


    # Dicionários para armazenar widgets e status
    salvar_adicionar = {}
    status_labels = {}

    # Lista de campos
    campos_fixos = ["Codigo", "Descrição", "Tipo", "Unidade", "Pos.IPI/NCM", "Origem", "Empresa"]

    # Campos que precisam ser validados
    campos_a_validar = ["Codigo", "Descrição", "Tipo", "Unidade"]

    for nome in campos_fixos:
        frame = tk.Frame(aba_adicionar)
        frame.pack(fill="x", padx=10, pady=5)

        # Label do campo
        tk.Label(frame, text=f"{nome}:").pack(side="left", padx=(0, 10))

        # Campo de entrada ou combobox
        if nome == "Empresa":
            entry = ttk.Combobox(frame, values=["09ALFA", "01MEGA", "02AIZT"], state="normal")
            entry.set("09ALFA")  # Define um valor padrão
        else:
            entry = tk.Entry(frame)
            if nome in campos_a_validar:
                entry.bind("<KeyRelease>", verificar_e_habilitar_botao)

        entry.pack(side="left", fill="x", expand=True)
        salvar_adicionar[nome] = entry

        # Adicionar status visual (✅ ou ❌) apenas para campos validados
        if nome in campos_a_validar:
            status_label = tk.Label(frame, text="", font=("Arial", 12))
            status_label.pack(side="left", padx=10)
            status_labels[nome] = status_label

    # Botão "Adicionar"
    botao_A = tk.Button(aba_adicionar, text="Adicionar", command=salvar_em_txt, state="disabled")
    botao_A.pack(pady=20)
            

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

    for nome in campos_fixos:
        tk.Label(aba_fixos, text=nome + ":").pack(anchor="w", padx=10, pady=(10, 0))
        entry = tk.Entry(aba_fixos)
        entry.pack(fill="x", padx=10)
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

    # Adicionar novo campo
    frame_add = tk.Frame(frame_interno)
    frame_add.pack(fill="x", pady=(5, 10), padx=10)

    tk.Label(frame_add, text="Nome:").grid(row=0, column=0, sticky="w")
    entrada_nome_novo = tk.Entry(frame_add)
    entrada_nome_novo.grid(row=0, column=1, sticky="ew", padx=5)

    tk.Label(frame_add, text="Valor:").grid(row=0, column=2, sticky="w")
    entrada_valor_novo = tk.Entry(frame_add)
    entrada_valor_novo.grid(row=0, column=3, sticky="ew", padx=5)

    frame_add.columnconfigure(1, weight=1)
    frame_add.columnconfigure(3, weight=1)

    def remover_campo(nome, frame):
        try:
            mgr, _ = get_mgr()
            mgr.Delete(nome)
            frame.destroy()
            outras_props.pop(nome, None)
        except Exception as e:
            messagebox.showerror("Erro ao excluir", str(e))

    def adicionar_campo_interface(nome, valor):
        frame_linha = tk.Frame(frame_interno)
        frame_linha.pack(fill="x", padx=10, pady=3)

        tk.Label(frame_linha, text=nome + ":").pack(side="left", padx=(0, 10))
        entry = tk.Entry(frame_linha)
        entry.pack(side="left", fill="x", expand=True)
        entry.insert(0, valor)
        outras_props[nome] = entry

        botao_excluir = tk.Button(frame_linha, text="Excluir", command=lambda: remover_campo(nome, frame_linha))
        botao_excluir.pack(side="right", padx=(10, 0))

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

    btn_add = tk.Button(frame_add, text="Adicionar Campo", command=adicionar_novo_campo)
    btn_add.grid(row=0, column=4, padx=(10, 0))

    # Adicionar campos existentes
    outros_existentes = {k: v for k, v in propriedades_existentes.items() if k not in campos_fixos}
    if outros_existentes:
        for nome, val in sorted(outros_existentes.items()):
            adicionar_campo_interface(nome, val)
    else:
        tk.Label(frame_interno, text="Nenhuma outra propriedade encontrada.").pack(pady=(10, 0))

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
