import win32com.client
import tkinter as tk
from tkinter import messagebox, ttk


def abrir_editor(janela):
    campos_fixos = ["Codigo", "Descrição", "Tipo", "Unidade", "Pos.IPI/NCM", "Origem"]
    entradas_fixas = {}
    outras_props = {}

    for widget in janela.pack_slaves():
        if getattr(widget, "editor_embutido", False):
            widget.destroy()

    editor = tk.Frame(janela, height=300, bg="#f0f0f0")
    editor.pack(side="bottom", fill="both", pady=5)
    editor.editor_embutido = True

    notebook = ttk.Notebook(editor)
    notebook.pack(expand=True, fill="both", padx=5, pady=5)

    aba_fixos = tk.Frame(notebook)
    notebook.add(aba_fixos, text="Campos fixos")

    try:
        swApp = win32com.client.Dispatch("SldWorks.Application")
        swModel = swApp.ActiveDoc
        swCustPropMgr = swModel.Extension.CustomPropertyManager("")
        nomes = swCustPropMgr.GetNames
        propriedades_existentes = {
            nome: swCustPropMgr.Get(nome)[0] if isinstance(swCustPropMgr.Get(nome), tuple) else swCustPropMgr.Get(nome)
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

    # Seção para adicionar novos campos
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
            swApp = win32com.client.Dispatch("SldWorks.Application")
            swModel = swApp.ActiveDoc
            swCustPropMgr = swModel.Extension.CustomPropertyManager("")
            swCustPropMgr.Delete(nome)
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
            swApp = win32com.client.Dispatch("SldWorks.Application")
            swModel = swApp.ActiveDoc
            swCustPropMgr = swModel.Extension.CustomPropertyManager("")
            swCustPropMgr.Add3(nome_novo, 30, valor_novo, 2)
            swCustPropMgr.Set(nome_novo, valor_novo)

            adicionar_campo_interface(nome_novo, valor_novo)
            entrada_nome_novo.delete(0, tk.END)
            entrada_valor_novo.delete(0, tk.END)
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    btn_add = tk.Button(frame_add, text="Adicionar Campo", command=adicionar_novo_campo)
    btn_add.grid(row=0, column=4, padx=(10, 0))

    # Adiciona os campos que já existiam (excluindo os fixos)
    outros_existentes = {k: v for k, v in propriedades_existentes.items() if k not in campos_fixos}
    if outros_existentes:
        for nome, val in outros_existentes.items():
            adicionar_campo_interface(nome, val)
    else:
        tk.Label(frame_interno, text="Nenhuma outra propriedade encontrada.").pack(pady=(10, 0))

    def salvar_propriedades():
        try:
            swApp = win32com.client.Dispatch("SldWorks.Application")
            swModel = swApp.ActiveDoc
            if swModel is None:
                messagebox.showerror("Erro", "Nenhum documento aberto no SolidWorks!")
                return
            swCustPropMgr = swModel.Extension.CustomPropertyManager("")

            for nome, entry in entradas_fixas.items():
                valor = entry.get().strip()
                if nome:
                    swCustPropMgr.Add3(nome, 30, valor, 2)
                    swCustPropMgr.Set(nome, valor)

            for nome, entry in outras_props.items():
                valor = entry.get().strip()
                if nome:
                    swCustPropMgr.Add3(nome, 30, valor, 2)
                    swCustPropMgr.Set(nome, valor)

            messagebox.showinfo("Sucesso", "Propriedades salvas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    botao_salvar = tk.Button(editor, text="Salvar Propriedades", command=salvar_propriedades)
    botao_salvar.pack(pady=5)
