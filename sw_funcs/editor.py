import win32com.client
import tkinter as tk
from tkinter import messagebox, ttk

def abrir_editor(janela):
    campos_fixos = ["Codigo", "Descrição", "Tipo", "Unidade", "Pos.IPI/NCM", "Origem"]
    entradas_fixas = {}
    outras_props = {}

    for widget in janela.pack_slaves():
        if getattr(widget, "editor_embutido", False):
            widget.destroy()  # Evita múltiplas abas

    # Cria o frame que será o painel do editor, dentro da janela principal
    editor = tk.Frame(janela, height=300, bg="#f0f0f0")
    editor.pack(side="bottom", fill="both", pady=5)
    editor.editor_embutido = True  # Marca esse frame para fácil remoção futura

    notebook = ttk.Notebook(editor)
    notebook.pack(expand=True, fill="both", padx=5, pady=5)

    aba_fixos = tk.Frame(notebook)
    notebook.add(aba_fixos, text="Campos fixos")

    try:
        # Tenta acessar o SolidWorks e carregar as propriedades
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

    outros_existentes = {k: v for k, v in propriedades_existentes.items() if k not in campos_fixos}
    if outros_existentes:
        for nome, val in outros_existentes.items():
            tk.Label(frame_interno, text=nome + ":").pack(anchor="w", padx=25, pady=(10, 0))
            entry = tk.Entry(frame_interno)
            entry.pack(fill="x", padx=10)
            entry.insert(0, val)
            outras_props[nome] = entry
    else:
        tk.Label(frame_interno, text="Nenhuma outra propriedade encontrada.").pack(pady=20)

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
