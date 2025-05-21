# aba_adicionar.py
import os
import tkinter as tk
from tkinter import ttk
from sw_funcs.descri_tipo import regras
from sw_funcs.validador import validar_todos_campos


def extrair_prefixo(codigo):
    codigo = codigo.strip().upper()
    p = codigo[:2]
    if any(p in [pref.strip().upper() for pref in r["prefixos"]] for r in regras()):
        return p
    return None


def obter_regras(prefixo):
    return [r for r in regras() if prefixo in [pref.strip().upper() for pref in r["prefixos"]]]


def criar_aba_adicionar(notebook, janela):
    """
    Cria a aba 'Adicionar' em um ttk.Notebook já existente.
    Retorna (frame_da_aba, salvar_adicionar_dict, status_labels_dict, botao_Adicionar).
    """
    aba = tk.Frame(notebook)
    notebook.add(aba, text="Adicionar")

    salvar_adicionar = {}
    status_labels   = {}

    # validação de comprimento do código
    def limitar_codigo(novo_texto):
        return len(novo_texto) <= 8
    vcmd = janela.register(limitar_codigo)

    def on_codigo_change(event=None):
        # atualiza comboboxes sempre que o código muda
        cod = salvar_adicionar["Codigo"].get().strip().upper()
        pref = extrair_prefixo(cod) if cod else None
        regras_list = obter_regras(pref) if pref else []

        # Tipo
        tipos = sorted({r["tipo"] for r in regras_list})
        cb_tipo = salvar_adicionar["Tipo"]
        cb_tipo["values"] = tipos
        if tipos and cb_tipo.get() not in tipos:
            cb_tipo.set(tipos[0])

        # Armazém
        arm_vals = sorted({r["armazem"] for r in regras_list if r.get("armazem") and r["armazem"] != "A Definir"})
        cb_arm = salvar_adicionar["Armazem"]
        cb_arm["values"] = arm_vals
        if arm_vals and cb_arm.get() not in arm_vals:
            cb_arm.set(arm_vals[0])

        # Origem
        orig_vals = sorted({o for r in regras_list for o in r.get("origens", [])})
        cb_ori = salvar_adicionar["Origem"]
        cb_ori["values"] = orig_vals
        if orig_vals and cb_ori.get() not in orig_vals:
            cb_ori.set(orig_vals[0])

        on_validar()

    def on_validar(event=None):
        campos = {n: w.get().strip() for n, w in salvar_adicionar.items()}
        valido = validar_todos_campos(campos)
        for nome, lbl in status_labels.items():
            if campos[nome]:
                lbl.config(text="✅" if valido else "❌",
                           fg="green" if valido else "red")
            else:
                lbl.config(text="")
        botao_A["state"] = "normal" if valido else "disabled"

    def salvar_em_txt():
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        path    = os.path.join(desktop, "dados_adicionados.txt")
        with open(path, "w", encoding="utf-8") as f:
            for nome, w in salvar_adicionar.items():
                f.write(f"{nome}: {w.get()}\n")
        print("Dados salvos em:", path)

    campos_lista = ["Codigo", "Descrição", "Tipo", "Unidade", "NCM", "Origem", "Armazem"]
    for nome in campos_lista:
        frm = tk.Frame(aba)
        frm.pack(fill="x", padx=10, pady=5)
        tk.Label(frm, text=f"{nome}:").pack(side="left", padx=(0,10))

        if nome in ("Tipo", "Origem", "Armazem"):
            widget = ttk.Combobox(frm, state="readonly")
            widget.bind("<<ComboboxSelected>>", on_validar)
        else:
            italic = ("Arial", 10, "italic") if nome in ("Descrição", "Unidade") else None
            opts = {}
            if nome == "Codigo":
                opts["validate"] = "key"
                opts["validatecommand"] = (vcmd, "%P")
            widget = tk.Entry(frm, font=italic, **opts)
            # bind Enter e KeyRelease para validação
            widget.bind("<Return>", on_codigo_change if nome == "Codigo" else on_validar)
            widget.bind("<KeyRelease>", on_codigo_change if nome == "Codigo" else on_validar)

        widget.pack(side="left", fill="x", expand=True)
        salvar_adicionar[nome] = widget

        if nome in ("Codigo", "Descrição", "Unidade"):
            lbl = tk.Label(frm, text="", font=("Arial", 12, "italic"))
            lbl.pack(side="left", padx=10)
            status_labels[nome] = lbl

    botao_A = tk.Button(aba, text="Adicionar", command=salvar_em_txt, state="disabled")
    botao_A.pack(pady=20)

    # validação inicial
    on_codigo_change()
    return aba, salvar_adicionar, status_labels, botao_A