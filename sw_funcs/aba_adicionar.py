import os
import tkinter as tk
from tkinter import ttk, messagebox
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
    aba = tk.Frame(notebook)
    notebook.add(aba, text="Adicionar")

    salvar_adicionar = {}
    status_labels = {}

    def limitar_codigo(novo_texto):
        return len(novo_texto) <= 8
    vcmd = janela.register(limitar_codigo)

    def on_codigo_change(event=None):
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
                lbl.config(text="✅" if valido else "❌", fg="green" if valido else "red")
            else:
                lbl.config(text="")
        botao_A["state"] = "normal" if valido else "disabled"

    def salvar_em_txt():
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        path = os.path.join(desktop, "dados_adicionados.txt")
        with open(path, "w", encoding="utf-8") as f:
            for nome, w in salvar_adicionar.items():
                f.write(f"{nome}: {w.get()}\n")
        print("Dados salvos em:", path)

    # --- Criação dos campos usando grid para alinhamento perfeito ---
    campos_lista = ["Codigo", "Descrição", "Tipo", "Unidade", "NCM", "Origem", "Armazem"]
    for idx, nome in enumerate(campos_lista):
        # Label alinhado à direita
        lbl = tk.Label(aba, text=f"{nome}:", anchor="e", width=10)
        lbl.grid(row=idx, column=0, sticky="e", padx=(12, 8), pady=4)

        if nome in ("Tipo", "Origem", "Armazem"):
            widget = ttk.Combobox(aba, state="readonly")
            widget.grid(row=idx, column=1, sticky="ew", padx=(0, 4), pady=4)
            widget.bind("<<ComboboxSelected>>", on_validar)
            salvar_adicionar[nome] = widget

        elif nome == "NCM":
            cb_ncm = ttk.Combobox(aba, state="normal", width=6)
            cb_ncm.grid(row=idx, column=1, sticky="ew", padx=(0, 4), pady=4)

            def on_buscar_ncm():
                descricao = salvar_adicionar["Descrição"].get().strip()
                if len(descricao) < 1:
                    messagebox.showwarning("Aviso", "Descrição muito curta para buscar NCM.")
                    return
                ncms = obter_ncm_sugerido(descricao)
                if not ncms:
                    messagebox.showinfo("Info", "Nenhum NCM encontrado para esta descrição.")
                    return
                cb_ncm["values"] = ncms
                cb_ncm.set(ncms[0])

            btn_buscar = tk.Button(aba, text="Consultar IA", width=10, command=on_buscar_ncm)
            btn_buscar.grid(row=idx, column=2, padx=(0, 4), pady=4)
            salvar_adicionar[nome] = cb_ncm

        else:
            italic = ("Arial", 10, "italic") if nome in ("Descrição", "Unidade") else None
            opts = {}
            if nome == "Codigo":
                opts["validate"] = "key"
                opts["validatecommand"] = (vcmd, "%P")
            widget = tk.Entry(aba, font=italic, **opts)
            widget.grid(row=idx, column=1, sticky="ew", padx=(0, 4), pady=4)
            widget.bind("<Return>", on_codigo_change if nome == "Codigo" else on_validar)
            widget.bind("<KeyRelease>", on_codigo_change if nome == "Codigo" else on_validar)
            salvar_adicionar[nome] = widget

        # Status label (check/erro) ao lado do campo
        if nome in ("Codigo", "Descrição", "Unidade"):
            lbl_status = tk.Label(aba, text="", font=("Arial", 12, "italic"))
            lbl_status.grid(row=idx, column=3, sticky="w", padx=(4, 0))
            status_labels[nome] = lbl_status

    aba.columnconfigure(1, weight=1)  # Faz o campo expandir na horizontal

    botao_A = tk.Button(aba, text="Adicionar", command=salvar_em_txt, state="disabled")
    botao_A.grid(row=len(campos_lista), column=0, columnspan=4, pady=16)
    
    import requests

    def obter_ncm_sugerido(descricao):
        try:
            url = "http://localhost:8000/classificar-ncm"
            response = requests.post(url, json={"descricao": descricao}, timeout=5)
            if response.status_code == 200:
                data = response.json()
                ncms = data.get("ncms", [])
                
                if not ncms:
                    messagebox.showinfo("Info", "Nenhum NCM encontrado para esta descrição.")
                return ncms
            else:
                messagebox.showerror("Erro", f"Erro ao buscar NCM: {response.status_code}\n{response.text}")
                return []
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao buscar NCM: {e}")
            return []

    on_codigo_change()  # validação inicial
    return aba, salvar_adicionar, status_labels, botao_A