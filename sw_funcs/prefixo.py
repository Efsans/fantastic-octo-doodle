from tkinter import messagebox
from sw_funcs.descri_tipo import tipo_prefix
from sw_funcs.parametros import params


def prefixo(valor, mostrar=True):
    liber = None
    valor = valor.strip().upper()
    if valor.startswith("C3S"):
        prefixo_extraido = "C3S"
    elif valor.startswith("Z"):
        prefixo_extraido = "Z"
    else:
        prefixo_extraido = valor[:2]

    prefixos = params(p='SOL_MP')
    if prefixos:
        lista_prefixos = prefixos[0][0].split(',')
    else:
        msg = "❌ Banco de dados não acessado."
        if mostrar:
            messagebox.showinfo("Erro", msg)
        return msg, False

    if prefixo_extraido in lista_prefixos:
        msg = f"✅ Prefixo encontrado: {prefixo_extraido} — {tipo_prefix(prefixo_extraido)}"
        if mostrar:
            messagebox.showinfo("Sucesso", msg)
        return msg, True
    else:
        msg = f"❌ Prefixo {prefixo_extraido} não encontrado."
        if mostrar:
            messagebox.showinfo("Fracasso", msg)
        return msg, False

    

