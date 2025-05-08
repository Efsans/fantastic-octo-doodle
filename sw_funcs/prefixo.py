from parametros import params
from tkinter import messagebox
from descri_tipo import tipo_prefix

def prefixo(valor, mostrar=True):
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
        return msg

    if prefixo_extraido in lista_prefixos:
        msg = f"✅ Prefixo encontrado: {prefixo_extraido} — {tipo_prefix(prefixo_extraido)}"
        
        if mostrar:
            messagebox.showinfo("Sucesso", msg)
        return msg
    else:
        msg = f"❌ Prefixo {prefixo_extraido} não encontrado."
        if mostrar:
            messagebox.showinfo("Fracasso", msg)
        return msg

    

