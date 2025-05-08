from prefixo import prefixo
import tkinter as tk

def monitorar_entrada(nome, entry, parent_frame):
    if nome != "Codigo":
        return  # só monitora o campo "Codigo"

    # Label para mostrar o resultado abaixo do campo
    resultado_label = tk.Label(parent_frame, text="", fg="blue", wraplength=300, justify="left")
    resultado_label.pack(after=entry, anchor="w", padx=5, pady=(0, 5))

    def ao_apertar_enter(event=None):
        valor = entry.get().strip()
        msg = prefixo(valor, mostrar=False)
        resultado_label.config(text=msg)

    entry.bind("<Return>", ao_apertar_enter)