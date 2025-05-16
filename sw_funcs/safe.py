import win32com.client
import tkinter as tk
from tkinter import messagebox, ttk
from sw_funcs.editor import outras_props, entradas_fixas
from sw_funcs.validador import validar_todos_campos

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


