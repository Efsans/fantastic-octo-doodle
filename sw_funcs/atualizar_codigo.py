import win32com.client
import ctypes
import os
import tkinter as tk
from tkinter import messagebox, Tk, Toplevel,ttk

def atualizar_codigo():
    try:
        swApp = win32com.client.Dispatch("SldWorks.Application")
        swModel = swApp.ActiveDoc

        if swModel is not None:
            nomeArquivo = swModel.GetPathName
            if callable(nomeArquivo):
                nomeArquivo = nomeArquivo()
                
            if nomeArquivo:
                nome_arquivo = os.path.basename(nomeArquivo)
                codigo = os.path.splitext(nome_arquivo)[0]

                swCustPropMgr = swModel.Extension.CustomPropertyManager("")
                swCustPropMgr.Add3("Codigo", 30, codigo, 2)

                if hasattr(swModel, "ForceRebuild3") and callable(swModel.ForceRebuild3):
                    swModel.ForceRebuild3(False)

                messagebox.showinfo("Sucesso", f"Código '{codigo}' salvo com sucesso!")
            else:
                messagebox.showwarning("Aviso", "O arquivo ainda não foi salvo.")
        else:
            messagebox.showwarning("Aviso", "Nenhum documento aberto no SolidWorks.")

    except Exception as e:
        messagebox.showerror("Erro", str(e))