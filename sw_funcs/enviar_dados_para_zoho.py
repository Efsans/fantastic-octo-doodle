import win32com.client
import ctypes
import os
import tkinter as tk
from tkinter import messagebox, Tk, Toplevel,ttk
import sys
import json
import requests


def enviar_dados_para_zoho():
    try:
        # Inicializa o SolidWorks
        swApp = win32com.client.Dispatch("SldWorks.Application")
        swModel = swApp.ActiveDoc

        if swModel is None:
            messagebox.showwarning("Aviso", "Nenhum documento ativo no SolidWorks.")
            return

        # Obtém o gerenciador de propriedades personalizadas
        swCustPropMgr = swModel.Extension.CustomPropertyManager("")

        # Obtém os nomes das propriedades personalizadas
        propNames = swCustPropMgr.GetNames
        propNames = list(propNames) if propNames else []

        if not propNames:
            messagebox.showwarning("Aviso", "Nenhuma propriedade encontrada!")
            return

        # Cria o dicionário com os dados
        dados = {}
        for propName in propNames:
            # A função Get2 retorna dois valores: o valor e o valor resolvido
            resultado = swCustPropMgr.Get(propName)
            if isinstance(resultado, tuple) and len(resultado) == 2:
                 resolvedVal = resultado
            else:
                 resolvedVal = ""

            dados[propName] = resolvedVal

        # Converte os dados para JSON
        json_data = json.dumps({"dados": dados})

        # URL da API do Zoho
        api_url = "https://www.zohoapis.com/creator/custom/grupoaiz/SolidWorks?publickey=4WTWAfSnDWdjzatDCYr6gyJ4B"

        # Envia o JSON para a API
        headers = {"Content-Type": "application/json"}
        response = requests.post(api_url, data=json_data, headers=headers)

        if response.status_code == 200:
            messagebox.showinfo("Sucesso", "Enviado com sucesso!")
        else:
            messagebox.showwarning("Aviso", f"Erro ao enviar: {response.status_code}")
            messagebox.showwarning("Aviso", response.text)
    
    except Exception as e:
        messagebox.showerror("Erro", str(e))