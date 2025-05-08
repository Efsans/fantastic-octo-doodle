import win32com.client
import ctypes
import os
import tkinter as tk
from tkinter import messagebox, Tk, Toplevel,ttk
import sys
import json
import requests
import pyodbc
from sw_funcs.conectsql import get
from sw_funcs.atualizar_codigo import atualizar_codigo


def monitorar_propriedade():
    try:
        # Acessa o SolidWorks
        swApp = win32com.client.Dispatch("SldWorks.Application")
        swModel = swApp.ActiveDoc

        if swModel is None:
            messagebox.showerror("Erro", "Nenhum documento aberto no SolidWorks!")
            return

        # Acessa o gerenciador de propriedades
        swCustPropMgr = swModel.Extension.CustomPropertyManager("")

        # Propriedades monitoradas
        codigo_key = "Codigo"
        

        # Obtém os valores de 'Codigo' e 'Filial'
        codigo = swCustPropMgr.Get(codigo_key)
        filial = "09ALFA"

        if not codigo or codigo is None :
            messagebox.showwarning("Atenção", "Código não esta preenchidos! clique em ok para o auto preenchimento.")
            atualizar_codigo()
            return
                         
            

        conexao = pyodbc.connect(get())      
        cursor = conexao.cursor()

        query=("""
        SELECT 
        B1_COD ,
        B1_TIPO,
        B1_UM,
        RTRIM(B1_DESC) AS B1_DESC,
        B1_POSIPI,
        B1_ORIGEM,
        B1_FILIAL  
        FROM SB1010   SB1		
        


""")
        conditions = ["SB1.D_E_L_E_T_ = ''"]  # Condição fixa
        params = []

# Condições variáveis
        if filial:
            conditions.append("SB1.B1_FILIAL = ?")
            params.append(filial)
        else:
            messagebox.showwarning("Atencao", "Filial nao encontrada")
            return

        if codigo:
            conditions.append("SB1.B1_COD = ?")
            params.append(codigo)
        else:
            messagebox.showwarning("Atencao", "Codigo nao encontrado")
            return
# Concatena a cláusula WHERE somente se houver condições
        if conditions:
            query += " WHERE " + " AND ".join(conditions)  # Lista para acumular condições da cláusula WHERE
            

        cursor.execute(query, params)
        row = cursor.fetchone()
            
        if row is None:
            messagebox.showwarning("Atencao", "Código nao encontrado na base de dados!")
            messagebox.showinfo("chore", f"{row}")
            return

        cursor.close()
        conexao.close()

        
        processar_resposta(swModel, row)

    except Exception as e:
        messagebox.showerror("Erro", str(e))


def processar_resposta(swModel, row):
    if not row:
        messagebox.showwarning("Aviso", "Nenhum dado encontrado! Verifique a tabela ou os parâmetros.")
        return

    # Mapeamento do nome das colunas do banco para o nome das propriedades do SolidWorks
    mapeamento = {
        "B1_COD": "Codigo",
        "B1_TIPO": "Tipo",
        "B1_UM": "Unidade",
        "B1_LOCAL": "Arm. Empenho",
        "B1_DESC": "Descrição",
        "B1_POSIPI": "Pos.IPI/NCM",
        "B1_ORIGEM": "Origem"
    }

    swCustPropMgr = swModel.Extension.CustomPropertyManager("")
    campos_alterados = 0

    for coluna, nome_prop in mapeamento.items():
        valor = getattr(row, coluna, None)
        if valor is not None:
            valor_str = str(valor)
            try:
                # Tenta adicionar a propriedade (cria se não existir)
                resultado = swCustPropMgr.Add3(nome_prop, 30, valor_str, 2)
                if resultado == 0:
                    # Se já existe, atualiza o valor
                    swCustPropMgr.Set(nome_prop, valor_str)
                campos_alterados += 1
            except Exception as e:
                print(f"Erro ao definir a propriedade {nome_prop}: {e}")

    if campos_alterados > 0:
        messagebox.showinfo("Sucesso", f"Atualização concluída! {campos_alterados} campos alterados.")
    else:
        messagebox.showwarning("Aviso", "Nenhum campo foi atualizado.")