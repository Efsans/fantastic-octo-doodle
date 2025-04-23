import win32com.client
import ctypes
import os
import tkinter as tk
from tkinter import messagebox, ttk
import sys
import json
import requests
import pyodbc

def get():
  # Função que retorna a string de conexão com o banco de dados SQL Server
  dados_conexao = ("Driver={SQL Server};"
                  "Server=TOTVSAPL;"
                  "Database=protheus12_producao;"
                  "UID=consulta;"
                  "PWD=consulta;")
  return dados_conexao
# conexao = pyodbc.connect(get())      #
#   cursor = conexao.cursor()

##########################################################################################################
#////////////////////////////////////////////////////////////////////////////////////////////////////////#
##########################################################################################################

############################################
#     Alto preencher campos               #
############################################
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
            pegar_nome_do_arquivo()
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


###
#PEGADOR DE ARQUIVO
##
def pegar_nome_do_arquivo():
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
                monitorar_propriedade()
            else:
                messagebox.showwarning("Aviso", "O arquivo ainda não foi salvo.")
        else:
            messagebox.showwarning("Aviso", "Nenhum documento aberto no SolidWorks.")

    except Exception as e:
        messagebox.showerror("Erro", str(e))

 

############################################
#     Consultar Dados Zoho               #
############################################   

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


############################################
#     Atualizar Código no SolidWorks       #
############################################

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


##################################
# === Interface Gráfica ===      #   
##################################

def abrir_editor():
    def carregar_valor(event=None):
        nome = combo_props.get()
        if nome in propriedades_existentes:
            entry_valor.delete(0, tk.END)
            entry_valor.insert(0, propriedades_existentes[nome])
        else:
            entry_valor.delete(0, tk.END)

    def salvar_propriedade():
        nome = combo_props.get().strip()
        valor = entry_valor.get().strip()

        if not nome:
            messagebox.showwarning("Aviso", "O nome da propriedade não pode estar vazio!")
            return

        try:
            swApp = win32com.client.Dispatch("SldWorks.Application")
            swModel = swApp.ActiveDoc
            if swModel is None:
                messagebox.showerror("Erro", "Nenhum documento aberto no SolidWorks!")
                return
            swCustPropMgr = swModel.Extension.CustomPropertyManager("")
            resultado = swCustPropMgr.Add3(nome, 30, valor, 2)
            if resultado == 0:
                swCustPropMgr.Set(nome, valor)
            propriedades_existentes[nome] = valor
            combo_props["values"] = list(propriedades_existentes.keys())
            messagebox.showinfo("Sucesso", f"Propriedade '{nome}' salva com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    try:
        swApp = win32com.client.Dispatch("SldWorks.Application")
        swModel = swApp.ActiveDoc
        swCustPropMgr = swModel.Extension.CustomPropertyManager("")
        nomes = swCustPropMgr.GetNames
        propriedades_existentes = {}
        if nomes:
            for nome in nomes:
                val = swCustPropMgr.Get(nome)
                if isinstance(val, tuple):
                    val = val[0]
                propriedades_existentes[nome] = val
    except:
        propriedades_existentes = {}

    editor = tk.Toplevel(janela)
    editor.title("Editar Propriedade")
    editor.geometry("320x160")
    editor.attributes("-topmost", True)

    tk.Label(editor, text="Nome da Propriedade:").pack(pady=(10,0))
    combo_props = ttk.Combobox(editor, values=list(propriedades_existentes.keys()))
    combo_props.pack(pady=5)
    combo_props.bind("<KeyRelease>", carregar_valor)
    combo_props.bind("<<ComboboxSelected>>", carregar_valor)

    tk.Label(editor, text="Valor:",width=150).pack()
    entry_valor = tk.Entry(editor)
    entry_valor.pack(pady=5)

    botao_salvar = tk.Button(editor, text="Salvar", command=salvar_propriedade)
    botao_salvar.pack(pady=10)

janela = tk.Tk()
janela.title("SolidWorks - Ferramentas")
janela.attributes("-topmost", True)
janela.geometry("320x280")

label = tk.Label(janela, text="Escolha uma ação:")
label.pack(pady=15)

botao_atualizar = tk.Button(janela, text="Atualizar Código", command=atualizar_codigo, height=2, width=25)
botao_atualizar.pack(pady=5)

botao_processar = tk.Button(janela, text="Enviar dados para Zoho", command=enviar_dados_para_zoho, height=2, width=25)
botao_processar.pack(pady=5)

botao_consultar = tk.Button(janela, text="Consultar e Atualizar", command=monitorar_propriedade, height=2, width=25)
botao_consultar.pack(pady=5)

botao_editar = tk.Button(janela, text="Editar Propriedade", command=abrir_editor, height=2, width=25)
botao_editar.pack(pady=5)

janela.mainloop()
