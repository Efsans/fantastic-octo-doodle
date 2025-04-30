import win32com.client
import ctypes
import os
import tkinter as tk
from tkinter import messagebox, Tk, Toplevel,ttk
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

####################################################
#              tabela de resulatdo                 #
####################################################
def gerar_lista(componentes, nivel, numero=0, lista=[], listadebug=[]):
    for comp in componentes:
        numero += 1
        nome_comp = extrair_codigo_componente(comp.Name2)
        listadebug.append(f"numero: {numero} nivel: {nivel}")
        listadebug.append(f"{obter_nome_por_nivel(nivel)}: {nome_comp}")
        lista.append(nome_comp)
        subcomps = comp.GetChildren
        if subcomps:
            # Chamada recursiva para subcomponentes
            gerar_lista(subcomps, nivel + 1, numero, lista, listadebug)
    return lista

def consultar_lista(lista):
    conexao = pyodbc.connect(get())
    cursor = conexao.cursor()

    try:
        placeholders = ", ".join(["?"] * len(lista))
        query = f"""
            SELECT B1_COD, B1_DESC, B1_FILIAL 
            FROM SB1010 AS SB1 
            WHERE SB1.D_E_L_E_T_ = '' 
              AND SB1.B1_COD IN ({placeholders}) 
              AND SB1.B1_FILIAL = '09ALFA'
        """
        cursor.execute(query, lista)
        resultados = cursor.fetchall()
        return resultados

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao consultar a lista: {str(e)}")
        return []

    finally:
        cursor.close()
        conexao.close()



def exibir_tela(resultados, json_dados):
    # Janela principal
    janela = Toplevel()
    janela.title("Resultados da Consulta")
    janela.geometry("600x400")

    # Configurar Treeview
    colunas = ("Código", "Descrição", "Filial")
    tree = ttk.Treeview(janela, columns=colunas, show="headings", height=15)
    tree.heading("Código", text="Código")
    tree.heading("Descrição", text="Descrição")
    tree.heading("Filial", text="Filial")
    tree.column("Código", width=150)
    tree.column("Descrição", width=300)
    tree.column("Filial", width=100)

    # Adicionar dados ao Treeview
    for row in resultados:
        tree.insert("", "end", values=(row[0], row[1], row[2]))

    tree.pack(fill="both", expand=True)

    # Botão
    botao_enviar = tk.Button(janela, text="enviar", command=lambda:enviar_api(json_dados), height=2, width=25)
    botao_enviar.pack(pady=5) 
    botao_fechar = tk.Button(janela, text="fechar", command=janela.destroy, height=2, width=25)
    botao_fechar.pack(pady=5)
    

    

def extrair_codigo_arquivo(nome_arquivo):
    return os.path.splitext(os.path.basename(nome_arquivo))[0]

def obter_nome_por_nivel(nivel):
    return {
        1: "Modulo",
        2: "Conjunto",
        3: "Componente",
        4: "Subcomponente"
    }.get(nivel, f"Nivel{nivel}")

def extrair_codigo_componente(nome_completo):
    nome = nome_completo.split("/")[-1]
    return nome.split("-")[0]

def eh_codigo_padrao(codigo):
    return " " not in codigo and 1 <= len(codigo) <= 9

def gerar_json(comps, nivel):
    filhos = []
    contador = 0

    for comp in comps:
        nome_comp = extrair_codigo_componente(comp.Name2)
        if eh_codigo_padrao(nome_comp):
            contador += 1
            item = {
                obter_nome_por_nivel(nivel): nome_comp,
                "Numero": contador
            }
            subcomps = comp.GetChildren
            if subcomps:
                # Chamada recursiva para subcomponentes
                item["Filhos"] = json.loads(gerar_json(subcomps, nivel + 1))["Filhos"]
            filhos.append(item)

    return json.dumps({"Filhos": filhos})

def tabela_dados():
    root = Tk()
    root.withdraw()

    try:
        swApp = win32com.client.Dispatch("SldWorks.Application")
        swModel = swApp.ActiveDoc

        if swModel is None or swModel.GetType != 2:
            messagebox.showerror("Erro", "Nenhum assembly aberto no SolidWorks.")
            return

        caminho = swModel.GetPathName
        if not caminho:
            messagebox.showerror("Erro", "O documento não foi salvo. Salve o arquivo antes de continuar.")
            return

        codigo = extrair_codigo_arquivo(caminho)
        componentes = swModel.GetComponents(False)

        lista = gerar_lista(componentes, 1)
        resultados = consultar_lista(lista)

        json_dados = {
            "dados": {
                "Produto": codigo,
                **json.loads(gerar_json(componentes, 1))
            }
        }

        if resultados:
            exibir_tela(resultados, json_dados)
        else:
            messagebox.showinfo("Sem Resultados", "Nenhum dado encontrado para os componentes fornecidos.")

    except Exception as e:
        messagebox.showerror("Erro", str(e))


def enviar_api(json_dados):
    try:
        json_string = json.dumps(json_dados, indent=2)
        url = "https://www.zohoapis.com/creator/custom/grupoaiz/SolidWorks?publickey=4WTWAfSnDWdjzatDCYr6gyJ4B"
        headers = {"Content-Type": "application/json"}
        response = requests.post(url, headers=headers, data=json_string)

        if response.status_code == 200:
            messagebox.showinfo("Sucesso", "Dados enviados com sucesso!")
        else:
            messagebox.showerror("Erro", f"Erro ao enviar: {response.status_code}\n{response.text}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao enviar para API: {str(e)}")

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
janela.geometry("320x480")

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

botao_tabelar = tk.Button(janela, text="tabela de Propriedades", command=tabela_dados, height=2, width=25)
botao_tabelar.pack(pady=5)


janela.mainloop()
