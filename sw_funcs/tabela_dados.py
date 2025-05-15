import win32com.client
import ctypes
import os
import tkinter as tk
from tkinter import messagebox, Tk, Toplevel, ttk
import sys
import json
import requests
import pyodbc
from sw_funcs.conectsql import get

# Cache global
resultados_cache = None
json_cache = None



def gerar_lista(componentes, nivel, numero=0, lista=[], listadebug=[]):
    for comp in componentes:
        numero += 1
        nome_comp = extrair_codigo_componente(comp.Name2)
        listadebug.append(f"numero: {numero} nivel: {nivel}")
        listadebug.append(f"{obter_nome_por_nivel(nivel)}: {nome_comp}")
        lista.append(nome_comp)
        subcomps = comp.GetChildren
        if subcomps:
            gerar_lista(subcomps, nivel + 1, numero, lista, listadebug)
    return lista

def consultar_lista(lista):
    conexao = pyodbc.connect(get())
    cursor = conexao.cursor()
    try:
        placeholders = ", ".join(["?"] * len(lista))
        query = f"""
            SELECT B1_COD, B1_DESC, B1_TIPO, B1_UM, B1_POSIPI, B1_ORIGEM, B1_FILIAL 
            FROM SB1010 AS SB1 
            WHERE SB1.D_E_L_E_T_ = '' 
              AND SB1.B1_COD IN ({placeholders}) 
              AND SB1.B1_FILIAL = '09ALFA'
        """
        cursor.execute(query, lista)
        resultados = cursor.fetchall()

        # Limpa cada campo de cada linha
        resultados_limpos = []
        for row in resultados:
            row_limpo = tuple(
                str(valor).replace("\n", " ").replace("\r", " ").replace(";", " ").strip() if valor is not None else ""
                for valor in row
            )
            resultados_limpos.append(row_limpo)

        return resultados_limpos
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao consultar a lista: {str(e)}")
        return []
    finally:
        cursor.close()
        conexao.close()


def exibir_tela(janela, resultados, json_dados):
    for widget in janela.pack_slaves():
        if getattr(widget, "editor_embutido", False):
            widget.destroy()

    frame_resultado = tk.Frame(janela, height=300, bg="#f0f0f0")
    frame_resultado.pack(side="bottom", fill="both", pady=5)
    frame_resultado.editor_embutido = True

    colunas = ("Código", "Descrição", "Pos.IPI/NCM", "Origem", "Filial", "Unidade")
    tree = ttk.Treeview(frame_resultado, columns=colunas, show="headings", height=10)
    tree.pack(fill="both", expand=True, padx=10, pady=(10, 5))

    for col in colunas:
        tree.heading(col, text=col)
        if col == "Descrição":
            tree.column(col, width=300)
        elif col == "Código":
            tree.column(col, width=150)
        else:
            tree.column(col, width=120)

    for row in resultados:
        valores = (
            row[0],  # Código
            row[1],  # Descrição
            row[4],  # Pos.IPI/NCM
            row[5],  # Origem
            row[6],  # Filial
            row[3],  # Unidade
        )
        tree.insert("", "end", values=valores)

    # Função para editar célula com clique duplo
    def editar_celula(event):
        item_id = tree.identify_row(event.y)
        coluna = tree.identify_column(event.x)
        col_index = int(coluna[1:]) - 1

        if not item_id or col_index == 0:
            return  # Não edita coluna "Código"

        x, y, width, height = tree.bbox(item_id, column=coluna)
        valor_atual = tree.item(item_id, "values")[col_index]

        entry = tk.Entry(tree)
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, valor_atual)
        entry.focus()

        def salvar_alteracao(event):
            novo_valor = entry.get()
            valores = list(tree.item(item_id, "values"))
            valores[col_index] = novo_valor
            tree.item(item_id, values=valores)
            entry.destroy()

        entry.bind("<Return>", salvar_alteracao)
        entry.bind("<FocusOut>", lambda e: entry.destroy())

    tree.bind("<Double-1>", editar_celula)

    # Função para salvar propriedades nos arquivos dos componentes
    def salvar_propriedades():
        try:
            swApp = win32com.client.Dispatch("SldWorks.Application")
            swModel = swApp.ActiveDoc

            messagebox.showinfo("Debug", "Obtendo componentes iniciais do modelo ativo...")
            componentes_iniciais = swModel.GetComponents(False)
            messagebox.showinfo("Debug", f"Tipo de componentes_iniciais: {type(componentes_iniciais)}")

            if isinstance(componentes_iniciais, tuple):
                messagebox.showinfo("Debug", "Convertendo tupla para lista")
                componentes_iniciais = list(componentes_iniciais)

            todos_componentes = []
            fila = list(componentes_iniciais)
            messagebox.showinfo("Debug", f"Componentes iniciais na fila: {len(fila)}")

            while fila:
                comp = fila.pop(0)
                todos_componentes.append(comp)

                try:
                    filhos = comp.GetChildren()
                    if filhos:
                        messagebox.showinfo("Debug", f"Componente '{comp.Name2}' tem {len(filhos)} filhos")
                        fila.extend(list(filhos))
                    else:
                        messagebox.showinfo("Debug", f"Componente '{comp.Name2}' não tem filhos")
                except Exception as e:
                    messagebox.showinfo("Debug", f"Erro ao obter filhos de '{getattr(comp, 'Name2', 'sem nome')}': {e}")
                    continue

            messagebox.showinfo("Debug", f"Total de componentes coletados: {len(todos_componentes)}")

            cod_para_comp = {
                extrair_codigo_componente(comp.Name2): comp
                for comp in todos_componentes
                if hasattr(comp, "Name2")
            }
            messagebox.showinfo("Debug", f"Total de códigos mapeados: {len(cod_para_comp)}")

            erros = []

            for item in tree.get_children():
                valores = tree.item(item)["values"]
                codigo = valores[0]
                comp = cod_para_comp.get(codigo)

                if not comp:
                    erros.append(f"Componente não encontrado: {codigo}")
                    continue

                try:
                    suprimido = comp.IsSuppressed
                    messagebox.showinfo("Debug", f"Código {codigo} -> IsSuppressed: {suprimido}")
                    if suprimido:
                        erros.append(f"Componente suprimido: {codigo}")
                        continue
                except Exception as e:
                    messagebox.showinfo("Debug", f"Erro ao verificar IsSuppressed para {codigo}: {e}")

                model_path = ""
                modelo = None

                try:
                    model_path = comp.GetPathName()
                    messagebox.showinfo("Debug", f"Código {codigo} -> Caminho do modelo: {model_path}")
                    modelo = comp.GetModelDoc2()
                except Exception as e:
                    messagebox.showinfo("Debug", f"Erro ao obter caminho/modelo para {codigo}: {e}")

                if modelo is None and model_path and os.path.exists(model_path):
                    try:
                        tipo_doc = 1 if model_path.lower().endswith(".sldprt") else 2
                        modelo = swApp.OpenDoc6(model_path, tipo_doc, 0, "", 0, 0)
                        messagebox.showinfo("Debug", f"Código {codigo} -> Documento aberto com sucesso")
                    except Exception as e:
                        erros.append(f"Erro ao abrir o modelo de {codigo}: {e}")
                        continue

                if modelo:
                    try:
                        props = modelo.Extension.CustomPropertyManager("")
                        props.Set2("Descricao", valores[1])
                        props.Set2("NCM", valores[2])
                        props.Set2("Origem", valores[3])
                        props.Set2("Filial", valores[4])
                        props.Set2("Unidade", valores[5])
                        messagebox.showinfo("Debug", f"Código {codigo} -> Propriedades salvas com sucesso")
                    except Exception as e:
                        erros.append(f"Erro ao definir propriedades em {codigo}: {e}")
                else:
                    erros.append(f"Modelo não carregado: {codigo}")

            if erros:
                messagebox.showwarning("Avisos", "\n".join(erros))
            else:
                messagebox.showinfo("Sucesso", "Propriedades salvas com sucesso.")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar propriedades: {e}")





    botoes_frame = tk.Frame(frame_resultado, bg="#f0f0f0")
    botoes_frame.pack(pady=5)

    tk.Button(botoes_frame, text="Salvar Propriedades", command=salvar_propriedades, height=2, width=25).pack(side="left", padx=5)
    tk.Button(botoes_frame, text="Enviar", command=lambda: enviar_api(json_dados), height=2, width=25).pack(side="left", padx=5)
    tk.Button(botoes_frame, text="Fechar", command=frame_resultado.destroy, height=2, width=25).pack(side="left", padx=5)

    

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
                item["Filhos"] = json.loads(gerar_json(subcomps, nivel + 1))["Filhos"]
            filhos.append(item)
    return json.dumps({"Filhos": filhos})

def tabela_dados(janela_principal):
    global resultados_cache, json_cache

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

        if resultados_cache and json_cache:
            exibir_tela(janela_principal, resultados_cache, json_cache)
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
            resultados_cache = resultados
            json_cache = json_dados
            exibir_tela(janela_principal, resultados, json_dados)
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
