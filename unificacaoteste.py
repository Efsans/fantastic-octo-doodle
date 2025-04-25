import win32com.client
import os
import json
import requests
from tkinter import messagebox, Tk, Toplevel
from tkinter import ttk
import pyodbc

def get():
    # Função que retorna a string de conexão com o banco de dados SQL Server
    return ("Driver={SQL Server};"
            "Server=TOTVSAPL;"
            "Database=protheus12_producao;"
            "UID=consulta;"
            "PWD=consulta;")

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

def exibir_tela(resultados):
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

    # Botão de fechamento
    ttk.Button(janela, text="Fechar", command=janela.destroy).pack(pady=10)

    

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

def main():
    root = Tk()
    root.withdraw()

    try:
        swApp = win32com.client.Dispatch("SldWorks.Application")
        swModel = swApp.ActiveDoc

        if swModel is None or swModel.GetType != 2:  # 2 = swDocASSEMBLY
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

        if resultados:
            exibir_tela(resultados)
        else:
            messagebox.showinfo("Sem Resultados", "Nenhum dado encontrado para os componentes fornecidos.")


        
        json_dados = {
            "dados": {
                "Produto": codigo,
                **json.loads(gerar_json(componentes, 1))  # Gera estrutura hierárquica
            }
        }

        json_string = json.dumps(json_dados, indent=2)
        

        url = "https://www.zohoapis.com/creator/custom/grupoaiz/SolidWorks?publickey=4WTWAfSnDWdjzatDCYr6gyJ4"  # colocar um "B" no final
        headers = {"Content-Type": "application/json"}
        response = requests.post(url, headers=headers, data=json_string)

        if response.status_code == 200:
            messagebox.showinfo("Sucesso", "Dados enviados com sucesso!")
        else:
            messagebox.showerror("Erro", f"Erro ao enviar: {response.status_code}\n{response.text}")

    except Exception as e:
        messagebox.showerror("Erro", str(e))

if __name__ == "__main__":
    main()