import win32com.client
import os
import json
import requests
from tkinter import messagebox, Tk
import pyodbc

def get():
  # Função que retorna a string de conexão com o banco de dados SQL Server
  dados_conexao = ("Driver={SQL Server};"
                  "Server=TOTVSAPL;"
                  "Database=protheus12_producao;"
                  "UID=consulta;"
                  "PWD=consulta;")
#conexao = pyodbc.connect(get())      
#cursor = conexao.cursor()                  
  return dados_conexao

def gerar_lista(componentes, nivel, numero=0, lista=[], listadebug=[]):#gerar uma lista e salvar 
    for comp in componentes:
        numero += 1
        nome_comp = extrair_codigo_componente(comp.Name2)
        listadebug += f"numero: {numero} nivel: {nivel}",f"{obter_nome_por_nivel(nivel)}: {nome_comp}"
        lista.append(nome_comp)
        subcomps = comp.GetChildren
        if subcomps:
            # Chamada recursiva para subcomponentes
            gerar_lista(subcomps, nivel + 1)         
    if nivel == 1:
        messagebox.showinfo("Lista de Componentes debug", "\n".join(listadebug))
        messagebox.showinfo("lista de debug conponentes", "\n".join(lista))
        return lista

def consultar_lista(lista, resultado=None):
    conexao = pyodbc.connect(get())      
    cursor = conexao.cursor()
    try:
        placeholder = ", ".join(["?"] * len(lista))
        # messagebox.showinfo("placeholder", placeholder)
        query = f"SELECT B1_COD, B1_DESC, B1_FILIAL FROM SB1010 AS SB1 WHERE SB1.D_E_L_E_T_ = '' AND SB1.B1_COD IN ({placeholder}) AND SB1.B1_FILIAL = '09ALFA'"
    
        cursor.execute(query, lista)
    
        for row in cursor.fetchall():
            produto = row[1] if row[1] else "Indefinido"
            codigo = row[0] if row[0] else "Indefinido"
            filial = row[2] if row[2] else "Indefinido"
            messagebox.showinfo("Resultados da Consulta", f"Produto: {produto}\nCódigo: {codigo}\nFilial: {filial}")
            return resultado == row    
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao consultar a lista: {str(e)}")    


def tela():
    root = Tk()
    root.withdraw()
    root.title("Monitor de Propriedades")
    root.geometry("300x150")
    


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

        # gerar_lista(componentes, 1)
        lista = gerar_lista(componentes, 1)
        try:
            resultado=consultar_lista(lista)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao consultar a lista: {str(e)}")

        if not resultado:
            messagebox.showerror("Erro", "Nenhum resultado encontrado na consulta.")

        tela(resultado)


        json_dados = {
            "dados": {
                "Produto": codigo,
                **json.loads(gerar_json(componentes, 1))  # Gera estrutura hierárquica
            }
        }

        json_string = json.dumps(json_dados, indent=2)
        print(json_string)

        url = "https://www.zohoapis.com/creator/custom/grupoaiz/SolidWorks?publickey=4WTWAfSnDWdjzatDCYr6gyJ4"#colocar um "B" no final
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