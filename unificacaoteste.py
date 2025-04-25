import win32com.client
import os
import json
import requests
from tkinter import Tk, Toplevel, ttk, messagebox, Button
import pyodbc

# Função para conexão ao banco de dados
def get():
    return ("Driver={SQL Server};"
            "Server=TOTVSAPL;"
            "Database=protheus12_producao;"
            "UID=consulta;"
            "PWD=consulta;")

# Função para consulta ao banco de dados
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

# Função para exibir a interface gráfica
def exibir_tela(resultados, dados_hierarquicos, enviar_api_func):
    # Janela principal
    janela = Toplevel()
    janela.title("Resultados Hierárquicos")
    janela.geometry("800x600")

    # Configurar Treeview para exibir hierarquia
    tree = ttk.Treeview(janela, columns=("Descrição", "Filial"), show="tree headings")
    tree.heading("#0", text="Código")
    tree.heading("Descrição", text="Descrição")
    tree.heading("Filial", text="Filial")
    tree.column("Descrição", width=300)
    tree.column("Filial", width=100)

    # Função recursiva para preencher a hierarquia
    def adicionar_no_pai(pai_id, filhos):
        for filho in filhos:
            codigo = filho["codigo"]
            descricao = filho["descricao"]
            filial = filho.get("filial", "")
            
            # Adicionar o item no Treeview
            item_id = tree.insert(pai_id, "end", text=codigo, values=(descricao, filial))
            
            # Se o item tiver filhos, chamar recursivamente
            if "filhos" in filho and filho["filhos"]:
                adicionar_no_pai(item_id, filho["filhos"])

    # Preencher a hierarquia a partir do nível raiz
    adicionar_no_pai("", dados_hierarquicos)
    tree.pack(fill="both", expand=True)

    # Botão para enviar os dados para a API
    Button(janela, text="Enviar para API", command=enviar_api_func).pack(pady=10)

    # Botão para fechar a janela
    Button(janela, text="Fechar", command=janela.destroy).pack(pady=5)

# Função para estruturar os dados em formato hierárquico
def organizar_hierarquia(resultados):
    # Exemplo de organização hierárquica
    hierarquia = []
    for row in resultados:
        hierarquia.append({
            "codigo": row[0],
            "descricao": row[1],
            "filial": row[2],
            "filhos": []  # Adicione lógica para preencher filhos reais, se necessário
        })
    return hierarquia

# Função para enviar dados à API
def enviar_para_api(dados_hierarquicos):
    try:
        url = "https://www.zohoapis.com/creator/custom/grupoaiz/SolidWorks?publickey=4WTWAfSnDWdjzatDCYr6gyJ4"
        headers = {"Content-Type": "application/json"}
        response = requests.post(url, headers=headers, data=json.dumps(dados_hierarquicos, indent=2))

        if response.status_code == 200:
            messagebox.showinfo("Sucesso", "Dados enviados com sucesso!")
        else:
            messagebox.showerror("Erro", f"Erro ao enviar: {response.status_code}\n{response.text}")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao enviar para API: {str(e)}")

# Função principal
def main():
    root = Tk()
    root.withdraw()

    try:
        # Simulação de dados do SolidWorks
        lista_componentes = ['G1000553', 'I2001963', 'X3007890']
        resultados = consultar_lista(lista_componentes)

        if resultados:
            # Organizar dados em hierarquia
            dados_hierarquicos = organizar_hierarquia(resultados)

            # Exibir a interface gráfica com hierarquia e botão de envio
            exibir_tela(resultados, dados_hierarquicos, lambda: enviar_para_api(dados_hierarquicos))
        else:
            messagebox.showinfo("Sem Resultados", "Nenhum dado encontrado para os componentes fornecidos.")

    except Exception as e:
        messagebox.showerror("Erro", str(e))

if __name__ == "__main__":
    main()