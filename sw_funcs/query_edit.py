from sw_funcs.conectsql import get
import pyodbc
import tkinter as tk
from tkinter import messagebox

def consultar_protheu(cod):
    # Garante que cod é string e não vazio

    conexao = None
    cursor = None
    try:
        conexao = pyodbc.connect(get())
        cursor = conexao.cursor()

        query = """
            SELECT B1_COD
            FROM SB1010
            WHERE B1_COD = ?
        """
        cursor.execute(query, (cod))
        resultados = cursor.fetchall()
        if len(resultados) > 0 :
            return True
        else:
            return False 

    except Exception as e:
        print(f"Erro ao consultar Protheus: {e}")
        return False

    finally:
        if cursor:
            cursor.close()
        if conexao:
            conexao.close()


