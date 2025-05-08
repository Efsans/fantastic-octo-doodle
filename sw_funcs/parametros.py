import pyodbc
from conectsql import get
import tkinter as tk
from tkinter import messagebox
def params(p):
  conexao = pyodbc.connect(get())
  cursor = conexao.cursor()
#SOL_MP
  query = """
SELECT

SX6.X6_CONTENG 

FROM SX6010	SX6

WHERE	SX6.D_E_L_E_T_	=	''	
AND SX6.X6_VAR	=	?

"""
  
  paramas = (p)
  cursor.execute(query, paramas)
  resultados = cursor.fetchall()
  cursor.close()
  conexao.close()
  # messagebox.showinfo("Sucesso", f"Parametros encontrados: {resultados}")
  return resultados
