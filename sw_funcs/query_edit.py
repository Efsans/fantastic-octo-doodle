from sw_funcs.conectsql import get
import pyodbc


def consultar_protheu(cod="I2009128"):
  conexao = pyodbc.connect(get())
  cursor = conexao.cursor()

  query=("""
        SELECT 
    B1_COD
    FROM SB1010
    where B1_COD = (?)   
        
      """)
  params = (cod)


  cursor.execute(query,params)
  resultados = cursor.fetchall()
  cursor.close()
  conexao.close()

  if len(resultados) == 0:
    return False
  
  else: 
    return True

