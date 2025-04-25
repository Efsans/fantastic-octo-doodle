import pyodbc

def get():
  # Função que retorna a string de conexão com o banco de dados SQL Server
  dados_conexao = ("Driver={SQL Server};"
                  "Server=TOTVSAPL;"
                  "Database=protheus12_producao;"
                  "UID=consulta;"
                  "PWD=consulta;")
  return dados_conexao


filial = "09ALFA"
codigo = "I2009128"

conexao = pyodbc.connect(get())      
cursor = conexao.cursor()

query=("""
        SELECT 
    B1_COD ,
    B1_TIPO AS Tipo,
    B1_UM AS UnidadeMedida,
    
    B1_DESC AS DescricaoProduto,
    B1_POSIPI AS NCM,
    B1_ORIGEM AS Orige,
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
  print("Atencao", "Filial nao encontrada")
  

if codigo:
  conditions.append("SB1.B1_COD = ?")
  params.append(codigo)
else:
  print("Atencao", "Codigo nao encontrado")
  
# Concatena a cláusula WHERE somente se houver condições
if conditions:
  query += " WHERE " + " AND ".join(conditions)  # Lista para acumular condições da cláusula WHERE
            
cursor.execute(query, params)
row = cursor.fetchone()
print(query)
print(params)
print(row)