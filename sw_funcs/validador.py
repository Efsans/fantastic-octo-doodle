import tkinter as tk
from tkinter import messagebox
from sw_funcs.descri_tipo import regras

def obter_regra(codigo):
    """
    Obtém a primeira regra correspondente ao prefixo de 2 caracteres do código informado.

    Parâmetros:
        codigo (str): O código a ser analisado.

    Retorna:
        dict ou None: Retorna o dicionário da regra correspondente, ou None se não houver.
    """
    codigo = codigo.strip().upper()
    # Apenas tenta prefixos de 2 caracteres
    p = codigo[:2]
    for regra in regras():
        prefixos = [prf.strip().upper() for prf in regra["prefixos"]]
        if p in prefixos:
            return regra
    return None

def validar_todos_campos(campos):
    """
    Valida todos os campos de um formulário de cadastro, de acordo com as regras de prefixo.

    Parâmetros:
        campos (dict): Dicionário com as chaves:
            - "Codigo": código do item
            - "Tipo": tipo do item (deve bater com a regra)
            - "Armazem": armazém (deve bater com a regra, se houver)
            - "Origem": origem (deve estar entre as permitidas pela regra, se houver)
            - "Descricao": descrição (não pode estar vazia)
            - "Unidade": unidade (não pode estar vazia)

    Retorna:
        bool: True se todos os campos forem válidos conforme as regras, False caso contrário.
    """
    codigo = campos.get("Codigo", "").strip()
    if len(codigo) < 2:
        print("❌ Código muito curto.")
        return False

    regra = obter_regra(codigo)
    if not regra:
        print(f"❌ Prefixo desconhecido em '{codigo}'.")
        return False

    # 1) Tipo
    esperado = regra["tipo"].upper()
    if campos.get("Tipo", "").strip().upper() != esperado:
        print(f"❌ Tipo incorreto: esperado '{esperado}'.")
        return False

    # 2) Armazem
    armazem_regra = regra.get("armazem", "").strip()
    if armazem_regra and campos.get("Armazem", "").strip() != armazem_regra:
        print(f"❌ Armazém incorreto: esperado '{armazem_regra}'.")
        return False

    # 3) Origem
    origens = regra.get("origens", [])
    if origens and campos.get("Origem", "").strip() not in origens:
        print(f"❌ Origem inválida: permitidas {origens}.")
        return False

    # 4) Descricao e Unidade: não podem estar vazios
    for campo in ("DescriÇão", "Unidade"):
        if not campos.get(campo, "").strip():
            print(f"❌ O campo '{campo}' está vazio.")
            return True

    print("✅ Todos os campos válidos.")
    return True

# campos = {
#     "Codigo": "G123456",
#     "Tipo": "MP",
#     "Armazem": "01",
#     "Origem": "0",
#     "Descricao": "Teste",
#     "Unidade": "KG"
# }
# validar_todos_campos(campos)
