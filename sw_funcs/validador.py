import tkinter as tk
from tkinter import messagebox
from sw_funcs.descri_tipo import regras

def obter_regras_possiveis(codigo):
    codigo = codigo.strip().upper()
    p = codigo[:2]
    return [
        possiveis for possiveis in regras()
        if p in [pref.strip().upper() for pref in possiveis["prefixos"]]
    ]

def validar_todos_campos(campos):
    codigo = campos.get("Codigo", "").strip().upper()
    if len(codigo) < 8:
        print("❌ Código muito curto.")
        return False

    possiveis = obter_regras_possiveis(codigo)
    if not possiveis:
        print(f"❌ Prefixo desconhecido em '{codigo}'.")
        return False

    for regra in possiveis:
        # 1) Tipo
        if campos.get("Tipo", "").strip().upper() != regra["tipo"].upper():
            continue

        # 2) Armazém
        arm_regra = regra.get("armazem", "").strip()
        if arm_regra and campos.get("Armazem", "").strip() != arm_regra:
            continue

        # 3) Origem
        origens = regra.get("origens", [])
        if origens and campos.get("Origem", "").strip() not in origens:
            continue

        # 4) Descrição e Unidade
        if not campos.get("Descrição", "").strip() or not campos.get("Unidade", "").strip():
            continue

        print(f"✅ Regra válida: prefixo {codigo[:2]} → tipo {regra['tipo']} / empresa {regra['empresa']}")
        return True

    print("❌ Nenhuma regra compatível com todos os campos.")
    return False

# campos = {
#     "Codigo": "G123456",
#     "Tipo": "MP",
#     "Armazem": "01",
#     "Origem": "0",
#     "Descricao": "Teste",
#     "Unidade": "KG"
# }
# validar_todos_campos(campos)
