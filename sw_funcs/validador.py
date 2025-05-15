from sw_funcs.descri_tipo import regras, tipo_prefix


def validar_todos_campos(campos):
    """
    Valida todos os campos com base nas regras.
    :param campos: Dicionário com os nomes dos campos e os valores preenchidos.
    :return: Retorna True se todos os campos forem válidos, caso contrário, False.
    """
    for nome, valor in campos.items():
        valor = valor.strip().upper()  # Normaliza o valor
        if not valor:
            print(f"❌ O campo '{nome}' está vazio.")
            return False

        # Valida o prefixo usando as regras
        info = None
        for regra in regras():
            if valor in regra["prefixos"]:
                info = regra
                break

        if not info:
            print(f"❌ O prefixo '{valor}' no campo '{nome}' não é válido.")
            return False

        # Exibe informações adicionais (opcional)
        descricao = tipo_prefix(valor)
        print(f"✅ Campo '{nome}' válido. Prefixo: {valor}, Descrição: {descricao}, Tipo: {info['tipo']}.")

    return True