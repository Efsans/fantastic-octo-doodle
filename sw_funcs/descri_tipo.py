def tipo_prefix(valor):

  match valor:
    case "G1":
        descricao = "Destinado a Produtos de Mercado."
    case "G2":
        descricao = "Destinado Matéria Prima para Transformação."
    case "G3":
        descricao = "Destinado a produto consumível na Produção."
    case "G4":
        descricao = "Destinado a Produto de Uso e Consumo; Destinado a Ativo Imobilizado."
    case "G5":
        descricao = "Destinado a produto importado."
    case "I3" | "I4":
        descricao = "Destinado a Implemento Fabricado."
    case "I1" | "C1" | "M1":
        descricao = "Destinado a Partes Fabricadas."
    case "I2" | "C2" | "M2":
        descricao = "Destinado a Conjuntos Fabricados."
    case "M3" | "M4":
        descricao = "Destinado a Máquina Fabricada."
    case "C3" | "C4":
        descricao = "Destinado Cilindro Fabricado."
    case "C3S":
        descricao = "Destinado a subprodutos ou sucatas"
    case "SD":
        descricao = "Destinado a Serviços relacionados a Despesas"
    case "SC":
        descricao = "Destinado a Serviços relacionados a Custos de Produção"
    case "SR":
        descricao = "Destinado a Serviços relacionados a Venda de Serviços. Receita"
    case "Z":
        descricao = "Destinado produtos que não fazem parte do processo do Grupo. exemplo, Tela de Plasma, que será enviada para manutenção."
    case "I5":
        descricao = "Destinado a caminhões implementados."

    case "CA":
        descricao = "Destinado a caminhões para revenda, com código composto pelos 6 últimos dígitos do chassi."

    case "MA":
        descricao = "Destinado a máquinas, com código composto pelos 6 últimos dígitos do número de série."

    case "M5":
        descricao = "Destinado a máquinas implementadas, com código composto pelos 6 últimos dígitos do número de série."    
  
  return descricao 
  
def regras():
    REGRAS_PREFIXO = [
    {
        "prefixos": ["G1", "G2", "G3"],
        "tipo": "MP",
        "NCM": "1.1.0.30.20001",
        "armazem": "01",
        "origens": ["0", "3", "4", "5"],
        "empresa": "Aiz-Indústria"
    },
    {
        "prefixos": ["G5"],
        "tipo": "MP",
        "armazem": "01",
        "origens": ["1", "2", "6", "7"],
        "empresa": "Aiz-Indústria"
    },
    {
        "prefixos": ["I3", "I4", "C3", "C4", "M3", "M4"],
        "tipo": "PA",
        "armazem": "01",
        "origens": ["4"],
        "empresa": "Aiz-Indústria"
    },
    {
        "prefixos": ["I1", "I2", "C1", "C2", "M1", "M2"],
        "tipo": "PP",
        "armazem": "01",
        "origens": ["4"],
        "empresa": "Aiz-Indústria"
    },
    {
        "prefixos": ["I1", "I2", "C1", "C2", "M1", "M2"],
        "tipo": "PI",
        "armazem": "01",
        "origens": ["4"],
        "empresa": "Aiz-Indústria"
    },
    {
        "prefixos": ["C3S"],
        "tipo": "SB",
        "armazem": "A Definir",
        "origens": ["4"],
        "empresa": "Aiz-Indústria"
    },
    {
        "prefixos": ["SD", "SC", "SR"],
        "tipo": "SV",
        "armazem": "00",
        "origens": ["0"],
        "empresa": "Aiz-Indústria"
    },
    {
        "prefixos": ["G1", "G2", "G3", "I3", "I4", "I1", "C1", "M1", "I2", "C2", "M2", "M3", "M4", "C3", "C4"],
        "tipo": "MR",
        "armazem": "01",
        "origens": ["0", "3", "4", "5"],
        "empresa": "MegaPesados/Keera"
    },
    {
        "prefixos": ["SD", "SC", "SR"],
        "tipo": "SV",
        "armazem": "",
        "origens": ["0"],
        "empresa": "MegaPesados/Keera"
    },
    {
        "prefixos": ["G4"],
        "tipo": "MC",
        "armazem": "00",
        "origens": ["0"],
        "empresa": "MegaPesados"
    },
    {
        "prefixos": ["CA"],
        "tipo": "MC",
        "armazem": "A Definir",
        "origens": [],
        "empresa": "MegaPesados/Keera"
    },
    {
        "prefixos": ["I5"],
        "tipo": "MC",
        "armazem": "A Definir",
        "origens": [],
        "empresa": "MegaPesados/Keera"
    },
    {
        "prefixos": ["MA", "M5"],
        "tipo": "MC",
        "armazem": "A Definir",
        "origens": [],
        "empresa": "MegaPesados/Keera"
    }
]
    return REGRAS_PREFIXO

def prefixo_info(prefixo: str, Tipo: str ):
    prefixo = prefixo.upper().strip()
    for regra in regras():
        if prefixo  in regra["prefixos"] :
            if Tipo in regra["tipo"]:
                return {
                    "prefixo": prefixo,
                    "tipo": regra["tipo"],

                    "armazem": regra["armazem"],
                    "origens": regra["origens"],
                    "empresa": regra["empresa"]
                }
    return None  

#Debug
def debug(): 
    info = prefixo_info("G1", "MR")
    if info:
        print("Tipo:", info["tipo"])
        print("local:", info["armazem"])
    else:
        print("Prefixo não encontrado.")  

#debug()
  

  

      