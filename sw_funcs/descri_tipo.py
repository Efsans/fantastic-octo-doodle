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
  
  

  


  

  

      