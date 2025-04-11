import sys
import json
import requests


def enviar_para_api(json_data):
    api_url = "https://www.zohoapis.com/creator/custom/grupoaiz/SolidWorks?publickey=4WTWAfSnDWdjzatDCYr6gyJ4B"
    headers = {
        "Content-Type": "application/json"
    }

    try:
        response = requests.post(api_url, headers=headers, data=json_data.encode('utf-8'))
        print("Enviado com sucesso!")
        print("Resposta da API:", response.text)
    except Exception as e:
        print("Erro ao enviar para a API:", str(e))

def ver_na_api(codigo, filial="01mega"):
  api_url = f"http://177.155.130.19:8092/Rest/WSConsulta/Produto/{codigo}"
  headers = {
        "FILIAL": filial,
        "Accept": "application/json",
        "Authorization": "Basic U1VQT1JURS5QUk9USEVVUzpzdXBvcnRl"
    }

  try:
      response = requests.get(api_url, headers=headers)
      if response.status_code == 200:
          return response.json()
      else:
          return {"erro": f"Erro {response.status_code}"}
  except Exception as e:
      return {"erro": str(e)}


def processo(filhos, nivel=1, arquivo=None):
    for item in filhos:
        codigo = ""
        descricao = ""
        campo = ""

        for Key in item:
          if Key not in ["numero", "filhos"]:
            campo = Key
            codigo = item[Key]
            break
    if codigo:
      resultado = ver_na_api(codigo)
      descricao = resultado.get("Descricao", "Descrição não encontrada")

      linha = f"{'  ' * (nivel - 1)}{campo}: {codigo} - {descricao}\n"
      print(linha.strip())
    if arquivo:
                arquivo.write(linha)

        # Processa filhos recursivamente
    if "Filhos" in item:
        processo(item["Filhos"], nivel + 1, arquivo)



def consultar(json_data):
    dados = json.loads(json_data)
    filhos = dados.get("dados", {}).get("Filhos", [])

    with open("C:\\TEMP\\resultado.txt", "w", encoding="utf-8") as arq:
        processo(filhos, nivel=1, arquivo=arq)



def percorrer_estrutura(lista, nivel):
    indent = "  " * nivel
    for item in lista:
        for chave, valor in item.items():
            if chave != "Filhos":
                print(f"{indent}{chave}: {valor}")
        if "Filhos" in item and isinstance(item["Filhos"], list):
            percorrer_estrutura(item["Filhos"], nivel + 1)

def main():
    try:
        json_input = sys.stdin.read()  # Lê o JSON vindo do VBA
        json.loads(json_input)  # Verifica se é um JSON válido (opcional, segurança)
        enviar_para_api(json_input) and consultar(json_input)
        

        
    except Exception as e:
        print("Erro ao processar o JSON:", str(e))


if __name__ == "__main__":
    main()
