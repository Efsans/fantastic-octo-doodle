import win32com.client
import os

def nome_arquivo():
    try:
        swApp = win32com.client.Dispatch("SldWorks.Application")
        swModel = swApp.ActiveDoc
        if swModel is None:
            print("Nenhum documento aberto.")
            return None

        nomeArquivo = swModel.GetPathName
        if callable(nomeArquivo):
            nomeArquivo = nomeArquivo()

        if not nomeArquivo:
            print("Arquivo não salvo ou sem caminho.")
            return None

        nome_arquivo = os.path.basename(nomeArquivo)
        codigo = os.path.splitext(nome_arquivo)[0]
        return codigo
    except Exception as e:
        print(f"Erro: {e}")
        return None
