import requests
from tkinter import messagebox  # se não importou ainda

def obter_ncm_sugerido(descricao):
    try:
        url = "http://localhost:8000/classificar-ncm"
        response = requests.post(url, json={"descricao": descricao}, timeout=5)
        if response.status_code == 200:
            data = response.json()
            return data.get("ncms", [])
        else:
            return []
    except Exception as e:
        print("Erro ao obter NCM:", e)
        messagebox.showerror("Erro", f"Erro ao buscar NCM: {e}")
        return []