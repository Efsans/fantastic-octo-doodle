import win32com.client

def get_mgr():
    swApp = win32com.client.Dispatch("SldWorks.Application")
    swModel = swApp.ActiveDoc
    if swModel is None:
        raise RuntimeError("Nenhum documento ativo no SolidWorks.")
    return swModel.Extension.CustomPropertyManager(""), swModel