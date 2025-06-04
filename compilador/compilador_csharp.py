import subprocess
import sys
import os

def compilar_csharp(arquivo_cs, saida_dll):
    """
    Compila um arquivo C# (.cs) em uma DLL usando o .NET SDK.
    Requer que o .NET SDK esteja instalado e disponível no PATH.
    """
    if not os.path.isfile(arquivo_cs):
        print(f"Arquivo fonte não encontrado: {arquivo_cs}")
        return

    # Cria um diretório temporário para o projeto
    import tempfile
    with tempfile.TemporaryDirectory() as tempdir:
        # Cria um novo projeto de biblioteca
        subprocess.run(["dotnet", "new", "classlib", "-n", "TempLib", "-o", tempdir], check=True)
        # Substitui o arquivo gerado pelo arquivo do usuário
        csproj = os.path.join(tempdir, "TempLib.csproj")
        destino_cs = os.path.join(tempdir, os.path.basename(arquivo_cs))
        os.replace(arquivo_cs, destino_cs)
        # Remove o arquivo padrão criado pelo template
        padrao_cs = os.path.join(tempdir, "Class1.cs")
        if os.path.exists(padrao_cs):
            os.remove(padrao_cs)
        # Compila o projeto
        resultado = subprocess.run(
            ["dotnet", "build", csproj, "-c", "Release", "-o", tempdir],
            capture_output=True, text=True
        )
        print(resultado.stdout)
        if resultado.returncode == 0:
            # Move a DLL gerada para o destino desejado
            dll_gerada = os.path.join(tempdir, "TempLib.dll")
            os.replace(dll_gerada, saida_dll)
            print(f"Compilação bem-sucedida! DLL gerada em: {saida_dll}")
        else:
            print("Falha na compilação:")
            print(resultado.stderr)

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Uso: python compilador_csharp.py <arquivo.cs> <saida.dll>")
    else:
        compilar_csharp(sys.argv[1], sys.argv[2])
