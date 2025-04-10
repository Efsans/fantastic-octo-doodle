im swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swAssy As SldWorks.AssemblyDoc

Sub main()
    On Error GoTo ErrorHandler ' Habilita tratamento de erro
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "Nenhum documento está ativo no SolidWorks.", vbCritical, "Erro"
        Exit Sub
    End If

    If swModel.GetType = swDocASSEMBLY Then
        Set swAssy = swModel
        
        Dim nomeArquivo As String
        Dim inicio As Integer, fim As Integer
        Dim codigo As String
        
        ' Extrai o nome do arquivo sem caminho e extensão
        nomeArquivo = swModel.GetPathName
        If nomeArquivo = "" Then
            MsgBox "O arquivo ativo não possui um caminho válido.", vbCritical, "Erro"
            Exit Sub
        End If

        inicio = InStrRev(nomeArquivo, "\") + 1
        fim = InStrRev(nomeArquivo, ".")
        If inicio = 0 Or fim = 0 Then
            MsgBox "Falha ao extrair o nome do arquivo.", vbCritical, "Erro"
            Exit Sub
        End If

        codigo = Mid(nomeArquivo, inicio, fim - inicio)
        
        ' Gera JSON com nome do produto e hierarquia de componentes
        Dim jsonResult As String
        jsonResult = """Produto"":""" & codigo & """," & GerarJSON(swAssy.GetComponents(False), 1)
        
        ' Adiciona a estrutura de dados do JSON
        Dim json As String
        json = "{""dados"":{" & jsonResult & "}}"
        
        ' Envia o JSON para a API do Zoho
        Call EnviarParaAPI(json)
        
        Debug.Print json
    Else
        MsgBox "O documento ativo não é um assembly.", vbCritical, "Erro"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Ocorreu um erro: " & Err.Description, vbCritical, "Erro"
End Sub

Function GerarJSON(ByRef comps As Variant, ByVal nivel As Integer) As String
    On Error GoTo ErrorHandler ' Habilita tratamento de erro
    Dim i As Integer
    Dim swComp As SldWorks.Component2
    Dim childComps As Variant
    Dim compJSON As String
    Dim jsonList As String
    Dim nomeComp As String
    Dim campoNome As String
    Dim contador As Integer

    ' Define os nomes dos campos para até 7 níveis
    Select Case nivel
        Case 1: campoNome = "Modulo"
        Case 2: campoNome = "Conjunto"
        Case 3: campoNome = "Componente"
        Case 4: campoNome = "Subcomponente"
        Case 5: campoNome = "Elemento"
        Case 6: campoNome = "Detalhe"
        Case Else: campoNome = "Nivel" & nivel
    End Select

    jsonList = """Filhos"":["
    contador = 1

    For i = 0 To UBound(comps)
        Set swComp = comps(i)
        
        ' Extrai o último código após a última barra e remove o sufixo numérico
        Dim nomeCompleto As String
        Dim posBarra As Integer
        Dim codigoTemp As String
        Dim posHifen As Integer
        Dim codigoPadrao As String
        
        nomeCompleto = swComp.Name2
        posBarra = InStrRev(nomeCompleto, "/")
        If posBarra > 0 Then
            codigoTemp = Mid(nomeCompleto, posBarra + 1)
        Else
            codigoTemp = nomeCompleto
        End If

        posHifen = InStr(codigoTemp, "-")
        If posHifen > 0 Then
            codigoPadrao = Left(codigoTemp, posHifen - 1)
        Else
            codigoPadrao = codigoTemp
        End If

        ' Aplica o filtro
        If EhCodigoPadrao(codigoPadrao) Then
            nomeComp = codigoPadrao
            compJSON = "{" & """" & campoNome & """:""" & EscapeJSON(nomeComp) & """,""Numero"":" & contador
            contador = contador + 1

            childComps = swComp.GetChildren
            If Not IsEmpty(childComps) Then
                Dim filhosJSON As String
                filhosJSON = GerarJSON(childComps, nivel + 1)
                compJSON = compJSON & "," & filhosJSON
            End If
            compJSON = compJSON & "},"
            jsonList = jsonList & compJSON
        End If
    Next i
    
    If Right(jsonList, 1) = "," Then
        jsonList = Left(jsonList, Len(jsonList) - 1)
    End If
    
    jsonList = jsonList & "]"
    GerarJSON = jsonList
    Exit Function

ErrorHandler:
    MsgBox "Erro ao gerar JSON: " & Err.Description, vbCritical, "Erro"
    GerarJSON = ""
End Function


Function EhCodigoPadrao(codigo As String) As Boolean
    On Error GoTo ErrorHandler ' Habilita tratamento de erro
    ' Define o intervalo de tamanho esperado para os códigos
    Dim tamanhoMinimo As Integer: tamanhoMinimo = 1
    Dim tamanhoMaximo As Integer: tamanhoMaximo = 9

    ' Rejeita se tiver espaço ou for muito curto/longo
    If InStr(codigo, " ") > 0 Then
        EhCodigoPadrao = False
    ElseIf Len(codigo) < tamanhoMinimo Or Len(codigo) > tamanhoMaximo Then
        EhCodigoPadrao = False
    Else
        EhCodigoPadrao = True
    End If
    Exit Function

ErrorHandler:
    MsgBox "Erro ao validar código padrão: " & Err.Description, vbCritical, "Erro"
    EhCodigoPadrao = False
End Function

Function EscapeJSON(texto As String) As String
    On Error GoTo ErrorHandler ' Habilita tratamento de erro
    texto = Replace(texto, "\", "\\")
    texto = Replace(texto, """", "\""")
    EscapeJSON = texto
    Exit Function

ErrorHandler:
    MsgBox "Erro ao escapar JSON: " & Err.Description, vbCritical, "Erro"
    EscapeJSON = ""
End Function

Sub EnviarParaAPI(jsonData As String)
    On Error GoTo ErrorHandler ' Habilita tratamento de erro
    ' Declaração da variável para fazer requisições HTTP
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' URL da API do Zoho
    Dim apiUrl As String
    apiUrl = "https://www.zohoapis.com/creator/custom/grupoaiz/SolidWorks?publickey=4WTWAfSnDWdjzatDCYr6gyJ4B"
    
    ' Configura a requisição HTTP
    http.Open "POST", apiUrl, False
    http.setRequestHeader "Content-Type", "application/json" ' Define o cabeçalho Content-Type como JSON
    http.send jsonData ' Envia o JSON como corpo da requisição
    
    ' Obtém e exibe a resposta do servidor
    Dim response As String
    response = http.responseText
    MsgBox "Enviado com sucesso "
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao enviar para a API: " & Err.Description, vbCritical, "Erro"
End Sub
