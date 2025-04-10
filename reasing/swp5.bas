Dim swApp As SldWorks.SldWorks
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
        
        Dim nomeArquivo As String, codigo As String
        
        ' Extrai o nome do arquivo sem caminho e extensão
        nomeArquivo = swModel.GetPathName
        If nomeArquivo = "" Then
            MsgBox "O arquivo ativo não possui um caminho válido.", vbCritical, "Erro"
            Exit Sub
        End If
        
        codigo = ExtrairCodigoArquivo(nomeArquivo)
        If codigo = "" Then
            MsgBox "Falha ao extrair o código do arquivo.", vbCritical, "Erro"
            Exit Sub
        End If
        
        ' Gera JSON com nome do produto e hierarquia de componentes
        Dim jsonResult As String
        jsonResult = """Produto"":""" & codigo & """," & GerarJSON(swAssy.GetComponents(False), 1)
        
        Dim json As String
        json = "{""dados"":{" & jsonResult & "}}"
        
        ' Envia o JSON para a API e exibe no console
        Call EnviarParaAPI(json)
        Debug.Print "JSON Gerado:", json
        
        ' Percorre o JSON gerado para exibição
        Call LerEstruturaDoJSON(json)
    Else
        MsgBox "O documento ativo não é um assembly.", vbCritical, "Erro"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Ocorreu um erro: " & Err.Description, vbCritical, "Erro"
End Sub

Function ExtrairCodigoArquivo(nomeArquivo As String) As String
    Dim inicio As Integer, fim As Integer
    inicio = InStrRev(nomeArquivo, "\") + 1
    fim = InStrRev(nomeArquivo, ".")
    If inicio > 0 And fim > 0 Then
        ExtrairCodigoArquivo = Mid(nomeArquivo, inicio, fim - inicio)
    Else
        ExtrairCodigoArquivo = ""
    End If
End Function

Function GerarJSON(ByRef comps As Variant, ByVal nivel As Integer) As String
    Dim i As Integer, jsonList As String
    Dim campoNome As String, compJSON As String
    Dim contador As Integer

    campoNome = ObterCampoNome(nivel)
    jsonList = """Filhos"":["

    contador = 0 ' Inicia o contador no nível atual

    For i = LBound(comps) To UBound(comps)
        Dim swComp As SldWorks.Component2, nomeComp As String, childComps As Variant
        Set swComp = comps(i)

        nomeComp = ExtrairCodigoComponente(swComp.Name2)
        If EhCodigoPadrao(nomeComp) Then
            contador = contador + 1 ' Incrementa o contador do nível atual

            ' Cria objeto JSON com nome e número
            compJSON = "{" & """" & campoNome & """:""" & EscapeJSON(nomeComp) & """,""Numero"":" & contador

            ' Processa os filhos recursivamente
            childComps = swComp.GetChildren
            If Not IsEmpty(childComps) Then
                compJSON = compJSON & "," & GerarJSON(childComps, nivel + 1)
            End If

            compJSON = compJSON & "},"
            jsonList = jsonList & compJSON
        End If
    Next i

    If Right(jsonList, 1) = "," Then jsonList = Left(jsonList, Len(jsonList) - 1)
    jsonList = jsonList & "]"
    GerarJSON = jsonList
End Function


Function ObterCampoNome(nivel As Integer) As String
    Select Case nivel
        Case 1: ObterCampoNome = "Modulo"
        Case 2: ObterCampoNome = "Conjunto"
        Case 3: ObterCampoNome = "Componente"
        Case 4: ObterCampoNome = "Subcomponente"
        Case Else: ObterCampoNome = "Nivel" & nivel
    End Select
End Function

Function ExtrairCodigoComponente(nomeCompleto As String) As String
    Dim posBarra As Integer, posHifen As Integer
    Dim codigoTemp As String, codigoPadrao As String

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

    ExtrairCodigoComponente = codigoPadrao
End Function

Function EhCodigoPadrao(codigo As String) As Boolean
    Dim tamanhoMinimo As Integer: tamanhoMinimo = 1
    Dim tamanhoMaximo As Integer: tamanhoMaximo = 9

    EhCodigoPadrao = (InStr(codigo, " ") = 0 And Len(codigo) >= tamanhoMinimo And Len(codigo) <= tamanhoMaximo)
End Function

Sub LerEstruturaDoJSON(jsonTexto As String)
    Dim jsonObj As Object, dados As Object
    Set jsonObj = JsonConverter.ParseJson(jsonTexto)
    Set dados = jsonObj("dados")
    
    Debug.Print "Produto:", dados("Produto")
    If dados.Exists("Filhos") Then
        ProcurarItem dados("Filhos"), 1
    End If
End Sub

Sub ProcurarItem(filhos As Collection, nivel As Integer)
    Dim item As Object, indent As String
    indent = String(nivel * 2, " ")

    For Each item In filhos
        ' Verifica se o item atual é o que queremos (Número = 782)
        If item.Exists("Numero") And item("Numero") = 3 Then
            Debug.Print indent & "Módulo encontrado:", item("Modulo"), "Número:", item("Numero")
            
            ' Verifica se ele possui filhos
            If item.Exists("Filhos") Then
                Dim subitem As Object
                For Each subitem In item("Filhos")
                    ' Verifica se algum dos filhos tem Número = 4
                    If subitem.Exists("Numero") And subitem("Numero") = 1 Then
                        Debug.Print indent & "  Conjunto encontrado:", subitem("Conjunto"), "Número:", subitem("Numero")
                        Exit For ' Encontrou, pode sair do loop
                    End If
                Next subitem
            End If
            
            Exit Sub ' Já encontramos o módulo 782, não precisa continuar
        End If

        ' Se o item atual tiver filhos, continua a busca recursiva
        If item.Exists("Filhos") Then
            ProcurarItem item("Filhos"), nivel + 1
        End If
    Next item
End Sub


Function EscapeJSON(texto As String) As String
    texto = Replace(texto, "\", "\\")
    texto = Replace(texto, """", "\""")
    EscapeJSON = texto
End Function

Sub EnviarParaAPI(jsonData As String)
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    Dim apiUrl As String
    apiUrl = "https://www.zohoapis.com/creator/custom/grupoaiz/SolidWorks?publickey=4WTWAfSnDWdjzatDCYr6gyJ4B"
    
    http.Open "POST", apiUrl, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send jsonData
    
    MsgBox "Enviado com sucesso: " & http.responseText
End Sub
