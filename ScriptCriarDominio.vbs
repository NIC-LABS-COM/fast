' ============================================================
' ScriptCriarDominio.vbs
' Cria um dominio no SAP via SE11
' Recebe argumentos: domainName, domainText, dataType,
'                    dataLength, packageName, requestId
' ============================================================

Dim domainName
Dim domainText
Dim dataType
Dim dataLength
Dim packageName
Dim requestId

' --- Valores padrao (usados se nao receber argumentos) ---
domainName  = "Z_MM_CNPJ"
domainText  = "CNPJ"
dataType    = "CHAR"
dataLength  = "15"
packageName = "z_php"
requestId   = "A4HK904843"

' --- Recebe argumentos da linha de comando ---
If WScript.Arguments.Count >= 1 Then
   If Trim(CStr(WScript.Arguments(0))) <> "" Then
      domainName = CStr(WScript.Arguments(0))
   End If
End If

If WScript.Arguments.Count >= 2 Then
   If Trim(CStr(WScript.Arguments(1))) <> "" Then
      domainText = CStr(WScript.Arguments(1))
   End If
End If

If WScript.Arguments.Count >= 3 Then
   If Trim(CStr(WScript.Arguments(2))) <> "" Then
      dataType = CStr(WScript.Arguments(2))
   End If
End If

If WScript.Arguments.Count >= 4 Then
   If Trim(CStr(WScript.Arguments(3))) <> "" Then
      dataLength = CStr(WScript.Arguments(3))
   End If
End If

If WScript.Arguments.Count >= 5 Then
   If Trim(CStr(WScript.Arguments(4))) <> "" Then
      packageName = CStr(WScript.Arguments(4))
   End If
End If

If WScript.Arguments.Count >= 6 Then
   If Trim(CStr(WScript.Arguments(5))) <> "" Then
      requestId = CStr(WScript.Arguments(5))
   End If
End If

' --- Funcao para esperar um elemento carregar (ate 10 segundos) ---
Function EsperarElemento(strId)
   Dim obj, t
   Set obj = Nothing
   For t = 1 To 10
      On Error Resume Next
      Set obj = session.findById(strId)
      On Error GoTo 0
      If Not obj Is Nothing Then
         Set EsperarElemento = obj
         Exit Function
      End If
      WScript.Sleep 1000
   Next
   Set EsperarElemento = Nothing
End Function

' --- Conecta ao SAP GUI ---
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

' --- Navega para SE11 ---
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse11"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000

' --- Preenche nome do dominio e clica Criar ---
Dim campoDoma
Set campoDoma = EsperarElemento("wnd[0]/usr/ctxtRSRD1-DOMA_VAL")
If campoDoma Is Nothing Then
   WScript.Echo "Erro: Tela da SE11 nao carregou."
   WScript.Quit 1
End If
campoDoma.text = domainName
campoDoma.caretPosition = Len(domainName)

session.findById("wnd[0]/usr/btnPUSHADD").press
WScript.Sleep 1000

' --- Preenche Datatype ---
Dim campoDatatype
Set campoDatatype = EsperarElemento("wnd[0]/usr/tabsTAB_STRIP/tabpTAB1/ssubTS_SCREEN:SAPLSD11:1201/ctxtDD01D-DATATYPE")
If campoDatatype Is Nothing Then
   WScript.Echo "Erro: Tela de manutencao do dominio nao carregou."
   WScript.Quit 1
End If
campoDatatype.text = UCase(dataType)
campoDatatype.setFocus
campoDatatype.caretPosition = Len(UCase(dataType))

session.findById("wnd[0]").sendVKey 0
WScript.Sleep 500

' --- Preenche descricao ---
session.findById("wnd[0]/usr/txtDD01D-DDTEXT").text = domainText

' --- Preenche Tamanho ---
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpTAB1/ssubTS_SCREEN:SAPLSD11:1201/txtDD01D-LENG").text = dataLength
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpTAB1/ssubTS_SCREEN:SAPLSD11:1201/txtDD01D-LENG").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpTAB1/ssubTS_SCREEN:SAPLSD11:1201/txtDD01D-LENG").caretPosition = Len(dataLength)
WScript.Sleep 500

' --- Ativar (btn[27]) ---
session.findById("wnd[0]/tbar[1]/btn[27]").press
WScript.Sleep 1000

' --- Popup de pacote ---
Dim campoPacote
Set campoPacote = EsperarElemento("wnd[1]/usr/ctxtKO007-L_DEVCLASS")
If Not campoPacote Is Nothing Then
   campoPacote.text = packageName
   campoPacote.caretPosition = Len(packageName)
   session.findById("wnd[1]/tbar[0]/btn[7]").press
   WScript.Sleep 1000
End If

' --- Popup de request ---
Dim campoRequest
Set campoRequest = EsperarElemento("wnd[1]/usr/ctxtKO008-TRKORR")
If Not campoRequest Is Nothing Then
   campoRequest.text = requestId
   session.findById("wnd[1]/tbar[0]/btn[0]").press
   WScript.Sleep 1000
End If

WScript.Echo "Dominio " & domainName & " criado e ativado com sucesso."
