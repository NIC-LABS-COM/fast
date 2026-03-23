' ============================================================
' ScriptCriarDominio.vbs
' Cria um dominio no SAP via SE11
' Recebe argumentos: domainName, domainText, dataType,
'                    dataLength, packageName, requestId
'
' Regras:
' - Se packageName = "$TMP", objeto local e nao usa request
' - Se packageName <> "$TMP", requestId passa a ser obrigatorio
' ============================================================

Option Explicit

Dim application
Dim connection
Dim session
Dim SapGuiAuto

Dim domainName
Dim domainText
Dim dataType
Dim dataLength
Dim packageName
Dim requestId

' --- Valores padrao ---
domainName  = "Z_MM_CNPJ"
domainText  = "CNPJ"
dataType    = "CHAR"
dataLength  = "15"
packageName = "$TMP"
requestId   = ""

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

' --- Funcao para esperar um elemento carregar ---
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
   Set session = connection.Children(0)
End If

If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject application, "on"
End If

' --- Navega para SE11 ---
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse11"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000

' --- Marca opcao Dominio ---
Dim radioDoma
Set radioDoma = EsperarElemento("wnd[0]/usr/radRSRD1-DOMA")
If radioDoma Is Nothing Then
   WScript.Echo "Erro: opcao Dominio nao foi encontrada na SE11."
   WScript.Quit 1
End If

radioDoma.setFocus
radioDoma.select
WScript.Sleep 500

' --- Preenche nome do dominio e clica Criar ---
Dim campoDoma
Set campoDoma = EsperarElemento("wnd[0]/usr/ctxtRSRD1-DOMA_VAL")
If campoDoma Is Nothing Then
   WScript.Echo "Erro: campo do dominio nao carregou."
   WScript.Quit 1
End If

campoDoma.text = domainName
campoDoma.caretPosition = Len(domainName)

session.findById("wnd[0]/usr/btnPUSHADD").press
WScript.Sleep 1000

' --- Preenche comprimento ---
Dim campoLeng
Set campoLeng = EsperarElemento("wnd[0]/usr/tabsTAB_STRIP/tabpTAB1/ssubTS_SCREEN:SAPLSD11:1201/txtDD01D-LENG")
If campoLeng Is Nothing Then
   WScript.Echo "Erro: tela de manutencao do dominio nao carregou."
   WScript.Quit 1
End If

campoLeng.text = dataLength
WScript.Sleep 300

' --- Preenche datatype via F4/valor manual ---
Dim campoDatatype
Set campoDatatype = EsperarElemento("wnd[0]/usr/tabsTAB_STRIP/tabpTAB1/ssubTS_SCREEN:SAPLSD11:1201/ctxtDD01D-DATATYPE")
If campoDatatype Is Nothing Then
   WScript.Echo "Erro: campo DATATYPE nao foi encontrado."
   WScript.Quit 1
End If

campoDatatype.setFocus
campoDatatype.caretPosition = 0

session.findById("wnd[0]").sendVKey 4
WScript.Sleep 500

On Error Resume Next
session.findById("wnd[1]").sendVKey 12
On Error GoTo 0
WScript.Sleep 300

campoDatatype.text = LCase(dataType)
campoDatatype.caretPosition = Len(LCase(dataType))

session.findById("wnd[0]").sendVKey 0
WScript.Sleep 500

' --- Preenche descricao ---
Dim campoTexto
Set campoTexto = EsperarElemento("wnd[0]/usr/txtDD01D-DDTEXT")
If campoTexto Is Nothing Then
   WScript.Echo "Erro: campo de descricao do dominio nao foi encontrado."
   WScript.Quit 1
End If

campoTexto.text = domainText
campoTexto.caretPosition = Len(domainText)
WScript.Sleep 300

' --- Navega pelas abas como no script gravado ---
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpTAB2").select
WScript.Sleep 300
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpTAB3").select
WScript.Sleep 300
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpTAB1").select
WScript.Sleep 300

' --- Salvar/Ativar ---
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
' So usa request se o pacote NAO for $TMP
If UCase(Trim(packageName)) <> "$TMP" Then
   Dim campoRequest
   Set campoRequest = EsperarElemento("wnd[1]/usr/ctxtKO008-TRKORR")

   If campoRequest Is Nothing Then
      WScript.Echo "Erro: pacote informado nao e $TMP e o popup de request nao apareceu."
      WScript.Quit 1
   End If

   If Trim(requestId) = "" Then
      WScript.Echo "Erro: pacote nao e $TMP, mas nenhuma request foi informada."
      WScript.Quit 1
   End If

   campoRequest.text = requestId
   campoRequest.caretPosition = Len(requestId)
   session.findById("wnd[1]/tbar[0]/btn[0]").press
   WScript.Sleep 1000
End If

WScript.Echo "Dominio " & domainName & " criado com sucesso."
