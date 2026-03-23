' ============================================================
' ScriptCriarTabela.vbs
' Cria uma tabela no SAP via SE11
'
' Argumentos:
'   0 - tableName
'   1 - tableText
'   2 - packageName
'   3 - requestId
'   4 - tabArt
'   5 - tabKat
'   6 - fieldsSpec
'   7 - deliveryClass
'
' Regras:
' - Se packageName = "$TMP", objeto local e nao usa request
' - Se packageName <> "$TMP", requestId passa a ser obrigatorio
'
' Formato do fieldsSpec:
' FIELD|KEY|NOTNULL|ROLLNAME|DESC|TIPO|TAM|REFTAB|REFFIELD;...
'
' Exemplo:
' MANDT|1|1|MANDT|Mandante|CLNT|3||;
' CNPJ|1|1|Z_MM_CNPJ|CNPJ do fornecedor|CHAR|15||;
' NOME_FORNEC|0|1|Z_MM_NOME|Nome do Fornecedor|CHAR|50||;
' DATA_CAD|0|1|Z_MM_DATA|Data de Cadastro|DATS|8||;
' LIMITE_CRED|0|0|Z_MM_LIMITE|Limite de Credito|CURR|11|TCURC|WAERS
' ============================================================

Option Explicit

Dim application
Dim connection
Dim session
Dim SapGuiAuto

Dim tableName
Dim tableText
Dim packageName
Dim requestId
Dim tabArt
Dim tabKat
Dim fieldsSpec
Dim deliveryClass

tableName     = "ZMM_FORNECED10"
tableText     = "Tabela de Fornecedor"
packageName   = "$TMP"
requestId     = ""
tabArt        = "APPL0"
tabKat        = "3"
fieldsSpec    = "MANDT|1|1|MANDT|Mandante|CLNT|3||;CNPJ|1|1|Z_MM_CNPJ|CNPJ do fornecedor|CHAR|15||;NOME_FORNEC|0|1|Z_MM_NOME|Nome do Fornecedor|CHAR|50||;DATA_CAD|0|1|Z_MM_DATA|Data de Cadastro|DATS|8||;LIMITE_CRED|0|0|Z_MM_LIMITE|Limite de Credito|CURR|11|TCURC|WAERS"
deliveryClass = "A"

If WScript.Arguments.Count >= 1 Then
   If Trim(CStr(WScript.Arguments(0))) <> "" Then tableName = CStr(WScript.Arguments(0))
End If

If WScript.Arguments.Count >= 2 Then
   If Trim(CStr(WScript.Arguments(1))) <> "" Then tableText = CStr(WScript.Arguments(1))
End If

If WScript.Arguments.Count >= 3 Then
   If Trim(CStr(WScript.Arguments(2))) <> "" Then packageName = CStr(WScript.Arguments(2))
End If

If WScript.Arguments.Count >= 4 Then
   If Trim(CStr(WScript.Arguments(3))) <> "" Then requestId = CStr(WScript.Arguments(3))
End If

If WScript.Arguments.Count >= 5 Then
   If Trim(CStr(WScript.Arguments(4))) <> "" Then tabArt = CStr(WScript.Arguments(4))
End If

If WScript.Arguments.Count >= 6 Then
   If Trim(CStr(WScript.Arguments(5))) <> "" Then tabKat = CStr(WScript.Arguments(5))
End If

If WScript.Arguments.Count >= 7 Then
   If Trim(CStr(WScript.Arguments(6))) <> "" Then fieldsSpec = CStr(WScript.Arguments(6))
End If

If WScript.Arguments.Count >= 8 Then
   If Trim(CStr(WScript.Arguments(7))) <> "" Then deliveryClass = CStr(WScript.Arguments(7))
End If
                        
Function SplitSafe(txt, sep)
   If Trim(txt) = "" Then
      SplitSafe = Array()
   Else
      SplitSafe = Split(txt, sep)
   End If
End Function

Function GetPart(parts, idx)
   If UBound(parts) >= idx Then
      GetPart = Trim(parts(idx))
   Else
      GetPart = ""
   End If
End Function

Sub SetScrollIfNeeded(objTable, pos)
   On Error Resume Next
   objTable.verticalScrollbar.Position = pos
   On Error GoTo 0
   WScript.Sleep 400
End Sub

Sub PressEnter()
   session.findById("wnd[0]").sendVKey 0
   WScript.Sleep 400
End Sub

If Not IsObject(application) Then
   Set SapGuiAuto = GetObject("SAPGUI")
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

' ------------------------------------------------------------
' Navega para SE11
' ------------------------------------------------------------
session.findById("wnd[0]").maximize
WScript.Sleep 500
session.findById("wnd[0]/tbar[0]/okcd").text = "/NSE11"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000

session.findById("wnd[0]/usr/radRSRD1-TBMA").setFocus
session.findById("wnd[0]/usr/radRSRD1-TBMA").select
WScript.Sleep 500
' ------------------------------------------------------------
' Cria tabela
' ------------------------------------------------------------
Dim campoTabela
Set campoTabela = EsperarElemento("wnd[0]/usr/ctxtRSRD1-TBMA_VAL")
If campoTabela Is Nothing Then
   WScript.Echo "Erro: campo da tabela nao carregou."
   WScript.Quit 1
End If

campoTabela.text = tableName
campoTabela.caretPosition = Len(tableName)
session.findById("wnd[0]/usr/btnPUSHADD").press
WScript.Sleep 1000


session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN").select
WScript.Sleep 500                              

' ------------------------------------------------------------
' Cabecalho inicial - MINIMO
' ------------------------------------------------------------
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/txtDD02D-DDTEXT").text = tableText
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[0]").close
WScript.Sleep 300

session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLDS41:2202/txtVALUE_TEXT-DDTEXT").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLDS41:2202/txtVALUE_TEXT-DDTEXT").caretPosition = 0
session.findById("wnd[0]").sendVKey 4
WScript.Sleep 800

session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLDS41:2202/ctxtDD02D-CONTFLAG").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLDS41:2202/ctxtDD02D-CONTFLAG").caretPosition = 0
session.findById("wnd[0]").sendVKey 4                             
                              
session.findById("wnd[1]/usr/lbl[3,3]").setFocus
session.findById("wnd[1]/usr/lbl[3,3]").caretPosition = 20
session.findById("wnd[1]").sendVKey 2

