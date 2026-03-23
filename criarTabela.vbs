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

Function EsperarElemento(strId)
   Dim obj, t
   Set obj = Nothing

   For t = 1 To 15
      On Error Resume Next
      Set obj = session.findById(strId)
      On Error GoTo 0

      If Not obj Is Nothing Then
         Set EsperarElemento = obj
         Exit Function
      End If

      WScript.Sleep 500
   Next

   Set EsperarElemento = Nothing
End Function

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

' ------------------------------------------------------------
' Cabecalho inicial
' ------------------------------------------------------------
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLDS41:2202/ctxtDD02D-CONTFLAG").text = deliveryClass
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLDS41:2202/ctxtDD02D-CONTFLAG").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLDS41:2202/ctxtDD02D-CONTFLAG").caretPosition = 1
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 500
                           
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLDS41:2202/cmbDD02D-MAINFLAG").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLDS41:2202/cmbDD02D-MAINFLAG").key = "X"
WScript.Sleep 400

session.findById("wnd[0]/usr/txtDD02D-DDTEXT").text = tableText
session.findById("wnd[0]/usr/txtDD02D-DDTEXT").caretPosition = Len(tableText)
WScript.Sleep 300

session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpHEAD").select
WScript.Sleep 200
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN").select
WScript.Sleep 200
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF").select
WScript.Sleep 500

' ------------------------------------------------------------
' Preenche campos da aba DEF com suporte a N linhas
' ------------------------------------------------------------
Dim linhas, i, partes
Dim fieldName, keyFlag, notNullFlag, rollName, fieldDesc, fieldType, fieldLen, refTab, refField
Dim objTableDef, visibleRowsDef, visibleRow, scrollPos, currentScroll

linhas = SplitSafe(fieldsSpec, ";")

Set objTableDef = EsperarElemento("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLDS41:2201/tblSAPLSD41TC0")
If objTableDef Is Nothing Then
   WScript.Echo "Erro: grade da aba DEF nao foi encontrada."
   WScript.Quit 1
End If

visibleRowsDef = objTableDef.VisibleRowCount
currentScroll = -1

For i = 0 To UBound(linhas)
   If Trim(linhas(i)) <> "" Then
      partes = Split(linhas(i), "|")

      fieldName   = GetPart(partes, 0)
      keyFlag     = GetPart(partes, 1)
      notNullFlag = GetPart(partes, 2)
      rollName    = GetPart(partes, 3)
      fieldDesc   = GetPart(partes, 4)
      fieldType   = GetPart(partes, 5)
      fieldLen    = GetPart(partes, 6)
      refTab      = GetPart(partes, 7)
      refField    = GetPart(partes, 8)

      visibleRow = i Mod visibleRowsDef
      scrollPos = i - visibleRow

      If scrollPos <> currentScroll Then
         SetScrollIfNeeded objTableDef, scrollPos
         currentScroll = scrollPos
      End If

      session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLDS41:2201/tblSAPLSD41TC0/txtDD03D-FIELDNAME[0," & CStr(visibleRow) & "]").text = fieldName
      session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLDS41:2201/tblSAPLSD41TC0/txtDD03D-FIELDNAME[0," & CStr(visibleRow) & "]").setFocus
      session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLDS41:2201/tblSAPLSD41TC0/txtDD03D-FIELDNAME[0," & CStr(visibleRow) & "]").caretPosition = Len(fieldName)
      PressEnter

      If keyFlag = "1" Then
         session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLDS41:2201/tblSAPLSD41TC0/chkDD03P-KEYFLAG[1," & CStr(visibleRow) & "]").selected = True
      End If

      If notNullFlag = "1" Then
         session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLDS41:2201/tblSAPLSD41TC0/chkDD03P-NOTNULL[2," & CStr(visibleRow) & "]").selected = True
      End If

      session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLDS41:2201/tblSAPLSD41TC0/txtDD03D-ROLLNAME[3," & CStr(visibleRow) & "]").text = rollName
      session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLDS41:2201/tblSAPLSD41TC0/txtDD03D-ROLLNAME[3," & CStr(visibleRow) & "]").setFocus
      session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLDS41:2201/tblSAPLSD41TC0/txtDD03D-ROLLNAME[3," & CStr(visibleRow) & "]").caretPosition = Len(rollName)
      PressEnter
   End If
Next

' ------------------------------------------------------------
' Preenche REF_TAB e REF_FIELD na aba REFF, com suporte a N linhas
' ------------------------------------------------------------
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpF4V").select
WScript.Sleep 300
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpREFF").select
WScript.Sleep 500

Dim objTableRef, visibleRowsRef
Set objTableRef = EsperarElemento("wnd[0]/usr/tabsTAB_STRIP/tabpREFF/ssubTS_SCREEN:SAPLDS41:2103/tblSAPLSD41TC0")
If objTableRef Is Nothing Then
   WScript.Echo "Erro: grade da aba REFF nao foi encontrada."
   WScript.Quit 1
End If

visibleRowsRef = objTableRef.VisibleRowCount
currentScroll = -1

For i = 0 To UBound(linhas)
   If Trim(linhas(i)) <> "" Then
      partes = Split(linhas(i), "|")

      refTab   = GetPart(partes, 7)
      refField = GetPart(partes, 8)

      If refTab <> "" Or refField <> "" Then
         visibleRow = i Mod visibleRowsRef
         scrollPos = i - visibleRow

         If scrollPos <> currentScroll Then
            SetScrollIfNeeded objTableRef, scrollPos
            currentScroll = scrollPos
         End If

         If refTab <> "" Then
            session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpREFF/ssubTS_SCREEN:SAPLDS41:2103/tblSAPLSD41TC0/txtDD03P_D-REFTABLE[3," & CStr(visibleRow) & "]").text = refTab
            session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpREFF/ssubTS_SCREEN:SAPLDS41:2103/tblSAPLSD41TC0/txtDD03P_D-REFTABLE[3," & CStr(visibleRow) & "]").setFocus
            session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpREFF/ssubTS_SCREEN:SAPLDS41:2103/tblSAPLSD41TC0/txtDD03P_D-REFTABLE[3," & CStr(visibleRow) & "]").caretPosition = Len(refTab)
            PressEnter
         End If

         If refField <> "" Then
            session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpREFF/ssubTS_SCREEN:SAPLDS41:2103/tblSAPLSD41TC0/txtDD03P_D-REFFIELD[4," & CStr(visibleRow) & "]").text = refField
            session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpREFF/ssubTS_SCREEN:SAPLDS41:2103/tblSAPLSD41TC0/txtDD03P_D-REFFIELD[4," & CStr(visibleRow) & "]").setFocus
            session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpREFF/ssubTS_SCREEN:SAPLDS41:2103/tblSAPLSD41TC0/txtDD03P_D-REFFIELD[4," & CStr(visibleRow) & "]").caretPosition = Len(refField)
            PressEnter
         End If
      End If
   End If
Next

' ------------------------------------------------------------
' Ativacao inicial para abrir a tela de atributos adicionais
' ------------------------------------------------------------
session.findById("wnd[0]/tbar[1]/btn[27]").press
WScript.Sleep 1000

' ------------------------------------------------------------
' Popup de pacote
' ------------------------------------------------------------
Dim campoPacote
Set campoPacote = EsperarElemento("wnd[1]/usr/ctxtKO007-L_DEVCLASS")
If Not campoPacote Is Nothing Then
   campoPacote.text = packageName
   campoPacote.caretPosition = Len(packageName)

   On Error Resume Next
   session.findById("wnd[1]/tbar[0]/btn[7]").press
   If Err.Number <> 0 Then
      Err.Clear
      session.findById("wnd[1]/tbar[0]/btn[0]").press
   End If
   On Error GoTo 0

   WScript.Sleep 1000
End If

' ------------------------------------------------------------
' Popup de request
' ------------------------------------------------------------
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

' ------------------------------------------------------------
' Tela GNIRL: TABART e TABKAT
' ------------------------------------------------------------
Dim campoTabartFinal
Dim campoTabkatFinal

Set campoTabartFinal = EsperarElemento("wnd[0]/usr/tabsTABS/tabpGNRL/ssubTABS_SUBSC:SAPMSEDS1:0050/ctxtDD09V-TABART")
If Not campoTabartFinal Is Nothing Then
   campoTabartFinal.text = tabArt
   campoTabartFinal.caretPosition = Len(tabArt)
   PressEnter
End If

Set campoTabkatFinal = EsperarElemento("wnd[0]/usr/tabsTABS/tabpGNRL/ssubTABS_SUBSC:SAPMSEDS1:0050/ctxtDD09V-TABKAT")
If Not campoTabkatFinal Is Nothing Then
   campoTabkatFinal.text = tabKat
   campoTabkatFinal.setFocus
   campoTabkatFinal.caretPosition = Len(tabKat)
   PressEnter
End If

' ------------------------------------------------------------
' Salvar / voltar / ativar final
' ------------------------------------------------------------
On Error Resume Next
session.findById("wnd[0]/tbar[0]/btn[11]").press
WScript.Sleep 800
session.findById("wnd[0]/tbar[0]/btn[3]").press
WScript.Sleep 800
session.findById("wnd[0]/tbar[1]/btn[27]").press
On Error GoTo 0
WScript.Sleep 1000

WScript.Echo "Tabela " & tableName & " criada com sucesso."
