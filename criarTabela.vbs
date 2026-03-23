' ============================================================
' ScriptCriarTabela.vbs
' Cria uma tabela no SAP via SE11
' Recebe argumentos:
'   tableName, tableText, packageName, requestId,
'   tabArt, tabKat, fieldsSpec, deliveryClass
'
' Regras:
' - Se packageName = "$TMP", objeto local e nao usa request
' - Se packageName <> "$TMP", requestId passa a ser obrigatorio
'
' fieldsSpec formato:
'   NOMECAMPO|KEYFLAG|NOTNULL|ROLLNAME;NOME2|KEYFLAG|NOTNULL|ROLLNAME2
' Exemplo:
'   MANDT|1|1|MANDT;CNPJ|1|1|Z_MM_CNPJ
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
fieldsSpec    = "MANDT|1|1|MANDT;CNPJ|1|1|Z_MM_CNPJ"
deliveryClass = "A"

' --- Recebe argumentos ---
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

Function SplitSafe(txt, sep)
   If Trim(txt) = "" Then
      SplitSafe = Array()
   Else
      SplitSafe = Split(txt, sep)
   End If
End Function

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

' --- Navega para SE11 ---
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse11"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000

' --- Preenche nome da tabela ---
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

' --- Abas iniciais como no script gravado ---
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpHEAD").select
WScript.Sleep 200
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF").select
WScript.Sleep 200
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpF4V").select
WScript.Sleep 200
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpREFF").select
WScript.Sleep 200

' --- Delivery class / contflag ---
Dim campoContFlag
Set campoContFlag = EsperarElemento("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLDS41:2202/ctxtDD02D-CONTFLAG")
If campoContFlag Is Nothing Then
   WScript.Echo "Erro: campo DD02D-CONTFLAG nao foi encontrado."
   WScript.Quit 1
End If

campoContFlag.setFocus
campoContFlag.caretPosition = 0
WScript.Sleep 200

On Error Resume Next
session.findById("wnd[0]").sendVKey 4
WScript.Sleep 300
session.findById("wnd[1]/usr/lbl[3,3]").setFocus
session.findById("wnd[1]/usr/lbl[3,3]").caretPosition = 4
session.findById("wnd[0]").sendVKey 2
On Error GoTo 0
WScript.Sleep 300

Dim comboContFlag
Set comboContFlag = EsperarElemento("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLDS41:2202/cmbDD02D-CONTFLAG")
If Not comboContFlag Is Nothing Then
   comboContFlag.setFocus
   comboContFlag.key = deliveryClass
   WScript.Sleep 300
Else
   campoContFlag.text = deliveryClass
   WScript.Sleep 300
End If

' --- Aba definicao ---
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF").select
WScript.Sleep 300

' --- Texto da tabela ---
Dim campoTextoTabela
Set campoTextoTabela = EsperarElemento("wnd[0]/usr/textDD02D-DDTEXT")
If campoTextoTabela Is Nothing Then
   WScript.Echo "Erro: campo de descricao da tabela nao foi encontrado."
   WScript.Quit 1
End If

campoTextoTabela.text = tableText
campoTextoTabela.caretPosition = Len(tableText)
WScript.Sleep 300

' --- Volta nas abas conforme script gravado ---
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF").select
WScript.Sleep 200
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpF4V").select
WScript.Sleep 200
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpREFF").select
WScript.Sleep 200
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpHEAD").select
WScript.Sleep 300

' --- Preenche tipo de tabela / categoria / delivery class na principal ---
On Error Resume Next
Dim campoTabArt
Set campoTabArt = session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLDS41:2202/cmbDD02D-TABCLASS")
If Err.Number = 0 Then
   campoTabArt.key = tabArt
End If
Err.Clear

Dim campoTabKat
Set campoTabKat = session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLDS41:2202/cmbDD02D-MAINFLAG")
If Err.Number = 0 Then
   campoTabKat.setFocus
   campoTabKat.key = tabKat
End If
Err.Clear
On Error GoTo 0
WScript.Sleep 500

' --- Vai para aba de campos ---
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF").select
WScript.Sleep 500

Dim linhas, i, partes
Dim fieldName, keyFlag, notNullFlag, rollName
linhas = SplitSafe(fieldsSpec, ";")

For i = 0 To UBound(linhas)
   If Trim(linhas(i)) <> "" Then
      partes = Split(linhas(i), "|")

      If UBound(partes) >= 3 Then
         fieldName   = Trim(partes(0))
         keyFlag     = Trim(partes(1))
         notNullFlag = Trim(partes(2))
         rollName    = Trim(partes(3))

         ' Nome do campo
         session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLDS41:2201/tblSAPLSD41TC0/txtDD03D-FIELDNAME[0," & CStr(i) & "]").text = fieldName
         session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLDS41:2201/tblSAPLSD41TC0/txtDD03D-FIELDNAME[0," & CStr(i) & "]").setFocus
         session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLDS41:2201/tblSAPLSD41TC0/txtDD03D-FIELDNAME[0," & CStr(i) & "]").caretPosition = Len(fieldName)

         ' Enter para validar linha
         session.findById("wnd[0]").sendVKey 0
         WScript.Sleep 300

         ' Key flag
         If keyFlag = "1" Then
            session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLDS41:2201/tblSAPLSD41TC0/chkDD03P-KEYFLAG[1," & CStr(i) & "]").selected = True
         End If

         ' Not null
         If notNullFlag = "1" Then
            session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLDS41:2201/tblSAPLSD41TC0/chkDD03P-NOTNULL[2," & CStr(i) & "]").selected = True
         End If

         ' Rollname / elemento de dados
         session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLDS41:2201/tblSAPLSD41TC0/txtDD03D-ROLLNAME[3," & CStr(i) & "]").text = rollName
         session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLDS41:2201/tblSAPLSD41TC0/txtDD03D-ROLLNAME[3," & CStr(i) & "]").setFocus
         session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLDS41:2201/tblSAPLSD41TC0/txtDD03D-ROLLNAME[3," & CStr(i) & "]").caretPosition = Len(rollName)

         session.findById("wnd[0]").sendVKey 0
         WScript.Sleep 400
      End If
   End If
Next

' --- Salvar ---
session.findById("wnd[0]/tbar[0]/btn[11]").press
WScript.Sleep 1000

' --- Popup de pacote ---
Dim campoPacote
Set campoPacote = EsperarElemento("wnd[1]/usr/ctxtKO007-L_DEVCLASS")
If Not campoPacote Is Nothing Then
   campoPacote.text = packageName
   campoPacote.caretPosition = Len(packageName)

   On Error Resume Next
   session.findById("wnd[1]/tbar[0]/btn[0]").press
   If Err.Number <> 0 Then
      Err.Clear
      session.findById("wnd[1]/tbar[0]/btn[7]").press
   End If
   On Error GoTo 0

   WScript.Sleep 1000
End If

' --- Popup de request ---
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

' --- Ativar ---
session.findById("wnd[0]/tbar[1]/btn[27]").press
WScript.Sleep 1000

' --- Confirma popup final, se existir ---
On Error Resume Next
session.findById("wnd[1]").sendVKey 0
On Error GoTo 0
WScript.Sleep 500

WScript.Echo "Tabela " & tableName & " criada com sucesso."
