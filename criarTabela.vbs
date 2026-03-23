' ============================================================
' ScriptCriarTabela.vbs
' Versao minima: apenas abre SE11, seleciona Tabela,
' cria a tabela e preenche o cabecalho inicial
' ============================================================

Option Explicit

Dim application
Dim connection
Dim session
Dim SapGuiAuto

Dim tableName
Dim tableText

tableName = "ZMM_TT123"
tableText = "teste 002"

If WScript.Arguments.Count >= 1 Then
   If Trim(CStr(WScript.Arguments(0))) <> "" Then
      tableName = CStr(WScript.Arguments(0))
   End If
End If

If WScript.Arguments.Count >= 2 Then
   If Trim(CStr(WScript.Arguments(1))) <> "" Then
      tableText = CStr(WScript.Arguments(1))
   End If
End If

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

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse11"
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/radRSRD1-VIMA").setFocus
session.findById("wnd[0]/usr/radRSRD1-VIMA").select
session.findById("wnd[0]/usr/radRSRD1-TBMA").setFocus
session.findById("wnd[0]/usr/radRSRD1-TBMA").select

session.findById("wnd[0]/usr/ctxtRSRD1-TBMA_VAL").text = tableName
session.findById("wnd[0]/usr/ctxtRSRD1-TBMA_VAL").caretPosition = Len(tableName)
session.findById("wnd[0]/usr/btnPUSHADD").press

On Error Resume Next
session.findById("wnd[1]").sendVKey 0
Err.Clear
On Error GoTo 0

' ------------------------------------------------------------
' Cabecalho inicial
' ------------------------------------------------------------
On Error Resume Next
session.findById("wnd[0]/usr/txtDD02D-DDTEXT").text = tableText
If Err.Number <> 0 Then
   Err.Clear
   session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLSD41:2202/txtDD02D-DDTEXT").text = tableText
End If
On Error GoTo 0
WScript.Sleep 300

On Error Resume Next
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLSD41:2202/ctxtDD02D-CONTFLAG").text = "A"
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLSD41:2202/ctxtDD02D-CONTFLAG").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLSD41:2202/ctxtDD02D-CONTFLAG").caretPosition = 1
session.findById("wnd[0]").sendVKey 0

If Err.Number <> 0 Then
   Err.Clear
   session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLSD41:2202/ctxtDD02D-CONTFLAG").setFocus
   session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLSD41:2202/ctxtDD02D-CONTFLAG").caretPosition = 0
   session.findById("wnd[0]").sendVKey 4
   WScript.Sleep 800

   session.findById("wnd[1]/usr/cntlCUSTOM_CONTAINER/shellcont/shell").currentCellColumn = "_TEXT"
   session.findById("wnd[1]/usr/cntlCUSTOM_CONTAINER/shellcont/shell").selectedRows = "0"
   session.findById("wnd[1]/usr/cntlCUSTOM_CONTAINER/shellcont/shell").doubleClickCurrentCell
End If
On Error GoTo 0
WScript.Sleep 500

session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLSD41:2202/cmbDD02D-MAINFLAG").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLSD41:2202/cmbDD02D-MAINFLAG").key = "X"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 500
