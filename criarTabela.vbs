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

'=============================
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF").select

session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/txtDD03D-FIELDNAME[0,0]").text = "MANDT"
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/txtDD03D-FIELDNAME[0,0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/txtDD03D-FIELDNAME[0,0]").caretPosition = 5
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/chkDD03P-KEYFLAG[1,0]").selected = true
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/chkDD03P-NOTNULL[2,0]").selected = true
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/ctxtDD03D-ROLLNAME[3,0]").text = "MANDT"
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/ctxtDD03D-ROLLNAME[3,0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/ctxtDD03D-ROLLNAME[3,0]").caretPosition = 5
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/txtDD03D-FIELDNAME[0,1]").text = "CNPJ"
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/txtDD03D-FIELDNAME[0,1]").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/txtDD03D-FIELDNAME[0,1]").caretPosition = 4
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/chkDD03P-NOTNULL[2,1]").selected = true
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/ctxtDD03D-ROLLNAME[3,1]").text = "Z_MM_CNPJ"
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/ctxtDD03D-ROLLNAME[3,1]").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/ctxtDD03D-ROLLNAME[3,1]").caretPosition = 9
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0").columns.elementAt(0).width = 16
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/txtDD03D-FIELDNAME[0,2]").text = "LIMITE_CRED"
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/txtDD03D-FIELDNAME[0,2]").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/txtDD03D-FIELDNAME[0,2]").caretPosition = 11
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/ctxtDD03D-ROLLNAME[3,2]").text = "Z_MM_LIMITE"
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/ctxtDD03D-ROLLNAME[3,2]").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/ctxtDD03D-ROLLNAME[3,2]").caretPosition = 11
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpREFF").select
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpREFF/ssubTS_SCREEN:SAPLSD41:2103/tblSAPLSD41TC0/txtDD03P_D-REFTABLE[3,2]").text = "TCURC"
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpREFF/ssubTS_SCREEN:SAPLSD41:2103/tblSAPLSD41TC0/txtDD03P_D-REFTABLE[3,2]").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpREFF/ssubTS_SCREEN:SAPLSD41:2103/tblSAPLSD41TC0/txtDD03P_D-REFTABLE[3,2]").caretPosition = 5
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpREFF/ssubTS_SCREEN:SAPLSD41:2103/tblSAPLSD41TC0/txtDD03P_D-REFFIELD[4,2]").text = "WAERS"
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpREFF/ssubTS_SCREEN:SAPLSD41:2103/tblSAPLSD41TC0/txtDD03P_D-REFFIELD[4,2]").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpREFF/ssubTS_SCREEN:SAPLSD41:2103/tblSAPLSD41TC0/txtDD03P_D-REFFIELD[4,2]").caretPosition = 5
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/tbar[1]/btn[27]").press
session.findById("wnd[1]/usr/ctxtKO007-L_DEVCLASS").text = "$TMP"
session.findById("wnd[1]/usr/ctxtKO007-L_DEVCLASS").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[7]").press

session.findById("wnd[0]/usr/tabsTABS/tabpGNRL/ssubTABS_SUBSC:SAPMSEDS1:0050/ctxtDD09V-TABART").text = "APPL0"
session.findById("wnd[0]/usr/tabsTABS/tabpGNRL/ssubTABS_SUBSC:SAPMSEDS1:0050/ctxtDD09V-TABART").caretPosition = 5
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/tabsTABS/tabpGNRL/ssubTABS_SUBSC:SAPMSEDS1:0050/ctxtDD09V-TABKAT").text = "3"
session.findById("wnd[0]/usr/tabsTABS/tabpGNRL/ssubTABS_SUBSC:SAPMSEDS1:0050/ctxtDD09V-TABKAT").setFocus
session.findById("wnd[0]/usr/tabsTABS/tabpGNRL/ssubTABS_SUBSC:SAPMSEDS1:0050/ctxtDD09V-TABKAT").caretPosition = 1
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
