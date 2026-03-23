' ============================================================
' ScriptCriarTabela.vbs
' Fluxo validado + linhas dinamicas via fieldsSpec
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

tableName     = "ZMM_TT123"
tableText     = "teste 002"
packageName   = "$TMP"
requestId     = ""
tabArt        = "APPL0"
tabKat        = "3"
fieldsSpec    = "MANDT|1|1|MANDT||;CNPJ|1|1|Z_MM_CNPJ||;LIMITE_CRED|0|0|Z_MM_LIMITE|TCURC|WAERS"
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

tableText = Left(tableText, 60)
deliveryClass = UCase(Left(Trim(CStr(deliveryClass)), 1))
If deliveryClass = "" Then deliveryClass = "A"

Function SplitSafe(txt, sep)
   If Trim(CStr(txt)) = "" Then
      SplitSafe = Array()
   Else
      SplitSafe = Split(CStr(txt), sep)
   End If
End Function

Function GetPart(parts, idx)
   If UBound(parts) >= idx Then
      GetPart = Trim(CStr(parts(idx)))
   Else
      GetPart = ""
   End If
End Function

Sub SafeEnter()
   session.findById("wnd[0]").sendVKey 0
   WScript.Sleep 300
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
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLSD41:2202/ctxtDD02D-CONTFLAG").text = deliveryClass
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLSD41:2202/ctxtDD02D-CONTFLAG").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLSD41:2202/ctxtDD02D-CONTFLAG").caretPosition = Len(deliveryClass)
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

' ============================================================
' INICIO MONTAGEM DINAMICA DAS LINHAS
' fieldsSpec = FIELD|KEY|NOTNULL|ELEMENT|REF_TABLE|REF_FIELD;...
' ============================================================
Dim rows, i, rowData, cols
Dim fieldName, keyFlag, notNull, rollName, refTable, refField
Dim defTableId, reffTableId
Dim visibleRowsDef, visibleRowsReff
Dim visibleRow, scrollBase
Dim hadRef

defTableId = "wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0"
reffTableId = "wnd[0]/usr/tabsTAB_STRIP/tabpREFF/ssubTS_SCREEN:SAPLSD41:2103/tblSAPLSD41TC0"

session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF").select

visibleRowsDef = session.findById(defTableId).VisibleRowCount
rows = SplitSafe(fieldsSpec, ";")

For i = 0 To UBound(rows)
   rowData = Trim(CStr(rows(i)))

   If rowData <> "" Then
      cols = Split(rowData, "|")

      fieldName = GetPart(cols, 0)
      keyFlag   = GetPart(cols, 1)
      notNull   = GetPart(cols, 2)
      rollName  = GetPart(cols, 3)
      refTable  = GetPart(cols, 4)
      refField  = GetPart(cols, 5)

      visibleRow = i Mod visibleRowsDef
      scrollBase = i - visibleRow

      If scrollBase > 0 Then
         session.findById(defTableId).verticalScrollbar.Position = scrollBase
         WScript.Sleep 300
      End If

      session.findById(defTableId & "/txtDD03D-FIELDNAME[0," & CStr(visibleRow) & "]").text = fieldName
      session.findById(defTableId & "/txtDD03D-FIELDNAME[0," & CStr(visibleRow) & "]").setFocus
      session.findById(defTableId & "/txtDD03D-FIELDNAME[0," & CStr(visibleRow) & "]").caretPosition = Len(fieldName)
      SafeEnter

      If keyFlag = "1" Then
         session.findById(defTableId & "/chkDD03P-KEYFLAG[1," & CStr(visibleRow) & "]").selected = True
      Else
         session.findById(defTableId & "/chkDD03P-KEYFLAG[1," & CStr(visibleRow) & "]").selected = False
      End If

      If notNull = "1" Then
         session.findById(defTableId & "/chkDD03P-NOTNULL[2," & CStr(visibleRow) & "]").selected = True
      Else
         session.findById(defTableId & "/chkDD03P-NOTNULL[2," & CStr(visibleRow) & "]").selected = False
      End If

      session.findById(defTableId & "/ctxtDD03D-ROLLNAME[3," & CStr(visibleRow) & "]").text = rollName
      session.findById(defTableId & "/ctxtDD03D-ROLLNAME[3," & CStr(visibleRow) & "]").setFocus
      session.findById(defTableId & "/ctxtDD03D-ROLLNAME[3," & CStr(visibleRow) & "]").caretPosition = Len(rollName)
      SafeEnter

      hadRef = False
      If refTable <> "" Or refField <> "" Then
         hadRef = True
         session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpREFF").select
         WScript.Sleep 300

         visibleRowsReff = session.findById(reffTableId).VisibleRowCount
         visibleRow = i Mod visibleRowsReff
         scrollBase = i - visibleRow

         If scrollBase > 0 Then
            session.findById(reffTableId).verticalScrollbar.Position = scrollBase
            WScript.Sleep 300
         End If

         If refTable <> "" Then
            session.findById(reffTableId & "/txtDD03P_D-REFTABLE[3," & CStr(visibleRow) & "]").text = refTable
            session.findById(reffTableId & "/txtDD03P_D-REFTABLE[3," & CStr(visibleRow) & "]").setFocus
            session.findById(reffTableId & "/txtDD03P_D-REFTABLE[3," & CStr(visibleRow) & "]").caretPosition = Len(refTable)
            SafeEnter
         End If

         If refField <> "" Then
            session.findById(reffTableId & "/txtDD03P_D-REFFIELD[4," & CStr(visibleRow) & "]").text = refField
            session.findById(reffTableId & "/txtDD03P_D-REFFIELD[4," & CStr(visibleRow) & "]").setFocus
            session.findById(reffTableId & "/txtDD03P_D-REFFIELD[4," & CStr(visibleRow) & "]").caretPosition = Len(refField)
            SafeEnter
         End If
      End If

      If hadRef Then
         session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF").select
         WScript.Sleep 300
      End If
   End If
Next

' ============================================================
' FIM MONTAGEM DINAMICA DAS LINHAS
' ============================================================

session.findById("wnd[0]/tbar[1]/btn[27]").press
session.findById("wnd[1]/usr/ctxtKO007-L_DEVCLASS").text = packageName
session.findById("wnd[1]/usr/ctxtKO007-L_DEVCLASS").caretPosition = Len(packageName)
session.findById("wnd[1]/tbar[0]/btn[7]").press

If UCase(Trim(packageName)) <> "$TMP" Then
   session.findById("wnd[1]/usr/ctxtKO008-TRKORR").text = requestId
   session.findById("wnd[1]/usr/ctxtKO008-TRKORR").caretPosition = Len(requestId)
   session.findById("wnd[1]/tbar[0]/btn[0]").press
End If

session.findById("wnd[0]/usr/tabsTABS/tabpGNRL/ssubTABS_SUBSC:SAPMSEDS1:0050/ctxtDD09V-TABART").text = tabArt
session.findById("wnd[0]/usr/tabsTABS/tabpGNRL/ssubTABS_SUBSC:SAPMSEDS1:0050/ctxtDD09V-TABART").caretPosition = Len(tabArt)
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/tabsTABS/tabpGNRL/ssubTABS_SUBSC:SAPMSEDS1:0050/ctxtDD09V-TABKAT").text = tabKat
session.findById("wnd[0]/usr/tabsTABS/tabpGNRL/ssubTABS_SUBSC:SAPMSEDS1:0050/ctxtDD09V-TABKAT").setFocus
session.findById("wnd[0]/usr/tabsTABS/tabpGNRL/ssubTABS_SUBSC:SAPMSEDS1:0050/ctxtDD09V-TABKAT").caretPosition = Len(tabKat)
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
