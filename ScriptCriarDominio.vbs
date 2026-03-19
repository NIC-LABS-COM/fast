' If Not IsObject(application) Then
'    Set SapGuiAuto  = GetObject("SAPGUI")
'    Set application = SapGuiAuto.GetScriptingEngine
' End If
' If Not IsObject(connection) Then
'    Set connection = application.Children(0)
' End If
' If Not IsObject(session) Then
'    Set session    = connection.Children(0)
' End If
' If IsObject(WScript) Then
'    WScript.ConnectObject session,     "on"
'    WScript.ConnectObject application, "on"
' End If

 Dim domainName
 Dim domainText
Dim dataType
Dim dataLength
Dim packageName
Dim requestId

' Function WndExists(sessionObj, wndId)
'    On Error Resume Next
'    Dim wnd
'    Set wnd = sessionObj.findById(wndId)
'    WndExists = (Err.Number = 0)
'    Set wnd = Nothing
'    Err.Clear
'    On Error GoTo 0
' End Function

domainName = "Z_MM_CNPJ"
domainText = "CNPJ"
dataType = "CHAR"
dataLength = "15"
packageName = "z_php"
requestId = "A4HK904843"

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

' session.findById("wnd[0]").maximize
' session.findById("wnd[0]/tbar[0]/okcd").text = "/nse11"
' session.findById("wnd[0]").sendVKey 0
' session.findById("wnd[0]/usr/radRSRD1-DOMA").setFocus
' session.findById("wnd[0]/usr/radRSRD1-DOMA").select
' session.findById("wnd[0]/usr/ctxtRSRD1-DOMA_VAL").text = domainName
' session.findById("wnd[0]/usr/btnPUSHADD").press
' session.findById("wnd[0]/usr/txtDD01D-DDTEXT").text = domainText
' session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpTAB1/ssubTS_SCREEN:SAPLSD11:1201/ctxtDD01D-DATATYPE").setFocus
' session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpTAB1/ssubTS_SCREEN:SAPLSD11:1201/ctxtDD01D-DATATYPE").caretPosition = 0
' session.findById("wnd[0]").sendVKey 4
' If WndExists(session, "wnd[1]") Then
'    session.findById("wnd[1]").close
' End If
' session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpTAB1/ssubTS_SCREEN:SAPLSD11:1201/ctxtDD01D-DATATYPE").text = UCase(dataType)
' session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpTAB1/ssubTS_SCREEN:SAPLSD11:1201/ctxtDD01D-DATATYPE").caretPosition = Len(UCase(dataType))
' session.findById("wnd[0]").sendVKey 0
' If WndExists(session, "wnd[1]") Then
'    session.findById("wnd[1]").sendVKey 0
' End If
' If WndExists(session, "wnd[1]") Then
'    session.findById("wnd[1]").sendVKey 0
' End If

' session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpTAB1/ssubTS_SCREEN:SAPLSD11:1201/txtDD01D-LENG").text = dataLength
' session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpTAB1/ssubTS_SCREEN:SAPLSD11:1201/txtDD01D-LENG").setFocus
' session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpTAB1/ssubTS_SCREEN:SAPLSD11:1201/txtDD01D-LENG").caretPosition = Len(dataLength)
' session.findById("wnd[0]").sendVKey 0
' If WndExists(session, "wnd[1]") Then
'    session.findById("wnd[1]").sendVKey 0
' End If
' If WndExists(session, "wnd[1]") Then
'    session.findById("wnd[1]").sendVKey 0
' End If
' session.findById("wnd[0]/tbar[0]/btn[11]").press
' If WndExists(session, "wnd[1]") Then
'    session.findById("wnd[1]/usr/ctxtKO007-L_DEVCLASS").text = packageName
'    session.findById("wnd[1]/usr/ctxtKO007-L_DEVCLASS").caretPosition = Len(packageName)
'    session.findById("wnd[1]/tbar[0]/btn[0]").press
' End If
' If WndExists(session, "wnd[1]") Then
'    session.findById("wnd[1]/usr/ctxtKO008-TRKORR").text = requestId
'    session.findById("wnd[1]/tbar[0]/btn[0]").press
' End If
' session.findById("wnd[0]/tbar[1]/btn[27]").press
' If WndExists(session, "wnd[1]") Then
'    session.findById("wnd[1]/tbar[0]/btn[0]").press
' End If

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
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse11"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/radRSRD1-DOMA").setFocus
session.findById("wnd[0]/usr/radRSRD1-DOMA").select
session.findById("wnd[0]/usr/ctxtRSRD1-DOMA_VAL").text = "zteste_dominio"
session.findById("wnd[0]/usr/ctxtRSRD1-DOMA_VAL").caretPosition = 14
session.findById("wnd[0]/usr/btnPUSHADD").press
