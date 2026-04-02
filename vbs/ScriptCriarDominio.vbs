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

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse11"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/radRSRD1-DOMA").setFocus
session.findById("wnd[0]/usr/radRSRD1-DOMA").select
session.findById("wnd[0]/usr/ctxtRSRD1-DOMA_VAL").text = domainName
session.findById("wnd[0]/usr/ctxtRSRD1-DOMA_VAL").caretPosition = Len(domainName)
session.findById("wnd[0]/usr/btnPUSHADD").press
WScript.Sleep 1000
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpTAB1/ssubTS_SCREEN:SAPLSD11:1201/txtDD01D-LENG").text = dataLength
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpTAB1/ssubTS_SCREEN:SAPLSD11:1201/ctxtDD01D-DATATYPE").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpTAB1/ssubTS_SCREEN:SAPLSD11:1201/ctxtDD01D-DATATYPE").caretPosition = 0
session.findById("wnd[0]").sendVKey 4
WScript.Sleep 500
session.findById("wnd[1]").sendVKey 12
WScript.Sleep 500
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpTAB1/ssubTS_SCREEN:SAPLSD11:1201/ctxtDD01D-DATATYPE").text = dataType
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpTAB1/ssubTS_SCREEN:SAPLSD11:1201/ctxtDD01D-DATATYPE").caretPosition = Len(dataType)
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 800
CheckSapError "Validar tipo " & dataType
session.findById("wnd[0]/usr/txtDD01D-DDTEXT").text = domainText
session.findById("wnd[0]/usr/txtDD01D-DDTEXT").caretPosition = Len(domainText)
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpTAB2").select
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpTAB3").select
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpTAB1").select
session.findById("wnd[0]/tbar[1]/btn[27]").press
WScript.Sleep 1000
session.findById("wnd[1]/usr/ctxtKO007-L_DEVCLASS").text = packageName
session.findById("wnd[1]/usr/ctxtKO007-L_DEVCLASS").caretPosition = Len(packageName)
session.findById("wnd[1]/tbar[0]/btn[7]").press
WScript.Sleep 1000

If UCase(Trim(packageName)) <> "$TMP" Then
   session.findById("wnd[1]/usr/ctxtKO008-TRKORR").text = requestId
   session.findById("wnd[1]/usr/ctxtKO008-TRKORR").caretPosition = Len(requestId)
   session.findById("wnd[1]/tbar[0]/btn[0]").press
   WScript.Sleep 1000
End If

WScript.Sleep 1000
CheckSapError "Ativar dominio " & domainName

WScript.Echo "Dominio " & domainName & " criado com sucesso."

' ---- Sub para verificar erro na status bar do SAP ----
Sub CheckSapError(stepName)
    Dim sbarType, sbarText
    On Error Resume Next
    sbarType = session.findById("wnd[0]/sbar").MessageType
    sbarText = session.findById("wnd[0]/sbar").Text
    On Error GoTo 0
    If sbarType = "E" Or sbarType = "A" Then
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
        session.findById("wnd[0]").sendVKey 0
        WScript.StdErr.Write "SAP_ERROR: [" & stepName & "] " & sbarText
        WScript.Quit 1
    End If
End Sub
