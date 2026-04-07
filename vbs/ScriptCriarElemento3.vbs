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

Dim elementName
Dim elementText
Dim domainName
Dim packageName
Dim requestId
Dim shortText
Dim mediumText
Dim longText
Dim reportText

elementName = "Z_MM_CN7003"
elementText = "elemento de test"
domainName  = "z_mm_cn00"
packageName = "$TMP"
requestId   = ""

' textos padrao
shortText  = Left(elementText, 5)
mediumText = Left(elementText, 10)
longText   = Left(elementText, 10)
reportText = elementText

If WScript.Arguments.Count >= 1 Then
   If Trim(CStr(WScript.Arguments(0))) <> "" Then
      elementName = CStr(WScript.Arguments(0))
   End If
End If

If WScript.Arguments.Count >= 2 Then
   If Trim(CStr(WScript.Arguments(1))) <> "" Then
      elementText = CStr(WScript.Arguments(1))
      shortText  = Left(elementText, 5)
      mediumText = Left(elementText, 10)
      longText   = Left(elementText, 10)
      reportText = elementText
   End If
End If

If WScript.Arguments.Count >= 3 Then
   If Trim(CStr(WScript.Arguments(2))) <> "" Then
      domainName = CStr(WScript.Arguments(2))
   End If
End If

If WScript.Arguments.Count >= 4 Then
   If Trim(CStr(WScript.Arguments(3))) <> "" Then
      packageName = CStr(WScript.Arguments(3))
   End If
End If

If WScript.Arguments.Count >= 5 Then
   If Trim(CStr(WScript.Arguments(4))) <> "" Then
      requestId = CStr(WScript.Arguments(4))
   End If
End If

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse11"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000

session.findById("wnd[0]/usr/radRSRD1-DDTYPE").setFocus
session.findById("wnd[0]/usr/radRSRD1-DDTYPE").select
WScript.Sleep 500

session.findById("wnd[0]/usr/ctxtRSRD1-DDTYPE_VAL").text = elementName
session.findById("wnd[0]/usr/ctxtRSRD1-DDTYPE_VAL").caretPosition = Len(elementName)

' --- Verifica se elemento ja existe (tenta Display) ---
session.findById("wnd[0]/usr/btnPUSHSHOW").press
WScript.Sleep 800

Dim elemExists
elemExists = False
On Error Resume Next
Dim testElemField
testElemField = session.findById("wnd[0]/usr/txtDD04D-DDTEXT").text
If Err.Number = 0 Then elemExists = True
Err.Clear
On Error GoTo 0

If elemExists Then
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.Echo "Elemento " & elementName & " ja existe. Ignorando."
    WScript.Quit 0
End If

' Elemento nao existe — criar
session.findById("wnd[0]/usr/ctxtRSRD1-DDTYPE_VAL").text = elementName
session.findById("wnd[0]/usr/ctxtRSRD1-DDTYPE_VAL").caretPosition = Len(elementName)
session.findById("wnd[0]/usr/btnPUSHADD").press
WScript.Sleep 1000

session.findById("wnd[1]/tbar[0]/btn[0]").press
WScript.Sleep 800

session.findById("wnd[0]/usr/txtDD04D-DDTEXT").text = elementText
session.findById("wnd[0]/usr/tabsTS/tabpTYPE/ssubSUB_DATA:SAPLSD51:1002/ctxtDD04D-DOMNAME").text = domainName
session.findById("wnd[0]/usr/tabsTS/tabpTYPE/ssubSUB_DATA:SAPLSD51:1002/ctxtDD04D-DOMNAME").setFocus
session.findById("wnd[0]/usr/tabsTS/tabpTYPE/ssubSUB_DATA:SAPLSD51:1002/ctxtDD04D-DOMNAME").caretPosition = Len(domainName)
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 800
CheckSapError "Validar dominio " & domainName

session.findById("wnd[0]/usr/tabsTS/tabpADDA").select
WScript.Sleep 300
session.findById("wnd[0]/usr/tabsTS/tabpTEXT").select
WScript.Sleep 800

session.findById("wnd[0]/usr/tabsTS/tabpTEXT/ssubSUB_DATA:SAPLSD51:1003/txtDD04D-SCRTEXT_S").text = shortText
session.findById("wnd[0]/usr/tabsTS/tabpTEXT/ssubSUB_DATA:SAPLSD51:1003/txtDD04D-SCRTEXT_M").text = mediumText
session.findById("wnd[0]/usr/tabsTS/tabpTEXT/ssubSUB_DATA:SAPLSD51:1003/txtDD04D-SCRTEXT_L").text = longText
session.findById("wnd[0]/usr/tabsTS/tabpTEXT/ssubSUB_DATA:SAPLSD51:1003/txtDD04D-REPTEXT").text = reportText
session.findById("wnd[0]/usr/tabsTS/tabpTEXT/ssubSUB_DATA:SAPLSD51:1003/txtDD04D-REPTEXT").setFocus
session.findById("wnd[0]/usr/tabsTS/tabpTEXT/ssubSUB_DATA:SAPLSD51:1003/txtDD04D-REPTEXT").caretPosition = Len(reportText)
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 800

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
CheckActivationPopup "Ativar elemento " & elementName
CheckSapError "Ativar elemento " & elementName

WScript.Echo "Elemento " & elementName & " criado com sucesso."

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

' ---- Sub para verificar popup de erro de ativacao (wnd[1]) ----
Sub CheckActivationPopup(stepName)
    Dim popup
    On Error Resume Next
    Set popup = session.findById("wnd[1]")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    If popup Is Nothing Then Exit Sub

    Dim popupTitle
    On Error Resume Next
    popupTitle = popup.Text
    On Error GoTo 0

    Dim lowerTitle
    lowerTitle = LCase(popupTitle)
    If InStr(lowerTitle, "erro") = 0 And InStr(lowerTitle, "error") = 0 _
       And InStr(lowerTitle, "ativ") = 0 And InStr(lowerTitle, "syntax") = 0 Then
        Exit Sub
    End If

    ' Tenta clicar em Nao/Processar para ver detalhes
    On Error Resume Next
    session.findById("wnd[1]/usr/btnBUTTON_2").press
    If Err.Number <> 0 Then
        Err.Clear
        popup.findById("tbar[0]/btn[2]").press
    End If
    On Error GoTo 0
    WScript.Sleep 1500

    ' Tenta ler erros detalhados do grid
    Dim shell, rowCount, i, errText, fullError
    fullError = ""
    On Error Resume Next
    Set shell = session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell")
    If Not shell Is Nothing Then
        rowCount = shell.RowCount
        Dim lineNr
        For i = 0 To rowCount - 1
            errText = shell.GetCellValue(i, "TEXT")
            If errText <> "" Then
                lineNr = shell.GetCellValue(i, "LINE")
                If lineNr <> "" Then
                    If fullError <> "" Then fullError = fullError & " | "
                    fullError = fullError & "Linha " & lineNr & ": " & errText
                Else
                    fullError = fullError & " " & errText
                End If
            End If
        Next
    End If
    On Error GoTo 0

    If fullError = "" Then fullError = popupTitle

    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0

    WScript.StdErr.Write "SAP_ERROR: [" & stepName & "] " & fullError
    WScript.Quit 1
End Sub
