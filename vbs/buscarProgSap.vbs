Option Explicit

Dim SapGuiAuto, application, connection, session
Dim fileName

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

fileName = ""

If WScript.Arguments.Count >= 1 Then
    fileName = Trim(CStr(WScript.Arguments(0)))
End If

If fileName = "" Then
    fileName = "z_php_test"
End If

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "se38"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").text = "Z_DONWLOAD_ARQUIVO"
session.findById("wnd[0]").sendVKey 8
CheckSapError "Abrir report"
session.findById("wnd[0]/usr/txtP_PROG").text = fileName
session.findById("wnd[0]/usr/txtP_PROG").caretPosition = Len(fileName)
session.findById("wnd[0]").sendVKey 8
CheckSapError "Executar report"
