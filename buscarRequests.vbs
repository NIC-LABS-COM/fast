Option Explicit

Dim SapGuiAuto, application, connection, session
Dim fileName

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
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").text = "Z_BUSCA_REQUESTS"
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/txtP_PROG").text = fileName
session.findById("wnd[0]/usr/txtP_PROG").caretPosition = Len(fileName)
session.findById("wnd[0]").sendVKey 8
