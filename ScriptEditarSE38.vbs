Option Explicit

Dim SapGuiAuto, application, connection, session
Dim programName, codigo

Sub Pause(seconds)
    Dim t
    t = Timer
    Do While Timer < t + seconds
    Loop
End Sub

Function NormalizeLineBreaks(text)
    text = Replace(text, "\r\n", vbCrLf)
    text = Replace(text, "\n", vbCrLf)
    text = Replace(text, "\r", vbCrLf)
    NormalizeLineBreaks = text
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

programName = ""
codigo = ""

If WScript.Arguments.Count >= 1 Then
    programName = Trim(CStr(WScript.Arguments(0)))
End If

If WScript.Arguments.Count >= 2 Then
    codigo = CStr(WScript.Arguments(1))
End If

If programName = "" Then
    programName = "zpar_impar_check"
End If

If codigo = "" Then
    codigo = _
    "REPORT zpar_impar_check." & vbCrLf & _
    "" & vbCrLf & _
    "START-OF-SELECTION." & vbCrLf & _
    "  WRITE: 'deu certo'."
End If

codigo = NormalizeLineBreaks(codigo)

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse38"
session.findById("wnd[0]").sendVKey 0
Pause 1

session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").text = programName
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").caretPosition = Len(programName)
Pause 0.5

session.findById("wnd[0]/usr/btnCHAP").press
Pause 1

session.findById("wnd[0]/usr/cntlEDITOR/shellcont/shell").deleteRange 1, 1, 10000, 10000
Pause 0.5

session.findById("wnd[0]/usr/cntlEDITOR/shellcont/shell").insertText codigo, 1, 1
Pause 0.5

session.findById("wnd[0]/tbar[0]/btn[11]").press
Pause 0.5

session.findById("wnd[0]").sendVKey 27
Pause 1.5

session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]").sendVKey 3
