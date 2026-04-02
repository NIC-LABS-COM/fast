Option Explicit

Dim SapGuiAuto, application, connection, session
Dim programName, codigo

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

' ---- Sub para verificar popup de erro de ativação (wnd[1]) ----
Sub CheckActivationPopup(stepName)
    Dim popupWnd, msgText, popupTitle
    On Error Resume Next
    Set popupWnd = session.findById("wnd[1]")
    On Error GoTo 0
    If popupWnd Is Nothing Then Exit Sub

    msgText = ""
    On Error Resume Next
    msgText = session.findById("wnd[1]/usr/txtMESSTXT1").Text
    On Error GoTo 0
    If msgText = "" Then
        On Error Resume Next
        msgText = session.findById("wnd[1]/usr/txtSPOPLI-TEXTLINE1").Text
        On Error GoTo 0
    End If
    If msgText = "" Then
        On Error Resume Next
        popupTitle = popupWnd.Text
        On Error GoTo 0
        If popupTitle <> "" Then msgText = popupTitle
    End If

    Dim lower
    lower = LCase(msgText)
    If InStr(lower, "erro") > 0 Or InStr(lower, "error") > 0 _
       Or InStr(lower, "falta") > 0 Or InStr(lower, "sintaxe") > 0 _
       Or InStr(lower, "syntax") > 0 Or InStr(lower, "encerrado") > 0 Then

        On Error Resume Next
        session.findById("wnd[1]/tbar[0]/btn[2]").press
        On Error GoTo 0
        Pause 0.5

        session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
        session.findById("wnd[0]").sendVKey 0
        WScript.StdErr.Write "SAP_ERROR: [" & stepName & "] " & msgText
        WScript.Quit 1
    End If
End Sub

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
CheckSapError "Salvar programa"

session.findById("wnd[0]").sendVKey 27
Pause 1.5
CheckActivationPopup "Ativar programa"
CheckSapError "Ativar programa"

WScript.Echo "Programa " & programName & " editado com sucesso."

session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]").sendVKey 3
