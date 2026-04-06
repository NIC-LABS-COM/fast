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

    ' Le titulo do popup
    Dim popupTitle
    On Error Resume Next
    popupTitle = popup.Text
    On Error GoTo 0

    ' Verifica se eh popup de erro de ativacao
    Dim lowerTitle
    lowerTitle = LCase(popupTitle)
    If InStr(lowerTitle, "erro") = 0 And InStr(lowerTitle, "error") = 0 _
       And InStr(lowerTitle, "ativ") = 0 And InStr(lowerTitle, "syntax") = 0 Then
        Exit Sub
    End If

    ' Clica em "Nao" / "Processar" (btnBUTTON_2) para mostrar lista de erros
    On Error Resume Next
    session.findById("wnd[1]/usr/btnBUTTON_2").press
    If Err.Number <> 0 Then
        Err.Clear
        popup.findById("tbar[0]/btn[2]").press
    End If
    On Error GoTo 0
    Pause 1.5

    ' Le erros detalhados do grid de sintaxe
    Dim shell, rowCount, i, errText, fullError
    fullError = ""
    On Error Resume Next
    Set shell = session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell")
    If Not shell Is Nothing Then
        rowCount = shell.RowCount
        For i = 0 To rowCount - 1
            errText = shell.GetCellValue(i, "TEXT")
            If errText <> "" Then
                If fullError <> "" Then fullError = fullError & " | "
                fullError = fullError & "Linha " & shell.GetCellValue(i, "LINE") & ": " & errText
            End If
        Next
    End If
    On Error GoTo 0

    If fullError = "" Then fullError = popupTitle

    ' Volta para tela inicial
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0

    WScript.StdErr.Write "SAP_ERROR: [" & stepName & "] " & fullError
    WScript.Quit 1
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
