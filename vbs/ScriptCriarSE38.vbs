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

Dim programName
Dim packageName
Dim requestId
Dim titulo
Dim sourceCode
Dim codigo

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
    WScript.Sleep 1500

    ' Le erros detalhados do grid de sintaxe
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

    ' Volta para tela inicial
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0

    WScript.StdErr.Write "SAP_ERROR: [" & stepName & "] " & fullError
    WScript.Quit 1
End Sub

programName = "ZMM_TESTE_PARIMPAR"
packageName = "$TMP"
requestId = ""
titulo = ""
sourceCode = ""

If WScript.Arguments.Count >= 1 Then
   If Trim(CStr(WScript.Arguments(0))) <> "" Then
      programName = CStr(WScript.Arguments(0))
   End If
End If

If WScript.Arguments.Count >= 2 Then
   If Trim(CStr(WScript.Arguments(1))) <> "" Then
      packageName = CStr(WScript.Arguments(1))
   End If
End If

If WScript.Arguments.Count >= 3 Then
   If Trim(CStr(WScript.Arguments(2))) <> "" Then
      requestId = CStr(WScript.Arguments(2))
   End If
End If

If WScript.Arguments.Count >= 4 Then
   If Trim(CStr(WScript.Arguments(3))) <> "" Then
      titulo = CStr(WScript.Arguments(3))
   End If
End If

If WScript.Arguments.Count >= 8 Then
   If Trim(CStr(WScript.Arguments(7))) <> "" Then
      sourceCode = CStr(WScript.Arguments(7))
   End If
End If

If Trim(CStr(sourceCode)) <> "" Then
   codigo = sourceCode
Else
   codigo = _
   "REPORT " & LCase(programName) & "." & vbCrLf & _
   "" & vbCrLf & _
   "DATA lv_num TYPE i VALUE 5." & vbCrLf & _
   "" & vbCrLf & _
   "IF lv_num MOD 2 = 0." & vbCrLf & _
   "  WRITE: / 'Numero par'." & vbCrLf & _
   "ELSE." & vbCrLf & _
   "  WRITE: / 'Numero impar'." & vbCrLf & _
   "ENDIF."
End If

' Normaliza quebras de linha caso o backend envie \n como texto
codigo = Replace(codigo, "\r\n", vbCrLf)
codigo = Replace(codigo, "\n", vbCrLf)
codigo = Replace(codigo, "\r", vbCrLf)

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse38"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 800

session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").text = programName
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").caretPosition = Len(programName)
WScript.Sleep 300

' Clicar em Criar / New
session.findById("wnd[0]/usr/btnNEW").press
WScript.Sleep 800

' Popup de criação do programa
session.findById("wnd[1]/usr/txtRS38M-REPTI").text = programName
session.findById("wnd[1]/usr/cmbTRDIR-SUBC").setFocus
session.findById("wnd[1]/usr/cmbTRDIR-SUBC").key = "1"
session.findById("wnd[1]/usr/cmbTRDIR-RSTAT").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
WScript.Sleep 800

' Popup do pacote
session.findById("wnd[2]/usr/ctxtKO007-L_DEVCLASS").text = packageName
session.findById("wnd[2]/usr/ctxtKO007-L_DEVCLASS").caretPosition = Len(packageName)
session.findById("wnd[2]").sendVKey 0
WScript.Sleep 1000

' [ALTERADO] Gravacao direta na REQUEST recebida (persist definitivo)
' Antes: gravava na request somente quando packageName <> "$TMP"
' Agora: grava diretamente na request se requestId for fornecido
If requestId <> "" Then
   On Error Resume Next
   session.findById("wnd[3]/usr/ctxtKO008-TRKORR").text = requestId
   session.findById("wnd[3]/usr/ctxtKO008-TRKORR").caretPosition = Len(requestId)
   session.findById("wnd[3]").sendVKey 0
   Err.Clear
   On Error GoTo 0
   WScript.Sleep 800
End If

' Espera o editor abrir
WScript.Sleep 1000

' Apaga o conteúdo padrão gerado pelo SAP
session.findById("wnd[0]/usr/cntlEDITOR/shellcont/shell").deleteRange 1,1,9,1
WScript.Sleep 500

' Escreve o código no editor
session.findById("wnd[0]/usr/cntlEDITOR/shellcont/shell").insertText codigo, 1, 1
WScript.Sleep 500

' Salvar
session.findById("wnd[0]/tbar[0]/btn[11]").press
WScript.Sleep 500
CheckSapError "Salvar programa"

' Ativar
session.findById("wnd[0]").sendVKey 27
WScript.Sleep 1500

' Se surgir popup de confirmacao (nao de erro), confirma com Enter
On Error Resume Next
Dim wndConfirm
Set wndConfirm = session.findById("wnd[1]")
If Not wndConfirm Is Nothing Then
    Dim wndConfirmTitle
    wndConfirmTitle = LCase(wndConfirm.Text)
    If InStr(wndConfirmTitle, "erro") = 0 And InStr(wndConfirmTitle, "error") = 0 _
       And InStr(wndConfirmTitle, "syntax") = 0 Then
        wndConfirm.sendVKey 0
        WScript.Sleep 800
    End If
End If
On Error GoTo 0

CheckActivationPopup "Ativar programa"
CheckSapError "Ativar programa"

' Volta para a tela inicial do SAP
session.findById("wnd[0]/tbar[0]/okcd").text = "/n12365"
session.findById("wnd[0]").sendVKey 0

WScript.Echo "Programa " & programName & " criado com sucesso."
