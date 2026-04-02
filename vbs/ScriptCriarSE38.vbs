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
session.findById("wnd[2]/tbar[0]/btn[0]").press
WScript.Sleep 1000

' Se nao for $TMP e vier request, tenta preencher
If UCase(Trim(packageName)) <> "$TMP" Then
   On Error Resume Next
   session.findById("wnd[3]/usr/ctxtKO008-TRKORR").text = requestId
   session.findById("wnd[3]/usr/ctxtKO008-TRKORR").caretPosition = Len(requestId)
   session.findById("wnd[3]/tbar[0]/btn[0]").press
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
