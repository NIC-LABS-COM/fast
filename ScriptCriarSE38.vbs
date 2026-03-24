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
Dim progType
Dim flow
Dim sourceCode
Dim codigo

programName = "ZMM_TESTE_PARIMPAR"
packageName = "$TMP"
requestId = ""
titulo = ""
progType = "1"
flow = ""
sourceCode = ""

' Ordem real dos args publicados hoje:
' 0 = programName
' 1 = packageName
' 2 = requestId
' 3 = titulo
' 4 = progType
' 5 = ignorado
' 6 = flow
' 7 = sourceCode

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

If WScript.Arguments.Count >= 5 Then
   If Trim(CStr(WScript.Arguments(4))) <> "" Then
      progType = CStr(WScript.Arguments(4))
   End If
End If

If WScript.Arguments.Count >= 7 Then
   If Trim(CStr(WScript.Arguments(6))) <> "" Then
      flow = CStr(WScript.Arguments(6))
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

' Normaliza quebras de linha vindas do backend para o formato do editor SAP
codigo = Replace(codigo, "\r\n", vbCr)
codigo = Replace(codigo, "\n", vbCr)
codigo = Replace(codigo, "\r", vbCr)

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse38"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 800

session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").text = programName
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").caretPosition = Len(programName)
WScript.Sleep 300

' Criar / New
session.findById("wnd[0]/usr/btnNEW").press
WScript.Sleep 800

' Popup de criação do programa
session.findById("wnd[1]/usr/txtRS38M-REPTI").text = programName
session.findById("wnd[1]/usr/cmbTRDIR-SUBC").setFocus
session.findById("wnd[1]/usr/cmbTRDIR-SUBC").key = progType
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
session.findById("wnd[0]/usr/cntlEDITOR/shellcont/shell").text = "" + vbCr + ""
session.findById("wnd[0]/usr/cntlEDITOR/shellcont/shell").setSelectionIndexes 0,0
WScript.Sleep 500

' Escreve o código no editor
session.findById("wnd[0]/usr/cntlEDITOR/shellcont/shell").text = codigo
WScript.Sleep 500

' Salvar
session.findById("wnd[0]/tbar[0]/btn[11]").press
WScript.Sleep 500

' Ativar conforme fluxo
flow = UCase(Trim(CStr(flow)))
If flow = "CRIAR E ATIVAR" Or flow = "ATIVAR" Then
   session.findById("wnd[0]").sendVKey 27
End If
