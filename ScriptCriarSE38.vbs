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
Dim codigo

programName = "ZMM_TESTE_PARIMPAR"
packageName = "$TMP"

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
