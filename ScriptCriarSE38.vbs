' ============================================================
' ScriptCriarSE38.vbs
' Cria programa ABAP na transacao SE38 via SAP GUI Scripting
'
' Argumentos posicionais (WScript.Arguments):
'   0: programName  - Nome do programa (ex: ZREPORT_VENDAS)
'   1: packageName  - Pacote ($TMP para local)
'   2: requestId    - Numero da request (vazio se $TMP)
'   3: titulo       - Descricao do programa
'   4-6: reservado
'   7: sourceCode   - Codigo ABAP com \n no lugar de quebra de linha
' ============================================================
Option Explicit

Dim SapGuiAuto, application, connection, session
If Not IsObject(application) Then
    Set SapGuiAuto  = GetObject("SAPGUI")
    Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
    Set connection = application.Children(0)
End If
If Not IsObject(session) Then
    Set session = connection.Children(0)
End If

Dim programName, packageName, requestId, titulo, sourceCode, codigo

programName = "ZTEST_DEFAULT"
packageName = "$TMP"
requestId   = ""
titulo      = ""
sourceCode  = ""

If WScript.Arguments.Count >= 1 Then
    If Trim(CStr(WScript.Arguments(0))) <> "" Then
        programName = UCase(Trim(CStr(WScript.Arguments(0))))
    End If
End If
If WScript.Arguments.Count >= 2 Then
    If Trim(CStr(WScript.Arguments(1))) <> "" Then
        packageName = Trim(CStr(WScript.Arguments(1)))
    End If
End If
If WScript.Arguments.Count >= 3 Then
    requestId = Trim(CStr(WScript.Arguments(2)))
End If
If WScript.Arguments.Count >= 4 Then
    titulo = Trim(CStr(WScript.Arguments(3)))
End If
If WScript.Arguments.Count >= 8 Then
    If Trim(CStr(WScript.Arguments(7))) <> "" Then
        sourceCode = CStr(WScript.Arguments(7))
    End If
End If

' Decodificar \n para quebras de linha reais
If Trim(sourceCode) <> "" Then
    codigo = sourceCode
    codigo = Replace(codigo, "\r\n", vbCrLf)
    codigo = Replace(codigo, "\n",   vbCrLf)
    codigo = Replace(codigo, "\r",   vbCrLf)
Else
    codigo = "REPORT " & LCase(programName) & "." & vbCrLf & vbCrLf & "* TODO: implementar"
End If

' ---- Navegar para SE38 ------------------------------------------
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse38"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 800

' ---- Preencher nome e clicar em Criar ---------------------------
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").text = programName
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").caretPosition = Len(programName)
WScript.Sleep 300
session.findById("wnd[0]/usr/btnNEW").press
WScript.Sleep 800

' ---- Popup de atributos wnd[1] ----------------------------------
session.findById("wnd[1]/usr/txtRS38M-REPTI").text = titulo
session.findById("wnd[1]/usr/cmbTRDIR-SUBC").setFocus
session.findById("wnd[1]/usr/cmbTRDIR-SUBC").key = "1"
session.findById("wnd[1]/usr/cmbTRDIR-RSTAT").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
WScript.Sleep 800

' ---- Popup de pacote wnd[2] -------------------------------------
session.findById("wnd[2]/usr/ctxtKO007-L_DEVCLASS").text = packageName
session.findById("wnd[2]/usr/ctxtKO007-L_DEVCLASS").caretPosition = Len(packageName)
session.findById("wnd[2]/tbar[0]/btn[0]").press
WScript.Sleep 1000

' ---- Request (apenas quando nao for $TMP) -----------------------
If UCase(Trim(packageName)) <> "$TMP" Then
    On Error Resume Next
    session.findById("wnd[3]/usr/ctxtKO008-TRKORR").text = requestId
    session.findById("wnd[3]/usr/ctxtKO008-TRKORR").caretPosition = Len(requestId)
    session.findById("wnd[3]/tbar[0]/btn[0]").press
    Err.Clear
    On Error GoTo 0
    WScript.Sleep 800
End If

' ---- Aguardar o editor ABAP abrir -------------------------------
WScript.Sleep 1000

' ---- Inserir codigo usando setContent (substitui todo o conteudo)
session.findById("wnd[0]/usr/cntlEDITOR/shellcont/shell").setContent(codigo)
WScript.Sleep 500

' ---- Salvar (Ctrl+S) --------------------------------------------
session.findById("wnd[0]/tbar[0]/btn[11]").press
WScript.Sleep 800

' ---- Ativar (tbar[1]/btn[3]) ------------------------------------
On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[3]").press
WScript.Sleep 800

' Fechar popup de ativacao se abrir
If session.findById("wnd[1]").Text <> "" Then
    session.findById("wnd[1]/tbar[0]/btn[0]").press
End If
On Error GoTo 0
WScript.Sleep 400

' ---- Resultado --------------------------------------------------
Dim finalMsg
On Error Resume Next
finalMsg = session.findById("wnd[0]/sbar").Text
On Error GoTo 0

WScript.Echo "Programa " & programName & " criado. Status: " & finalMsg
WScript.Quit 0
