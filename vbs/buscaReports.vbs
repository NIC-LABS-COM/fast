' ============================================================
' buscaReports.vbs
' Executa programa ABAP Z_BUSCA_REPORTS que gera arquivo
' em C:\temp, depois le o conteudo e retorna via stdout.
'
' Argumentos:
'   0: fileName - Nome do arquivo sem extensao (default: reports)
'
' Saida:
'   stdout: conteudo do arquivo (capturado pelo ConsumerTT)
' ============================================================
Option Explicit

Dim SapGuiAuto, application, connection, session
Dim fileName
Dim fso, filePath, fRead, conteudo

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
    fileName = "reports"
End If

' ---- Executa programa ABAP na SE38 ----
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "se38"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 500

session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "Z_BUSCA_REPORTS"
session.findById("wnd[0]").sendVKey 8
WScript.Sleep 500
CheckSapError "Abrir report"

' Tenta setar P_PROG e executar; se nao existir, programa ja rodou com defaults
On Error Resume Next
session.findById("wnd[0]/usr/txtP_PROG").Text = fileName
If Err.Number = 0 Then
    session.findById("wnd[0]/usr/txtP_PROG").caretPosition = Len(fileName)
    session.findById("wnd[0]").sendVKey 8
End If
Err.Clear
On Error GoTo 0

' ---- Aguarda o txt ser gerado ----
WScript.Sleep 2000

' ---- Le o arquivo em C:\temp ----
Set fso = CreateObject("Scripting.FileSystemObject")
filePath = "C:\temp\" & fileName & ".txt"

If Not fso.FileExists(filePath) Then
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.StdErr.Write "Arquivo nao encontrado apos execucao do report: " & filePath
    WScript.Quit 1
End If

On Error Resume Next
Set fRead = fso.OpenTextFile(filePath, 1, False)
If Err.Number <> 0 Then
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.StdErr.Write "Erro ao abrir arquivo: " & filePath & " | " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0

conteudo = fRead.ReadAll
fRead.Close

If Trim(conteudo) = "" Then
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.StdErr.Write "Arquivo encontrado mas esta vazio: " & filePath
    WScript.Quit 1
End If

' Volta para tela inicial SAP
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0

conteudo = Replace(conteudo, vbCrLf, "\n")
conteudo = Replace(conteudo, vbCr, "\n")
conteudo = Replace(conteudo, vbLf, "\n")

WScript.Echo conteudo
WScript.Quit 0
