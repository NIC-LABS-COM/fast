' ============================================================
' buscarRequests.vbs
' Executa programa ABAP Z_BUSCA_REQUEST que gera arquivo
' em C:\temp, depois le o conteudo e retorna via stdout.
'
' Argumentos:
'   0: fileName - Nome do arquivo sem extensao (default: workbench_requests)
'
' Saida:
'   stdout: conteudo do arquivo (capturado pelo ConsumerTT)
' ============================================================
Option Explicit

Dim SapGuiAuto, application, connection, session
Dim fileName
Dim fso, filePath, fRead, conteudo

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
    fileName = "workbench_requests"
End If

' ---- Executa programa ABAP na SE38 ----
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "se38"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 500

session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "Z_BUSCA_REQUEST"
session.findById("wnd[0]").sendVKey 8
WScript.Sleep 500

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
    WScript.StdErr.Write "Arquivo nao encontrado apos execucao: " & filePath
    WScript.Quit 1
End If

On Error Resume Next
Set fRead = fso.OpenTextFile(filePath, 1, False)
If Err.Number <> 0 Then
    WScript.StdErr.Write "Erro ao abrir arquivo: " & filePath & " | " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0

conteudo = fRead.ReadAll
fRead.Close

If Trim(conteudo) = "" Then
    WScript.StdErr.Write "Arquivo encontrado mas esta vazio: " & filePath
    WScript.Quit 1
End If

conteudo = Replace(conteudo, vbCrLf, "\n")
conteudo = Replace(conteudo, vbCr, "\n")
conteudo = Replace(conteudo, vbLf, "\n")

WScript.Echo conteudo
WScript.Quit 0
