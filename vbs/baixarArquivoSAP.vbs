' ============================================================
' baixarArquivoSAP.vbs
' Executa programa ABAP Z_DONWLOAD_ARQUIVO que baixa arquivo
' para C:\temp, depois le o conteudo e retorna via stdout.
'
' Argumentos:
'   0: fileName - Nome do arquivo (ex: z_php_test)
'
' Saida:
'   stdout: conteudo do arquivo (capturado pelo ConsumerTT)
' ============================================================
Option Explicit

Dim SapGuiAuto, application, connection, session
Dim fileName

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
    fileName = "z_php_test"
End If

' ---- Executa programa ABAP que baixa o arquivo ----
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "se38"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 500

session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").text = "Z_DONWLOAD_ARQUIVO"
session.findById("wnd[0]").sendVKey 8
WScript.Sleep 500

session.findById("wnd[0]/usr/txtP_PROG").text = fileName
session.findById("wnd[0]/usr/txtP_PROG").caretPosition = Len(fileName)
session.findById("wnd[0]").sendVKey 8

' ---- Aguarda download completar ----
WScript.Sleep 2000

' ---- Agora le o arquivo baixado em C:\temp ----
Dim fso, filePath, fRead, conteudo
Set fso = CreateObject("Scripting.FileSystemObject")
filePath = "C:\temp\" & fileName & ".txt"

' Verifica se arquivo existe
If Not fso.FileExists(filePath) Then
    WScript.StdErr.Write "Arquivo nao encontrado apos download: " & filePath
    WScript.Quit 1
End If

' Le o conteudo
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

' ---- Codifica quebras de linha para transporte ----
conteudo = Replace(conteudo, vbCrLf, "\n")
conteudo = Replace(conteudo, vbCr,   "\n")
conteudo = Replace(conteudo, vbLf,   "\n")

' ---- Retorna conteudo via stdout (capturado pelo ConsumerTT) ----
WScript.Echo conteudo
WScript.Quit 0
