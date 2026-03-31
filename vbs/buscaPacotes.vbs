' ============================================================
' buscaPackages.vbs
' Executa programa ABAP Z_BUSCA_PACKAGES que gera arquivo
' em C:\temp\packages_tdevc.txt, depois le o conteudo
' e retorna via stdout.
' ============================================================
Option Explicit

Dim SapGuiAuto, application, connection, session
Dim fso, filePath, fRead, conteudo

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children(0)
Set session = connection.Children(0)

session.findById("wnd[0]").maximize

' ---- Vai para SE38 ----
session.findById("wnd[0]/tbar[0]/okcd").Text = "/NSE38"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000

' ---- Informa o programa ----
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = "Z_BUSCA_PACKAGES"
session.findById("wnd[0]").sendVKey 8
WScript.Sleep 3000

' ---- Caso apareça popup/mensagem, tenta fechar ----
On Error Resume Next
If session.Children.Count > 1 Then
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    WScript.Sleep 500
End If
Err.Clear
On Error GoTo 0

' ---- Le o arquivo gerado ----
Set fso = CreateObject("Scripting.FileSystemObject")
filePath = "C:\temp\packages_tdevc.txt"

If Not fso.FileExists(filePath) Then
    WScript.StdErr.Write "Arquivo nao encontrado apos execucao do report: " & filePath
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
