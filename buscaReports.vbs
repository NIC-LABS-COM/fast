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

Dim fileName
fileName = "reports"

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse38"
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").text = "z_busca_reports"
session.findById("wnd[0]").sendVKey 8

' ---- Aguarda download completar ----
WScript.Sleep 2000

' ---- Agora le o arquivo baixado em C:\temp ----
Dim fso, filePath, fRead, conteudo
Set fso = CreateObject("Scripting.FileSystemObject")

filePath = "C:\temp\" & fileName & ".txt"

If Not fso.FileExists(filePath) Then
    WScript.StdErr.Write "Arquivo nao encontrado apos download: " & filePath
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
