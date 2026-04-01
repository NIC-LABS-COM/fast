' ============================================================
' buscaCategoryByFileName.vbs
' Executa programa ABAP Z_GET_CATEGORY_BY_FILE_NAME que gera
' arquivo em C:\temp, depois le o conteudo e retorna via stdout.
'
' Argumentos:
'   0: fileName - Nome do objeto ABAP
'
' Saida:
'   stdout: categoria do objeto (ex: PROGRAM, CLASS, FUNC...)
' ============================================================
Option Explicit

Dim SapGuiAuto, application, connection, session
Dim i, fso, filePath, fRead, conteudo
Dim argFileName

Const REPORT_NAME = "Z_GET_CATEGORY_BY_FILE_NAME"
Const FILE_PATH   = "C:\temp\file_category.txt"

' ---- Leitura dos argumentos ----
argFileName = ""

If WScript.Arguments.Count >= 1 Then
    argFileName = Trim(CStr(WScript.Arguments(0)))
End If

If argFileName = "" Then
    WScript.StdErr.Write "Argumento fileName e obrigatorio."
    WScript.Quit 1
End If

' ---- Conexao SAP GUI ----
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

session.findById("wnd[0]").maximize

' ---- Vai para SE38 ----
session.findById("wnd[0]/tbar[0]/okcd").Text = "/NSE38"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000

' ---- Informa o report ----
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = REPORT_NAME
session.findById("wnd[0]").sendVKey 8
WScript.Sleep 1000

' ---- Preenche parametro fileName ----
On Error Resume Next
session.findById("wnd[0]/usr/txtP_FILE").Text = argFileName
Err.Clear
On Error GoTo 0

' ---- Executa (F8) ----
session.findById("wnd[0]").sendVKey 8
WScript.Sleep 3000

' Se aparecer popup, tenta fechar
On Error Resume Next
session.findById("wnd[1]/tbar[0]/btn[0]").press
Err.Clear
On Error GoTo 0

WScript.Sleep 2000

' ---- Aguarda o arquivo ser gerado (ate 10 segundos) ----
Set fso = CreateObject("Scripting.FileSystemObject")
filePath = FILE_PATH

For i = 1 To 10
    If fso.FileExists(filePath) Then
        Exit For
    End If
    WScript.Sleep 1000
Next

If Not fso.FileExists(filePath) Then
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.StdErr.Write "Arquivo nao encontrado apos execucao: " & filePath
    WScript.Quit 1
End If

' ---- Le o conteudo do arquivo ----
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

conteudo = Trim(conteudo)

WScript.Echo conteudo
WScript.Quit 0
