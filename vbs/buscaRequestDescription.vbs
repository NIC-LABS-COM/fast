' ============================================================
' buscaRequestDescription.vbs
' Executa programa ABAP Z_GET_REQUEST_DESCRIPTION que gera
' arquivo em C:\temp, depois le o conteudo e retorna via stdout.
'
' Argumentos:
'   0: requestId - ID da request (ex: USDK900001)
'
' Saida:
'   stdout: descricao da request (string simples)
' ============================================================
Option Explicit

Dim SapGuiAuto, application, connection, session
Dim i, fso, filePath, fRead, conteudo
Dim argRequestId

Const REPORT_NAME = "Z_GET_REQUEST_DESCRIPTION"
Const FILE_PATH   = "C:\temp\request_description.txt"

' ---- Leitura dos argumentos ----
argRequestId = ""

If WScript.Arguments.Count >= 1 Then
    argRequestId = Trim(CStr(WScript.Arguments(0)))
End If

If argRequestId = "" Then
    WScript.StdErr.Write "Argumento requestId e obrigatorio."
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

' ---- Deleta arquivo anterior se existir ----
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(FILE_PATH) Then
    fso.DeleteFile FILE_PATH
End If

' ---- Vai para SE38 ----
session.findById("wnd[0]/tbar[0]/okcd").Text = "/NSE38"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000

' ---- Informa o report ----
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = REPORT_NAME
session.findById("wnd[0]").sendVKey 8
WScript.Sleep 1000

' ---- Preenche parametro requestId ----
Dim field
Set field = session.findById("wnd[0]/usr/ctxtP_REQ")

field.SetFocus
field.Text = argRequestId
field.caretPosition = Len(argRequestId)

WScript.Sleep 500
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 500

' ---- Valida se o campo foi preenchido ----
If Trim(field.Text) = "" Then
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.StdErr.Write "Campo P_REQ nao foi preenchido no SAP"
    WScript.Quit 1
End If

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
