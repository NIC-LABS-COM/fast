' ============================================================
' buscaConteudoPorVersao.vbs
' Executa programa ABAP Z_GET_ARQUIVO_POR_VERSAO que gera arquivo
' em C:\temp, depois le o conteudo e retorna via stdout.
'
' Argumentos:
'   0: fileName  - Nome do programa ABAP
'   1: category  - Categoria (PROGRAM, FUNCTION_MODULE, CLASS)
'   2: versionId - ID da versao
'
' Saida:
'   stdout: conteudo do arquivo (capturado pelo ConsumerTT)
' ============================================================
Option Explicit

Dim SapGuiAuto, application, connection, session
Dim i, fso, filePath, fRead, conteudo
Dim argFileName, argCategory, argVersionId

Const REPORT_NAME = "Z_GET_ARQUIVO_POR_VERSAO"
Const FILE_PATH   = "C:\temp\arquivo_por_versao.txt"

' ---- Leitura dos argumentos ----
argFileName  = ""
argCategory  = ""
argVersionId = ""

If WScript.Arguments.Count >= 1 Then
    argFileName = Trim(CStr(WScript.Arguments(0)))
End If

If WScript.Arguments.Count >= 2 Then
    argCategory = Trim(CStr(WScript.Arguments(1)))
End If

If WScript.Arguments.Count >= 3 Then
    argVersionId = Trim(CStr(WScript.Arguments(2)))
End If

If argFileName = "" Then
    WScript.StdErr.Write "Argumento fileName e obrigatorio."
    WScript.Quit 1
End If

If argVersionId = "" Then
    WScript.StdErr.Write "Argumento versionId e obrigatorio."
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

' ---- Preenche parametros do report ----
On Error Resume Next
session.findById("wnd[0]/usr/txtP_FNAME").Text = argFileName
Err.Clear
If argCategory <> "" Then
    session.findById("wnd[0]/usr/txtP_CAT").Text = argCategory
    Err.Clear
End If
session.findById("wnd[0]/usr/txtP_VERS").Text = argVersionId
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

conteudo = Replace(conteudo, vbCrLf, "\n")
conteudo = Replace(conteudo, vbCr, "\n")
conteudo = Replace(conteudo, vbLf, "\n")

WScript.Echo conteudo
WScript.Quit 0
