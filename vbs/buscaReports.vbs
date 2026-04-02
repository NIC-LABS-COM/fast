' ============================================================
' buscaReports.vbs
' Executa programa ABAP Z_BUSCA_REPORTS que gera arquivo
' em C:\temp, depois le o conteudo e retorna via stdout.
'
' Argumentos:
'   (nenhum obrigatorio)
'
' Saida:
'   stdout: conteudo do arquivo (capturado pelo ConsumerTT)
' ============================================================
Option Explicit

Dim SapGuiAuto, application, connection, session
Dim i, fso, filePath, fRead, conteudo

Const REPORT_NAME = "Z_BUSCA_REPORTS"
Const FILE_PATH   = "C:\temp\reports.txt"

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
CheckSapError "Abrir report"

' ---- Preenche parametro (opcional) ----
On Error Resume Next
session.findById("wnd[0]/usr/txtP_PROG").Text = "reports"
Err.Clear
On Error GoTo 0

' ---- Executa (F8) ----
session.findById("wnd[0]").sendVKey 8
WScript.Sleep 3000
CheckSapError "Executar report"

' Se aparecer popup, tenta fechar
On Error Resume Next
session.findById("wnd[1]/tbar[0]/btn[0]").press
Err.Clear
On Error GoTo 0

WScript.Sleep 2000

' ---- Aguarda o arquivo ser gerado (ate 10 segundos) ----
For i = 1 To 10
    If fso.FileExists(FILE_PATH) Then
        Exit For
    End If
    WScript.Sleep 1000
Next

If Not fso.FileExists(FILE_PATH) Then
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.StdErr.Write "Arquivo nao encontrado apos execucao: " & FILE_PATH
    WScript.Quit 1
End If

' ---- Le o conteudo do arquivo ----
On Error Resume Next
Set fRead = fso.OpenTextFile(FILE_PATH, 1, False)
If Err.Number <> 0 Then
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.StdErr.Write "Erro ao abrir arquivo: " & FILE_PATH & " | " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0

conteudo = fRead.ReadAll
fRead.Close

If Trim(conteudo) = "" Then
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.StdErr.Write "Arquivo encontrado mas esta vazio: " & FILE_PATH
    WScript.Quit 1
End If

' ---- Cleanup e output ----
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0

conteudo = Replace(conteudo, vbCrLf, "\n")
conteudo = Replace(conteudo, vbCr, "\n")
conteudo = Replace(conteudo, vbLf, "\n")

WScript.Echo conteudo
WScript.Quit 0
