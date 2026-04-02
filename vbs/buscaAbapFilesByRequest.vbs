' ============================================================
' buscaAbapFilesByRequest.vbs
' Executa programa ABAP Z_GET_ABAP_FILES_BY_REQUEST que gera
' arquivo em C:\temp, depois le o conteudo e retorna via stdout.
'
' Argumentos:
'   0: requests - Lista de requests separadas por virgula
'                 Ex: "USDK900001,USDK900002"
'
' Saida:
'   stdout: conteudo do arquivo gerado pelo report
' ============================================================
Option Explicit

Dim SapGuiAuto, application, connection, session
Dim i, fso, filePath, fRead, conteudo
Dim argRequests, arrRequests

Const REPORT_NAME = "Z_GET_ABAP_FILES_BY_REQUEST"
Const FILE_PATH   = "C:\temp\request_files.txt"

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

' ---- Leitura dos argumentos ----
argRequests = ""

If WScript.Arguments.Count >= 1 Then
    argRequests = Trim(CStr(WScript.Arguments(0)))
End If

If argRequests = "" Then
    WScript.StdErr.Write "Argumento requests e obrigatorio."
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
CheckSapError "Abrir report"

' ---- Preenche requests via SELECT-OPTIONS (multipla selecao) ----
arrRequests = Split(argRequests, ",")

' Preenche o primeiro valor no campo LOW
On Error Resume Next
session.findById("wnd[0]/usr/ctxtP_REQ-LOW").Text = Trim(arrRequests(0))
Err.Clear
On Error GoTo 0

' Se houver mais valores, usa botao de multipla selecao
If UBound(arrRequests) > 0 Then
    On Error Resume Next
    ' Clica no botao de multipla selecao (icone amarelo ao lado do campo)
    session.findById("wnd[0]/usr/btn%_P_REQ_%_APP_%-VALU_PUSH").press
    WScript.Sleep 500

    Dim j
    For j = 1 To UBound(arrRequests)
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & (j - 1) & "]").Text = Trim(arrRequests(j))
    Next

    ' Confirma a selecao multipla
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    WScript.Sleep 500
    Err.Clear
    On Error GoTo 0
End If

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

' ---- Aguarda o arquivo ser gerado (ate 15 segundos) ----
filePath = FILE_PATH

For i = 1 To 15
    If fso.FileExists(filePath) Then
        Exit For
    End If
    WScript.Sleep 1000
Next

If Not fso.FileExists(filePath) Then
    ' Volta para tela inicial antes de sair
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

conteudo = Trim(conteudo)

' ---- Volta para tela inicial SAP ----
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0

WScript.Echo conteudo
WScript.Quit 0
