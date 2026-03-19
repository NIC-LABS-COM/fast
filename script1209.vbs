Option Explicit

Dim SapGuiAuto
Dim application
Dim connection
Dim session
Dim shell

' Função para logar o campo e ID de um campo específico
Sub logCampoID(campoID)
    On Error Resume Next
    Dim campo
    Set campo = session.findById(campoID)
    If Not campo Is Nothing Then
        Log "Campo ID: " & campoID & " encontrado com valor: " & campo.text
    Else
        Log "Campo ID: " & campoID & " não encontrado."
    End If
    On Error GoTo 0
End Sub

Function WndExists(sessionObj, wndId)
    On Error Resume Next
    Dim wnd
    Set wnd = sessionObj.findById(wndId)
    WndExists = (Err.Number = 0)
    Set wnd = Nothing
    Err.Clear
    On Error GoTo 0
End Function

Sub StopScript(msg)
    WScript.Echo "ERRO: " & msg
    WScript.Quit 1
End Sub

'========================================================
' Conecta na sessão SAP já aberta
'========================================================
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

If session Is Nothing Then
    StopScript "Não foi possível obter a sessão SAP."
End If

session.findById("wnd[0]").maximize
Log "Sessão SAP conectada e maximizada"

'========================================================
' Vai para a SE16N
'========================================================
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse16n"
session.findById("wnd[0]").sendVKey 0
Log "Navegou para SE16N"

WScript.Sleep 800

'========================================================
' Informa a tabela E070
'========================================================
If WndExists(session, "wnd[0]/usr/ctxtGD-TAB") Then
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "E070"
    session.findById("wnd[0]").sendVKey 0
    Log "Preencheu campo Tabela E070"
Else
    StopScript "Campo da tabela GD-TAB não encontrado."
End If

WScript.Sleep 800

'========================================================
' Remove limite máximo de linhas
'========================================================
If WndExists(session, "wnd[0]/usr/txtGD-MAX_LINES") Then
    session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = ""
End If

'========================================================
' Preenche critérios de seleção
' Pela sua tela:
' linha 0 = TRKORR
' linha 2 = TRSTATUS
'========================================================
logCampoID("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SE16N_SELFIELDS-LOW[2,0]") ' TRKORR
logCampoID("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SE16N_SELFIELDS-LOW[2,2]") ' TRSTATUS

If WndExists(session, "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,0]") Then
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,0]").setFocus
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,0]").text = "DDR*"
    Log "Preencheu TRKORR com DDR* na linha 2, coluna 0"
Else
    StopScript "Campo LOW da linha TRKORR na coluna 0 não encontrado."
End If

If WndExists(session, "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,2]") Then
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS__SELFIELDS-LOW[2,2]").setFocus
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,2]").text = "D"
    Log "Preencheu TRSTATUS com D na linha 2, coluna 2"
Else
    StopScript "Campo LOW da linha TRSTATUS na coluna 2 não encontrado."
End If

WScript.Sleep 500

'========================================================
' Executa
'========================================================
If WndExists(session, "wnd[0]/tbar[1]/btn[8]") Then
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    Log "Executou a seleção"
Else
    StopScript "Botão Executar não encontrado."
End If

WScript.Sleep 1200

'========================================================
' Obtém ALV de resultado
'========================================================
If WndExists(session, "wnd[0]/usr/cntlRESULT_LIST/shellcont/shell") Then
    Set shell = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
    Log "ALV de resultados encontrado"
Else
    StopScript "ALV de resultado não encontrado."
End If

'========================================================
' Lê as 3 primeiras ordens
'========================================================
Dim ordem1, ordem2, ordem3
ordem1 = ""
ordem2 = ""
ordem3 = ""

On Error Resume Next
ordem1 = shell.getCellValue(0, "TRKORR")
ordem2 = shell.getCellValue(1, "TRKORR")
ordem3 = shell.getCellValue(2, "TRKORR")
On Error GoTo 0

Log "Ordens encontradas: " & ordem1 & ", " & ordem2 & ", " & ordem3
WScript.Echo "Ordens encontradas: " & ordem1 & ", " & ordem2 & ", " & ordem3

' Função de log para acompanhamento
Sub Log(msg)
    WScript.Echo "[LOG] " & msg
End Sub
