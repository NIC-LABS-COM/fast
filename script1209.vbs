' If Not IsObject(application) Then
'    Set SapGuiAuto  = GetObject("SAPGUI")
'    Set application = SapGuiAuto.GetScriptingEngine
' End If
' If Not IsObject(connection) Then
'    Set connection = application.Children(0)
' End If
' If Not IsObject(session) Then
'    Set session    = connection.Children(0)
' End If
' If IsObject(WScript) Then
'    WScript.ConnectObject session,     "on"
'    WScript.ConnectObject application, "on"
' End If
' session.findById("wnd[0]").maximize
' session.findById("wnd[0]/tbar[0]/okcd").text = "/nse09"
' session.findById("wnd[0]").sendVKey 0
' session.findById("wnd[0]/usr/subCOMMONSUBSCREEN:RDDM0001:0220/btn%_AUTOTEXT028").press
' session.findById("wnd[0]/usr").horizontalScrollbar.position = 1
' session.findById("wnd[0]/usr").horizontalScrollbar.position = 2
' session.findById("wnd[0]/usr").horizontalScrollbar.position = 1
' session.findById("wnd[0]/usr").horizontalScrollbar.position = 0
' session.findById("wnd[0]/tbar[0]/okcd").text = "/nse16n"
' session.findById("wnd[0]").sendVKey 0
' session.findById("wnd[0]/usr/ctxtGD-TAB").text = "e070"
' session.findById("wnd[0]/usr/ctxtGD-TAB").caretPosition = 4
' session.findById("wnd[0]").sendVKey 0
' session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = ""
' session.findById("wnd[0]/usr/txtGD-MAX_LINES").setFocus
' session.findById("wnd[0]/usr/txtGD-MAX_LINES").caretPosition = 0
' session.findById("wnd[0]/tbar[1]/btn[7]").press
' session.findById("wnd[1]").close
' session.findById("wnd[0]/tbar[1]/btn[8]").press
' session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").firstVisibleRow = 35
' session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").firstVisibleRow = 63
' session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").setCurrentCell -1,""
' session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").firstVisibleRow = 0
' session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectAll
' session.findById("wnd[1]").close

' Dim ordem1, ordem2, ordem3
' ordem1 = shell.getCellValue(0, "Ordem/tarefa")
' ordem2 = shell.getCellValue(1, "Ordem/tarefa")
' ordem3 = shell.getCellValue(2, "Ordem/tarefa")

' WScript.Echo "Ordens marcadas: " & ordem1 & ", " & ordem2 & ", " & ordem3
Option Explicit

Dim SapGuiAuto
Dim application
Dim connection
Dim session
Dim shell

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

'========================================================
' Vai para a SE16N
'========================================================
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse16n"
session.findById("wnd[0]").sendVKey 0

WScript.Sleep 800

'========================================================
' Informa a tabela E070
'========================================================
If WndExists(session, "wnd[0]/usr/ctxtGD-TAB") Then
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "E070"
    session.findById("wnd[0]").sendVKey 0
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
If WndExists(session, "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SE16N_SELFIELDS-LOW[1,0]") Then
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SE16N_SELFIELDS-LOW[1,0]").text = "DDR*"
Else
    StopScript "Campo LOW da linha TRKORR não encontrado."
End If

If WndExists(session, "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SE16N_SELFIELDS-LOW[1,2]") Then
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SE16N_SELFIELDS-LOW[1,2]").text = "D"
Else
    StopScript "Campo LOW da linha TRSTATUS não encontrado."
End If

WScript.Sleep 500

'========================================================
' Executa
'========================================================
If WndExists(session, "wnd[0]/tbar[1]/btn[8]") Then
    session.findById("wnd[0]/tbar[1]/btn[8]").press
Else
    StopScript "Botão Executar não encontrado."
End If

WScript.Sleep 1200

'========================================================
' Obtém ALV de resultado
'========================================================
If WndExists(session, "wnd[0]/usr/cntlRESULT_LIST/shellcont/shell") Then
    Set shell = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
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

WScript.Echo "Ordens encontradas: " & ordem1 & ", " & ordem2 & ", " & ordem3