' ============================================================
' buscarProgSap.vbs
' Le o codigo fonte de um programa ABAP existente via SE38
' e retorna via WScript.Echo (stdout) com \n como separador
'
' Programa fixo:
'   Z05_TESTE_1
' ============================================================
Option Explicit

Dim SapGuiAuto, application, connection, session
If Not IsObject(application) Then
    Set SapGuiAuto  = GetObject("SAPGUI")
    Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
    Set connection = application.Children(0)
End If
If Not IsObject(session) Then
    Set session = connection.Children(0)
End If

Dim programName
programName = "Z05_TESTE_1"

' ---- Navegar para SE38 e abrir o editor do programa -------------
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse38"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 800

session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").text = programName
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").caretPosition = Len(programName)
WScript.Sleep 200
session.findById("wnd[0]").sendVKey 0  ' Enter abre o editor
WScript.Sleep 1000

' ---- Tentativa 1: currentContent --------------------------------
Dim sourceCode
sourceCode = ""
On Error Resume Next
sourceCode = session.findById("wnd[0]/usr/cntlEDITOR/shellcont/shell").currentContent
On Error GoTo 0

' ---- Tentativa 2: export via menu (Download) --------------------
If Trim(sourceCode) = "" Then
    Dim fso, tempPath, fRead
    Set fso  = CreateObject("Scripting.FileSystemObject")
    tempPath = fso.GetSpecialFolder(2) & "\sap_export_" & programName & ".txt"

    On Error Resume Next
    session.findById("wnd[0]/mbar/menu[3]/menu[7]/menu[1]/menu[1]").Select
    WScript.Sleep 800
    session.findById("wnd[1]/usr/ctxtDY_PATH").text     = fso.GetSpecialFolder(2) & "\"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "sap_export_" & programName & ".txt"
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    On Error GoTo 0
    WScript.Sleep 800

    If fso.FileExists(tempPath) Then
        Set fRead  = fso.OpenTextFile(tempPath, 1, False)
        sourceCode = fRead.ReadAll
        fRead.Close
        On Error Resume Next
        fso.DeleteFile tempPath
        On Error GoTo 0
    End If
End If

If Trim(sourceCode) = "" Then
    WScript.StdErr.Write "not found: Programa " & programName & " nao encontrado ou sem codigo."
    WScript.Quit 1
End If

' ---- Codificar quebras de linha para transporte -----------------
sourceCode = Replace(sourceCode, vbCrLf, "\n")
sourceCode = Replace(sourceCode, vbCr,   "\n")
sourceCode = Replace(sourceCode, vbLf,   "\n")

WScript.Echo sourceCode
WScript.Quit 0
