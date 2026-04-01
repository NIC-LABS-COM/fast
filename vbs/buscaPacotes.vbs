Option Explicit

Dim SapGuiAuto
Dim application
Dim connection
Dim session
Dim i
Dim fso
Dim filePath

Const REPORT_NAME = "Z_GET_ALL_PACKAGES"
Const FILE_PATH   = "C:\temp\packages_tdevc.txt"

Sub WaitSeconds(seconds)
    Dim startTime
    startTime = Timer

    Do While Timer < startTime + seconds
    Loop
End Sub

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

' Vai para SE38
session.findById("wnd[0]/tbar[0]/okcd").Text = "/NSE38"
session.findById("wnd[0]").sendVKey 0
Call WaitSeconds(1)

' Informa o report
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = REPORT_NAME
session.findById("wnd[0]").sendVKey 8
Call WaitSeconds(3)

' Se aparecer popup, tenta fechar
On Error Resume Next
session.findById("wnd[1]/tbar[0]/btn[0]").press
Err.Clear
On Error GoTo 0

Call WaitSeconds(1)

' Verifica se o arquivo foi gerado
Set fso = CreateObject("Scripting.FileSystemObject")
filePath = FILE_PATH

For i = 1 To 10
    If fso.FileExists(filePath) Then
        Exit For
    End If
    Call WaitSeconds(1)
Next

If fso.FileExists(filePath) Then
    MsgBox "Arquivo gerado com sucesso: " & filePath
Else
    MsgBox "Arquivo nao encontrado: " & filePath
End If
