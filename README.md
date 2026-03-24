Dim fso, resultado
Set fso = CreateObject("Scripting.FileSystemObject")
resultado = ""

' Verifica os componentes COM do SAP
Dim componentes
componentes = Array("SAP.Functions", "SAP.LogonControl", _
                    "SAP.LogonControl.1", "SAP.TableFactory", _
                    "SAP.Rfc.Connection", "SAPGUI")

Dim comp
For Each comp In componentes
    On Error Resume Next
    Dim obj
    Set obj = CreateObject(comp)
    If Err.Number = 0 Then
        resultado = resultado & comp & " = DISPONIVEL" & vbCrLf
        Set obj = Nothing
    Else
        resultado = resultado & comp & " = NAO DISPONIVEL" & vbCrLf
    End If
    Err.Clear
    On Error GoTo 0
Next

' Verifica se DLLs existem
Dim pastas, pasta
pastas = Array( _
    "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\", _
    "C:\Program Files\SAP\FrontEnd\SAPgui\", _
    "C:\Windows\System32\", _
    "C:\Windows\SysWOW64\" _
)

Dim dlls
dlls = Array("wdtfuncs.ocx", "librfc32.dll", "sapnco.dll", "wdtlog.ocx")

resultado = resultado & vbCrLf & "=== DLLs ===" & vbCrLf

Dim p, d
For Each p In pastas
    For Each d In dlls
        If fso.FileExists(p & d) Then
            resultado = resultado & "ENCONTRADO: " & p & d & vbCrLf
        End If
    Next
Next

MsgBox resultado, vbInformation, "Diagnostico SAP Components"

' Salva em arquivo tambem
Dim WshShell
Set WshShell = CreateObject("WScript.Shell")
Dim arq
Set arq = fso.CreateTextFile(WshShell.SpecialFolders("Desktop") & "\diagnostico_sap.txt", True)
arq.Write resultado
arq.Close

MsgBox "Resultado salvo em diagnostico_sap.txt na Area de Trabalho", vbInformation
