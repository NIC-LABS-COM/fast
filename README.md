On Error Resume Next

Dim objRFC
Set objRFC = CreateObject("SAP.Functions")

If Err.Number <> 0 Then
    MsgBox "SAP.Functions NAO disponivel." & vbCrLf & Err.Description, vbCritical
    WScript.Quit
End If

MsgBox "SAP.Functions EXISTE! Componente disponivel.", vbInformation
Set objRFC = Nothing
