If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

'--- Funcao para esperar um elemento carregar (ate 10 segundos) ---
Function EsperarElemento(strId)
    Dim obj, t
    Set obj = Nothing
    For t = 1 To 10
        On Error Resume Next
        Set obj = session.findById(strId)
        On Error GoTo 0
        If Not obj Is Nothing Then
            Set EsperarElemento = obj
            Exit Function
        End If
        WScript.Sleep 1000
    Next
    MsgBox "Erro: Elemento nao encontrado apos 10s:" & vbCrLf & strId, vbCritical
    WScript.Quit
End Function

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse11"
session.findById("wnd[0]").sendVKey 0

'--- Espera a tela da SE11 carregar ---
Dim campoTbma
Set campoTbma = EsperarElemento("wnd[0]/usr/ctxtRSRD1-TBMA_VAL")
campoTbma.text = "ztab_test"
campoTbma.caretPosition = 9

session.findById("wnd[0]/usr/btnPUSHADD").press

'--- Espera a tela de manutencao carregar apos o press ---
Dim campoContFlag
Set campoContFlag = EsperarElemento("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLSD41:2202/ctxtDD02D-CONTFLAG")
campoContFlag.text = "A"
campoContFlag.setFocus
campoContFlag.caretPosition = 1

session.findById("wnd[0]").sendVKey 0
WScript.Sleep 500

session.findById("wnd[0]/usr/txtDD02D-DDTEXT").text = "TABELA TESTE"
session.findById("wnd[0]/usr/txtDD02D-DDTEXT").caretPosition = 12

'--- Troca para aba DEF e espera carregar ---
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF").select
WScript.Sleep 1000

Dim campoField
Set campoField = EsperarElemento("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/txtDD03D-FIELDNAME[0,0]")
campoField.text = "MANDT"
campoField.setFocus
campoField.caretPosition = 5

session.findById("wnd[0]").sendVKey 0
WScript.Sleep 500

session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/chkDD03P-KEYFLAG[1,0]").selected = true
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/chkDD03P-NOTNULL[2,0]").selected = true

Dim campoRoll
Set campoRoll = EsperarElemento("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/ctxtDD03D-ROLLNAME[3,0]")
campoRoll.text = "MANDT"
campoRoll.setFocus
campoRoll.caretPosition = 5

session.findById("wnd[0]").sendVKey 0
WScript.Sleep 500

'--- Volta para aba MAIN e espera carregar ---
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN").select
WScript.Sleep 1000

Dim campoMainFlag
Set campoMainFlag = EsperarElemento("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLSD41:2202/cmbDD02D-MAINFLAG")
campoMainFlag.setFocus
campoMainFlag.key = "X"
