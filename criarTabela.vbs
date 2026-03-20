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

Dim tableName
Dim tableText
Dim packageName
Dim requestId
Dim tabArt
Dim tabKat
Dim fieldsSpec
Dim deliveryClass

tableName = "ZMM_FORNECED10"
tableText = "Tabela de Fornecedor"
packageName = "z_php"
requestId = "A4HK904843"
tabArt = "APPL0"
tabKat = "3"
fieldsSpec = "MANDT|1|1|MANDT;CNPJ|1|1|Z_MM_CNPJ"
deliveryClass = "A"

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse11"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRSRD1-TBMA_VAL").text = tableName
session.findById("wnd[0]/usr/ctxtRSRD1-TBMA_VAL").caretPosition = 13
session.findById("wnd[0]/usr/btnPUSHADD").press

session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpHEAD").select
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabDEF").select
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpF4V").select
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpREFF").select
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLDS41:2202/ctxtDD02D-CONTFLAG").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLDS41:2202/ctxtDD02D-CONTFLAG").caretPosition = 0

session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/lbl[3,3]").setFocus
session.findById("wnd[1]/usr/lbl[3,3]").caretPosition = 4
session.findById("wnd[0]").sendVKey 2

session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLDS41:2202/cmbDD02D-CONTFLAG").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLDS41:2202/cmbDD02D-CONTFLAG").key = "X"
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF").select

session.findById("wnd[0]/usr/textDD02D-DDTEXT").text = tableText
session.findById("wnd[0]/usr/textDD02D-DDTEXT").caretPosition = 12

session.findById("wnd[0]/usr/tabsTAB_STRIP/tabDEF").select
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpF4V").select
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpREFF").select
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpHEAD").select



