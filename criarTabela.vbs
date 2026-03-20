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

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse11"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRSRD1-TBMA_VAL").text = "z tabela_teste"
session.findById("wnd[0]/usr/ctxtRSRD1-TBMA_VAL").caretPosition = 13
session.findById("wnd[0]/usr/btnPUSHADD").press

session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpHEAD").select
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabDEF").select
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpF4V").select
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpREFF").select

session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLSD41:2202/ctxtDD02D-CONTFLAG").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLSD41:2202/ctxtDD02D-CONTFLAG").caretPosition = 0

session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/lbl[3,3]").setFocus
session.findById("wnd[1]/usr/lbl[3,3]").caretPosition = 4
session.findById("wnd[0]").sendVKey 2

session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLSD41:2202/cmbDD02D-MAINFLAG").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLSD41:2202/cmbDD02D-MAINFLAG").key = "X"
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabDEF").select

session.findById("wnd[0]/usr/textDD02D-DDTEXT").text = "tabela teste"
session.findById("wnd[0]/usr/textDD02D-DDTEXT").caretPosition = 12

session.findById("wnd[0]/usr/tabsTAB_STRIP/tabDEF").select
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpF4V").select
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpREFF").select
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpHEAD").select
