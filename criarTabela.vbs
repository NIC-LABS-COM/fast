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
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse11"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRSRD1-TBMA_VAL").text = "ztab_test"
session.findById("wnd[0]/usr/ctxtRSRD1-TBMA_VAL").caretPosition = 9
session.findById("wnd[0]/usr/btnPUSHADD").press
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLSD41:2202/ctxtDD02D-CONTFLAG").text = "A"
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLSD41:2202/ctxtDD02D-CONTFLAG").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLSD41:2202/ctxtDD02D-CONTFLAG").caretPosition = 1
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/txtDD02D-DDTEXT").text = "TABELA TESTE"
session.findById("wnd[0]/usr/txtDD02D-DDTEXT").caretPosition = 12
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF").select
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/txtDD03D-FIELDNAME[0,0]").text = "MANDT"
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/txtDD03D-FIELDNAME[0,0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/txtDD03D-FIELDNAME[0,0]").caretPosition = 5
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/chkDD03P-KEYFLAG[1,0]").selected = true
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/chkDD03P-NOTNULL[2,0]").selected = true
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/ctxtDD03D-ROLLNAME[3,0]").text = "MANDT"
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/ctxtDD03D-ROLLNAME[3,0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF/ssubTS_SCREEN:SAPLSD41:2201/tblSAPLSD41TC0/ctxtDD03D-ROLLNAME[3,0]").caretPosition = 5
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN").select
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLSD41:2202/cmbDD02D-MAINFLAG").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLSD41:2202/cmbDD02D-MAINFLAG").key = "X"
