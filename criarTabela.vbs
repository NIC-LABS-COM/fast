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

' Definir as variáveis
Dim tableName, tableText, packageName, requestId, tabArt, tabKat, fieldsSpec, deliveryClass
tableName = "ZMM_FORNECED10"
tableText = "Tabela de Fornecedor"
packageName = "z_php"
requestId = "A4HK904843"
tabArt = "APPL0"
tabKat = "3"
fieldsSpec = "MANDT|1|1|MANDT;CNPJ|1|1|Z_MM_CNPJ"
deliveryClass = "A"

' Maximizar janela
session.findById("wnd[0]").maximize
WScript.Sleep 500 ' Espera 500ms para garantir que a janela foi maximizada

' Enviar comando para SAP
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse11"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 500 ' Espera para garantir que o SAP processou o comando

' Preencher campos
session.findById("wnd[0]/usr/ctxtRSRD1-TBMA_VAL").text = tableName
session.findById("wnd[0]/usr/ctxtRSRD1-TBMA_VAL").caretPosition = 13
session.findById("wnd[0]/usr/btnPUSHADD").press
WScript.Sleep 500 ' Espera para garantir que o botão foi pressionado

' Selecionar abas
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpHEAD").select
WScript.Sleep 500 ' Espera para garantir que a aba foi selecionada
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabDEF").select
WScript.Sleep 500 ' Espera para garantir que a aba foi selecionada
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpF4V").select
WScript.Sleep 500 ' Espera para garantir que a aba foi selecionada
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpREFF").select
WScript.Sleep 500 ' Espera para garantir que a aba foi selecionada

' Focar no campo e setar a posição do cursor
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLDS41:2202/ctxtDD02D-CONTFLAG").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLDS41:2202/ctxtDD02D-CONTFLAG").caretPosition = 0
WScript.Sleep 500 ' Espera para garantir que o campo foi focado

' Enviar comando para SAP
session.findById("wnd[0]").sendVKey 4
WScript.Sleep 500 ' Espera para garantir que o SAP processou o comando

' Focar no próximo campo
session.findById("wnd[1]/usr/lbl[3,3]").setFocus
session.findById("wnd[1]/usr/lbl[3,3]").caretPosition = 4
WScript.Sleep 500 ' Espera para garantir que o campo foi focado

' Enviar outro comando VKey
session.findById("wnd[0]").sendVKey 2
WScript.Sleep 500 ' Espera para garantir que o SAP processou o comando

' Focar e interagir com o campo CONTFLAG
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLDS41:2202/cmbDD02D-CONTFLAG").setFocus
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpMAIN/ssubTS_SCREEN:SAPLDS41:2202/cmbDD02D-CONTFLAG").key = "X"
WScript.Sleep 500 ' Espera para garantir que o campo foi modificado

' Selecionar outra aba
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpDEF").select
WScript.Sleep 500 ' Espera para garantir que a aba foi selecionada

' Preencher o campo de texto
session.findById("wnd[0]/usr/textDD02D-DDTEXT").text = tableText
session.findById("wnd[0]/usr/textDD02D-DDTEXT").caretPosition = 12
WScript.Sleep 500 ' Espera para garantir que o campo foi preenchido

' Finalizar com a seleção das abas
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabDEF").select
WScript.Sleep 500 ' Espera para garantir que a aba foi selecionada
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpF4V").select
WScript.Sleep 500 ' Espera para garantir que a aba foi selecionada
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpREFF").select
WScript.Sleep 500 ' Espera para garantir que a aba foi selecionada
session.findById("wnd[0]/usr/tabsTAB_STRIP/tabpHEAD").select
WScript.Sleep 500 ' Espera para garantir que a aba foi selecionada
