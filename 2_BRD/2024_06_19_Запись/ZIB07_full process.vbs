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
session.findById("wnd[0]/tbar[0]/okcd").text = "ZIB07"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtP_EQUNR").text = "522443"
session.findById("wnd[0]/usr/ctxtP_WERKS2").text = "2001"
session.findById("wnd[0]/usr/ctxtP_WERKS2").setFocus
session.findById("wnd[0]/usr/ctxtP_WERKS2").caretPosition = 4
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/chkJOB").selected = false
session.findById("wnd[0]/usr/chkJOB").setFocus
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 0,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").currentCellRow = 1
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 1,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 2,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").currentCellRow = 2
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 3,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").currentCellRow = 3
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 4,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").currentCellRow = 4
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 5,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").currentCellRow = 5
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 6,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").currentCellRow = 6
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 7,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").currentCellRow = 7
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 8,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").currentCellRow = 8
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 9,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").currentCellRow = 9
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 10,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").currentCellRow = 10
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 11,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").currentCellRow = 11
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 12,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").currentCellRow = 12
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 13,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").currentCellRow = 13
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 14,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").currentCellRow = 14
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 15,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").currentCellRow = 15
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 16,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").currentCellRow = 16
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 17,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").currentCellRow = 17
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 18,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").currentCellRow = 18
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 19,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").currentCellRow = 19
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 19,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 20,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").currentCellRow = 20
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
