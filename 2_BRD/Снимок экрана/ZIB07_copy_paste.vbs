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
session.findById("wnd[0]/usr/chkJOB").selected = false
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 0,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 1,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 2,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 3,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 4,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 5,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 6,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 7,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 8,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 9,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 10,"TEMPLATE","BUY-2001"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").currentCellRow = 7
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").triggerModified
