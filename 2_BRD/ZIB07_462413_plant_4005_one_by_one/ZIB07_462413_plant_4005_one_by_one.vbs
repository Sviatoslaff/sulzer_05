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
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").modifyCell 1,"TEMPLATE","BUY-4005"
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").currentCellRow = 1
session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").pressEnter
session.findById("wnd[1]/tbar[0]/btn[0]").press
