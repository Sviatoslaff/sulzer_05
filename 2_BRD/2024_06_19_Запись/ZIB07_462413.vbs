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
session.findById("wnd[0]/usr/ctxtP_EQUNR").text = "462413"
session.findById("wnd[0]/usr/ctxtP_WERKS2").text = "2001"
session.findById("wnd[0]/usr/ctxtP_WERKS2").setFocus
session.findById("wnd[0]/usr/ctxtP_WERKS2").caretPosition = 4
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
