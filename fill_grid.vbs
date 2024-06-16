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

'session.findById("wnd[0]").maximize

intRow = 0
qtyRows = session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell").rowCount
MsgBox "Rows amount: " & qtyRows
' Цикл для каждой строки
'On Error Resume Next
Do Until intRow > qtyRows
    'Err.Clear
    Set grid = session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell")
	MsgBox "Row: " & intRow
	grid.modifyCell intRow, "TEMPLATE", "BUY-2001"
 	grid.currentCellRow = intRow 
'	sapRow = grid.currentRow               
	grid.triggerModified
    intRow = intRow + 1
Loop

MsgBox "Finished!", vbSystemModal Or vbInformation


