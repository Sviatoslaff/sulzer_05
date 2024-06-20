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

Set grid = session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell")

strTemplate = "BUY-2001"
qtyRows = grid.rowCount - 1
MsgBox "Rows amount: " & qtyRows
visibleRows = grid.VisibleRowCount
MsgBox "Visible Rows amount: " & qtyRows

' Цикл для каждой строки
'On Error Resume Next
Do Until intRow > qtyRows
    'Err.Clear
'	MsgBox "Row: " & intRow
	grid.modifyCell intRow, "TEMPLATE", strTemplate
 	grid.currentCellRow = intRow 
'	sapRow = grid.currentRow
    intRow = intRow + 1
Loop
grid.triggerModified


MsgBox "Finished!", vbSystemModal Or vbInformation


