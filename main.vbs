Option Explicit
Public Const serRow = 9

Dim qtn, plant, sorg, template, serno, grid
Dim qtyRows, visibleRows, intRow, grid
'Запрашиваем файл QTN
Dim excelFile
excelFile = selectExcel()
Dim arrSerno : arrSerno = GetUniqSerNumbersArray()
'WScript.Echo Join(arrSerno)

'StartTransaction("ZIB07")
session.findById("wnd[0]/tbar[0]/okcd").text = "ZIB07"
session.findById("wnd[0]").sendVKey 0

For Each serno In arrSerno
  
  session.findById("wnd[0]/usr/ctxtP_EQUNR").text = serno
  session.findById("wnd[0]/usr/ctxtP_WERKS2").text = plant
  session.findById("wnd[0]/tbar[1]/btn[8]").press
  WScript.Sleep 500     'Delay for SAP processing
  If session.findById("wnd[0]/usr/ctxtP_EQUNR",False) Is Nothing Then
    If session.findById("wnd[0]/usr/chkJOB",False) Is Nothing Then
      session.findById("wnd[1]/tbar[0]/btn[8]").press
    End If
  
    Set grid = session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell")

    qtyRows = grid.rowCount - 1
    'MsgBox "Rows amount: " & qtyRows
    visibleRows = grid.VisibleRowCount
    'MsgBox "Visible Rows amount: " & qtyRows
    
    ' Цикл для каждой строки
    'On Error Resume Next
    intRow = 0
    Do Until intRow > qtyRows
        'Err.Clear
    '	MsgBox "Row: " & intRow
        grid.modifyCell intRow, "TEMPLATE", template
        grid.currentCellRow = intRow 
    '	sapRow = grid.currentRow
        intRow = intRow + 1
    Loop
    grid.triggerModified  
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press

  Else
    ' Same selection window - doing nothing
  End If

  'Exit For
Next

MsgBox "The script finished!", vbSystemModal Or vbInformation

'returns an unique array of serial numbers from Excel file chosen by user
Function GetUniqSerNumbersArray()
    'excelFile = "C:\VBScript\articles.xlsx" ' Полный путь к выбранному файлу
    Dim ArticlesExcel, objWorkbook
    Set ArticlesExcel = CreateObject("Excel.Application")
    Set objWorkbook = ArticlesExcel.Workbooks.Open (excelFile)
    qtn   = ArticlesExcel.Cells(22, 4).Value
    plant = ArticlesExcel.Cells(21, 4).Value
    sorg  = ArticlesExcel.Cells(3, 4).Value
    template = "BUY-" & sorg
    
    Dim arrSerno()

    ' Считаем, что в 25 строке - начало таблицы для обработки
    Dim intRow : intRow = 25
    ' Цикл для каждой строки
    On Error Resume Next
    Do Until ArticlesExcel.Cells(intRow, serRow).Value = ""
      ReDim Preserve arrSerno(intRow - 25)
      'WScript.Echo ArticlesExcel.Cells(intRow, serRow).Value
      arrSerno(intRow - 25) = ArticlesExcel.Cells(intRow, serRow).Value
      intRow = intRow + 1
    Loop
    objWorkbook.Close False
    ArticlesExcel.Quit
    'WScript.Echo Join(arrSerno)
    Dim arrUniqSerno : arrUniqSerno = uniqFE(arrSerno)
    'WScript.Echo Join(arrUniqSerno)

    GetUniqSerNumbersArray = arrUniqSerno

End Function


' returns an array of the unique items in for-each-able collection fex
Function uniqFE(fex)
  Dim dicTemp : Set dicTemp = CreateObject("Scripting.Dictionary")
  Dim xItem
  For Each xItem In fex
      dicTemp(xItem) = 0
  Next
  uniqFE = dicTemp.Keys()
End Function