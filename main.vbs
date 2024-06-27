Option Explicit
Public Const serRow = 9, resOK = "Added successfully", resExists = "Already exists in SAP" 
Public Const resNoTemplate = " template not found. Check the template.  "
Public Const resNoBOM = "Nothing is inside this BOM. First make the BOM."

Dim qtn, plant, sorg, template, serno
Dim qtyRows, visibleRows, intRow, grid, bExit, bAbort
'Запрашиваем файл QTN
Dim excelFile
excelFile = selectExcel()
Dim arrSerno : arrSerno = GetUniqSerNumbersArray()

' Make a structure for a report
Dim qtySerno : qtySerno= UBound(arrSerno)
Dim arrReport, strReport
Dim dicReport : Set dicReport  = CreateObject("Scripting.Dictionary")

'WScript.Echo Join(arrSerno)

'StartTransaction("ZIB07")
session.findById("wnd[0]/tbar[0]/okcd").text = "ZIB07"
session.findById("wnd[0]").sendVKey 0

bAbort = vbFalse
For Each serno In arrSerno
  bExit = vbFalse
  session.findById("wnd[0]/usr/ctxtP_EQUNR").text  = serno
  session.findById("wnd[0]/usr/ctxtP_WERKS2").text = plant
  session.findById("wnd[0]/tbar[1]/btn[8]").press
  WScript.Sleep 500     'Delay for SAP processing
  If session.findById("wnd[0]/usr/ctxtP_EQUNR",False) Is Nothing Then
    Do While session.findById("wnd[0]/usr/chkJOB", False) Is Nothing
      If session.findById("wnd[1]/usr/txtLV_MATNR1", False) Is Nothing Then     
        dicReport.Add serno, resNoBOM
        bExit = vbTrue
        session.findById("wnd[1]").sendVKey 0
        Exit Do
      Else
        session.findById("wnd[1]/tbar[0]/btn[8]").press       'V
        'session.findById("wnd[1]/tbar[0]/btn[2]").press       'X        
      End If 
    Loop

    If Not bExit Then
      session.findById("wnd[0]/usr/chkJOB").selected = false
      session.findById("wnd[0]/usr/chkJOB").setFocus  

      Set grid = session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell")

      qtyRows = grid.rowCount - 1
      'MsgBox "Rows amount: " & qtyRows
      visibleRows = grid.VisibleRowCount
      
      ' Цикл для каждой строки
      'On Error Resume Next
      intRow = 0
      Do Until intRow > qtyRows
          'Err.Clear
          'MsgBox "Row: " & intRow
          grid.modifyCell intRow, "TEMPLATE", template
          grid.currentCellRow = intRow 
          intRow = intRow + 1
      Loop
      grid.triggerModified  
      session.findById("wnd[0]/tbar[1]/btn[8]").press
      '    MsgBox "Next Control - btn[3]", vbSystemModal Or vbInformation
  
      ' It can be error that mat number not found - If for that
      If session.findById("wnd[1]/tbar[0]/btn[0]", False) Is Nothing Then
      Else
        bAbort = vbTrue
        dicReport.Add serno, template & resNoTemplate
        session.findById("wnd[1]/tbar[0]/btn[0]").press          
      End If 

      If Not bAbort Then
        session.findById("wnd[0]/tbar[0]/btn[3]").press
        '    MsgBox "Next Control - wnd[1]/tbar[0]/btn[0]", vbSystemModal Or vbInformation
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        dicReport.Add serno, resOK        
      End If  
    End If  
  Else
    ' Same selection window - doing nothing
    dicReport.Add serno, resExists
  End If
Next

OutputToExcel
MsgBox "Script finished! ", vbSystemModal Or vbInformation

'====== Functions ans Subs ========

'returns an unique array of serial numbers from Excel file chosen by user
Function GetUniqSerNumbersArray()
    'excelFile = "C:\VBScript\articles.xlsx" ' Полный путь к выбранному файлу
    Dim ArticlesExcel, objWorkbook
    Set ArticlesExcel = CreateObject("Excel.Application")
    Set objWorkbook = ArticlesExcel.Workbooks.Open (excelFile)
    qtn   = ArticlesExcel.Cells(22, 4).Value
    plant = ArticlesExcel.Cells(21, 4).Value
    sorg  = ArticlesExcel.Cells(3, 4).Value
    template = "BUY-" & plant
    
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

Sub OutputToExcel 
  Dim ReportExcel, objWorkbook
  Set ReportExcel = CreateObject("Excel.Application")
  Set objWorkbook = ReportExcel.Workbooks.Add()
  ReportExcel.Visible = True

  arrReport = dicReport.Items
  intRow = 0
  For Each serno In arrSerno
    strReport = strReport & serno & " : " & arrReport(intRow) & vbCrLf
    objWorkbook.cells(intRow,1).value = serno
    objWorkbook.cells(intRow,2).value = arrReport(intRow) 
    intRow = intRow + 1
  Next
End Sub