Option Explicit
Public Const serRow = 9

Dim qtn, plant, sorg, template

Dim arrSerno : GetUniqSerNumbersArray()
WScript.Echo Join(arrSerno)



MsgBox "The script finished!", vbSystemModal Or vbInformation

'returns an unique array of serial numbers from Excel file chosen by user
Function GetUniqSerNumbersArray()
    'Запрашиваем файл QTN
    Dim excelFile
    excelFile = selectExcel()

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
    Dim intRow = 25
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